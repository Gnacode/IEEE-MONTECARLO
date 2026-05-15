#!/usr/bin/env python3
"""
backup_montecarlo.py - Backup the Monte Carlo data tree with hash verification.

Default behaviour
-----------------
1. Copies U:/MONTECARLO/data into U:/MONTECARLO/backups/<timestamp>/data/
2. Optionally also copies U:/MONTECARLO/src/  (the simulator source)
3. Computes SHA-256 of every source AND every destination file
4. Verifies every file copied correctly (mismatched hashes -> exit 1)
5. Writes manifest.json with file list, sizes, hashes, and elapsed time

Usage
-----
    # Default - copy U:/MONTECARLO/data to U:/MONTECARLO/backups/<timestamp>/
    python backup_montecarlo.py

    # Customize source and destination
    python backup_montecarlo.py --source U:/MONTECARLO/data --dest D:/backups

    # Make a single .zip archive instead of a copied tree
    python backup_montecarlo.py --zip

    # Skip the src/ folder (data only)
    python backup_montecarlo.py --no-src

    # Quiet output (only errors and summary)
    python backup_montecarlo.py --quiet

Author: GNACODE INC, January 2026
"""

import argparse
import hashlib
import json
import os
import shutil
import sys
import time
import zipfile
from datetime import datetime
from pathlib import Path

# =============================================================================
# Defaults (override with command-line args)
# =============================================================================

DEFAULT_SOURCE = 'U:/MONTECARLO/data'
DEFAULT_SRC_DIR = 'U:/MONTECARLO/src'        # simulator source code
DEFAULT_DEST_ROOT = 'U:/MONTECARLO/backups'  # destination parent directory

# =============================================================================
# Helpers
# =============================================================================

def fmt_bytes(n: int) -> str:
    """Format byte count as human-readable string."""
    if n >= 1 << 30:
        return f"{n / (1 << 30):.2f} GB"
    if n >= 1 << 20:
        return f"{n / (1 << 20):.2f} MB"
    if n >= 1 << 10:
        return f"{n / (1 << 10):.2f} KB"
    return f"{n} B"


def sha256_file(path: Path, chunk: int = 1 << 20) -> str:
    """Compute SHA-256 of a file. 1 MiB chunks - fast for large npz files."""
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        while True:
            data = f.read(chunk)
            if not data:
                break
            h.update(data)
    return h.hexdigest()


def walk_files(root: Path):
    """Yield (relative_path_str, absolute_path) for every file under root."""
    if not root.exists():
        return
    for dirpath, _, filenames in os.walk(root):
        for fn in filenames:
            full = Path(dirpath) / fn
            rel = full.relative_to(root)
            yield rel.as_posix(), full


def dir_size(root: Path) -> int:
    """Sum of file sizes under root."""
    total = 0
    for _, full in walk_files(root):
        total += full.stat().st_size
    return total


# =============================================================================
# Main backup procedure
# =============================================================================

def backup(source: Path, dest_root: Path,
           include_src: bool = True,
           src_dir: Path = None,
           use_zip: bool = False,
           quiet: bool = False) -> int:
    """Run the backup. Returns exit code (0 success, 1 verification failure)."""

    def info(msg):
        if not quiet:
            print(msg)

    t0 = time.time()
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

    # ────────────────────────────────────────────────────────────────────────
    # Pre-flight
    # ────────────────────────────────────────────────────────────────────────
    info('=' * 75)
    info(f'Monte Carlo Data Backup  ({timestamp})')
    info('=' * 75)

    if not source.exists():
        print(f"ERROR: source directory not found: {source}", file=sys.stderr)
        return 1

    src_size = dir_size(source)
    info(f"  Source:  {source}  ({fmt_bytes(src_size)})")

    if include_src and src_dir and src_dir.exists():
        src_code_size = dir_size(src_dir)
        info(f"  Src/:    {src_dir}  ({fmt_bytes(src_code_size)})")
    else:
        src_code_size = 0
        if include_src and src_dir:
            info(f"  Src/:    {src_dir}  (not found - skipping)")
        include_src = False

    # ────────────────────────────────────────────────────────────────────────
    # Build the source manifest (hash every source file BEFORE copying).
    # This protects against silent disk corruption: if the source file
    # changes between hashing and reading, we'd see it.
    # ────────────────────────────────────────────────────────────────────────
    info('  Hashing source files...')

    src_manifest = {}  # rel_path -> {'size', 'sha256', 'group'}

    for rel, full in walk_files(source):
        size = full.stat().st_size
        digest = sha256_file(full)
        src_manifest[f"data/{rel}"] = {
            'size': size,
            'sha256': digest,
            'group': 'data',
        }

    if include_src:
        for rel, full in walk_files(src_dir):
            size = full.stat().st_size
            digest = sha256_file(full)
            src_manifest[f"src/{rel}"] = {
                'size': size,
                'sha256': digest,
                'group': 'src',
            }

    n_files = len(src_manifest)
    total_size = src_size + src_code_size
    info(f"  {n_files} files, {fmt_bytes(total_size)} to back up")

    # ────────────────────────────────────────────────────────────────────────
    # Create destination directory
    # ────────────────────────────────────────────────────────────────────────
    dest_root.mkdir(parents=True, exist_ok=True)
    backup_dir = dest_root / timestamp
    backup_dir.mkdir(parents=True, exist_ok=True)
    info(f"  Dest:    {backup_dir}")

    # ────────────────────────────────────────────────────────────────────────
    # Copy (or zip) the files
    # ────────────────────────────────────────────────────────────────────────
    info('  Copying files...')
    t_copy_start = time.time()

    if use_zip:
        zip_path = backup_dir / f"montecarlo_backup_{timestamp}.zip"
        with zipfile.ZipFile(zip_path, 'w',
                             compression=zipfile.ZIP_DEFLATED,
                             compresslevel=6) as zf:
            for rel, full in walk_files(source):
                zf.write(full, arcname=f"data/{rel}")
            if include_src:
                for rel, full in walk_files(src_dir):
                    zf.write(full, arcname=f"src/{rel}")
        info(f"  Zip:     {zip_path}  ({fmt_bytes(zip_path.stat().st_size)})")

        # Verification of zip: extract to memory and hash
        info('  Verifying zip contents (re-hashing from archive)...')
        verify_failures = []
        with zipfile.ZipFile(zip_path, 'r') as zf:
            for archive_name in zf.namelist():
                if archive_name not in src_manifest:
                    continue
                expected = src_manifest[archive_name]['sha256']
                with zf.open(archive_name) as fh:
                    h = hashlib.sha256()
                    while True:
                        data = fh.read(1 << 20)
                        if not data:
                            break
                        h.update(data)
                    actual = h.hexdigest()
                if actual != expected:
                    verify_failures.append({
                        'file': archive_name,
                        'expected': expected,
                        'actual': actual,
                    })
    else:
        # Tree copy
        dest_data = backup_dir / 'data'
        dest_data.mkdir(exist_ok=True)
        for rel, full in walk_files(source):
            target = dest_data / rel
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(full, target)

        if include_src:
            dest_src = backup_dir / 'src'
            dest_src.mkdir(exist_ok=True)
            for rel, full in walk_files(src_dir):
                target = dest_src / rel
                target.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(full, target)

        # Verification of tree: re-hash every destination file
        info('  Verifying destination files (re-hashing from disk)...')
        verify_failures = []
        for archive_name, info_dict in src_manifest.items():
            target = backup_dir / archive_name
            if not target.exists():
                verify_failures.append({
                    'file': archive_name,
                    'expected': info_dict['sha256'],
                    'actual': '<missing>',
                })
                continue
            expected = info_dict['sha256']
            actual = sha256_file(target)
            if actual != expected:
                verify_failures.append({
                    'file': archive_name,
                    'expected': expected,
                    'actual': actual,
                })

    t_copy_end = time.time()
    copy_seconds = t_copy_end - t_copy_start
    if copy_seconds > 0 and total_size > 0:
        rate_mb_s = (total_size / (1 << 20)) / copy_seconds
        info(f"  Copied in {copy_seconds:.1f}s ({rate_mb_s:.1f} MB/s)")

    # ────────────────────────────────────────────────────────────────────────
    # Manifest
    # ────────────────────────────────────────────────────────────────────────
    manifest = {
        'backup_timestamp': timestamp,
        'backup_started': datetime.fromtimestamp(t0).isoformat(),
        'backup_seconds': time.time() - t0,
        'source': {
            'data': str(source),
            'src': str(src_dir) if include_src else None,
        },
        'destination': str(backup_dir),
        'archive_format': 'zip' if use_zip else 'tree',
        'total_files': n_files,
        'total_bytes': total_size,
        'verification': {
            'method': 'sha256',
            'failures': len(verify_failures),
            'failed_files': verify_failures,
        },
        'files': src_manifest,
    }

    manifest_path = backup_dir / 'manifest.json'
    with open(manifest_path, 'w') as f:
        json.dump(manifest, f, indent=2)
    info(f"  Manifest: {manifest_path}")

    # ────────────────────────────────────────────────────────────────────────
    # Summary
    # ────────────────────────────────────────────────────────────────────────
    info('-' * 75)
    if verify_failures:
        print(f"  VERIFICATION FAILED: {len(verify_failures)} file(s) "
              f"with hash mismatches", file=sys.stderr)
        for vf in verify_failures[:10]:
            print(f"    - {vf['file']}", file=sys.stderr)
            print(f"      expected: {vf['expected']}", file=sys.stderr)
            print(f"      actual:   {vf['actual']}", file=sys.stderr)
        if len(verify_failures) > 10:
            print(f"    ... and {len(verify_failures) - 10} more "
                  "(see manifest.json)", file=sys.stderr)
        return 1

    info(f"  SUCCESS: {n_files} files, {fmt_bytes(total_size)}, "
         f"all hashes verified")
    info(f"  Total time: {time.time() - t0:.1f}s")
    info('=' * 75)
    return 0


# =============================================================================
# CLI
# =============================================================================

def main():
    parser = argparse.ArgumentParser(
        description='Backup Monte Carlo data with hash verification',
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument('--source', default=DEFAULT_SOURCE,
                        help='Source directory to back up')
    parser.add_argument('--src-dir', default=DEFAULT_SRC_DIR,
                        help='Optional source-code directory to also back up')
    parser.add_argument('--dest', default=DEFAULT_DEST_ROOT,
                        help='Backup destination root '
                             '(timestamped subdir created inside)')
    parser.add_argument('--no-src', action='store_true',
                        help='Skip backup of the src/ directory')
    parser.add_argument('--zip', action='store_true', dest='use_zip',
                        help='Create a single .zip archive instead of '
                             'a tree copy')
    parser.add_argument('--quiet', action='store_true',
                        help='Reduce output (errors and summary only)')

    args = parser.parse_args()

    # Resolve to Path objects (absolute paths)
    source = Path(args.source).resolve()
    dest_root = Path(args.dest).resolve()
    src_dir = Path(args.src_dir).resolve() if args.src_dir else None

    rc = backup(
        source=source,
        dest_root=dest_root,
        include_src=not args.no_src,
        src_dir=src_dir,
        use_zip=args.use_zip,
        quiet=args.quiet,
    )
    sys.exit(rc)


if __name__ == '__main__':
    main()