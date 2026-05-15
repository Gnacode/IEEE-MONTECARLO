#!/usr/bin/env python3
"""
Verification harness for classical comparators (Slots 1-4).
============================================================

Generates synthetic 1D time-series frames matching TSNFA's regime
(100 Hz sample rate, 128 samples per frame), runs each comparator,
and measures detection/false-alarm performance.

Three test scenarios:
  1. Noise-only stream (measure false-alarm rate, sanity-check that
     each detector trains/calibrates without spurious triggers)
  2. Event-bearing stream (measure detection rate at fixed SNR)
  3. SNR sweep (build crude detection-vs-SNR curves)

This is sanity verification, NOT the full Monte Carlo. The goal is to
confirm each algorithm's basic behavior is correct before plugging into
the full simulator.

Author: GNACODE INC, January 2026
"""

import numpy as np
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from model_verification_1 import (
    LipskiFFTMethod,
    CACFARMethod,
    OSCFARMethod,
    CUSUMMethod,
    SAMPLE_RATE_HZ,
    FRAME_SIZE,
    FRAME_DURATION_S,
)

# =============================================================================
# SYNTHETIC FRAME GENERATOR
# =============================================================================

def generate_noise_frame(rng, noise_power: float = 1.0) -> np.ndarray:
    """White Gaussian noise, no events, no EMI/environmental components."""
    return rng.normal(0.0, np.sqrt(noise_power), FRAME_SIZE)


def generate_event_frame(
    rng,
    noise_power: float = 1.0,
    event_snr_db: float = 18.0,
    event_freq_hz: float = 2.5,
    event_phase: float = 0.0,
) -> np.ndarray:
    """White noise plus a sinusoidal event in the 1-5 Hz band."""
    noise = rng.normal(0.0, np.sqrt(noise_power), FRAME_SIZE)
    event_amp = np.sqrt(noise_power * 10 ** (event_snr_db / 10.0))
    t = np.arange(FRAME_SIZE) / SAMPLE_RATE_HZ
    event = event_amp * np.sin(2 * np.pi * event_freq_hz * t + event_phase)
    return noise + event


def make_methods(seed_offset: int = 0):
    """Instantiate all four classical comparators with their canonical params."""
    return {
        "lipski_fft":  LipskiFFTMethod(node_id=1 + seed_offset),
        "ca_cfar":     CACFARMethod(node_id=1 + seed_offset),
        "os_cfar":     OSCFARMethod(node_id=1 + seed_offset),
        "cusum":       CUSUMMethod(node_id=1 + seed_offset),
    }


# =============================================================================
# TEST 1 -- noise-only stream, measure empirical false-alarm rate
# =============================================================================

def test_noise_only(n_frames: int = 1500, noise_power: float = 1.0, seed: int = 42):
    """Run each detector on pure-noise frames. Report empirical FAR."""
    print("=" * 75)
    print(f"TEST 1: Noise-only stream ({n_frames} frames)")
    print("=" * 75)

    rng = np.random.default_rng(seed)
    methods = make_methods()

    # Each detector has its own calibration window; we count triggers only
    # AFTER calibration is complete to measure the steady-state FAR.
    results = {name: {"cal_frames": 0, "post_cal_frames": 0, "triggers": 0,
                      "strengths": []}
               for name in methods}

    for frame_idx in range(n_frames):
        samples = generate_noise_frame(rng, noise_power)
        for name, method in methods.items():
            trigger, strength = method.process_frame(samples, noise_power)
            # Distinguish calibration vs detection phase
            stats = method.get_stats()
            in_cal = not stats.get("calibrated", True)
            if in_cal:
                results[name]["cal_frames"] += 1
            else:
                results[name]["post_cal_frames"] += 1
                if trigger:
                    results[name]["triggers"] += 1
                results[name]["strengths"].append(strength)

    print()
    print(f"  {'Method':<14} {'Cal':>5} {'Post-cal':>9} {'Triggers':>9} "
          f"{'FAR/frame':>11} {'Mean str':>9} {'Max str':>9}")
    print("  " + "-" * 73)
    for name, r in results.items():
        far = r["triggers"] / r["post_cal_frames"] if r["post_cal_frames"] > 0 else 0.0
        mean_str = np.mean(r["strengths"]) if r["strengths"] else 0.0
        max_str = np.max(r["strengths"]) if r["strengths"] else 0.0
        print(f"  {name:<14} {r['cal_frames']:>5} {r['post_cal_frames']:>9} "
              f"{r['triggers']:>9} {far:>11.4f} {mean_str:>9.3f} {max_str:>9.3f}")
    print()
    return results


# =============================================================================
# TEST 2 -- event-bearing stream at fixed SNR, measure empirical detection rate
# =============================================================================

def test_events(
    n_calibration: int = 200,
    n_detection: int = 1000,
    n_events: int = 50,
    event_duration_frames: int = 4,
    event_snr_db: float = 18.0,
    seed: int = 42,
):
    """Run each detector on a stream of n_calibration noise frames, then
    n_detection mixed (mostly noise + some event) frames. Events span
    event_duration_frames consecutive frames each.

    A frame is labeled "event" if it falls within ANY event's duration.
    This matches the simulator's behavior (multi-frame events) and lets
    CUSUM exercise its pulse-onset/offset state machine properly.

    Detection rate is measured per-EVENT (not per-frame): an event is
    'detected' if any detector trigger occurs within an event-window of
    [event_start - 1, event_end + 1] frames.

    Frame-level FAR is measured as: triggers in NON-event frames divided
    by total non-event frames.
    """
    print("=" * 75)
    print(f"TEST 2: Event-bearing stream (cal={n_calibration}, "
          f"detect={n_detection}, n_events={n_events}, "
          f"duration={event_duration_frames} frames, SNR={event_snr_db} dB)")
    print("=" * 75)

    rng = np.random.default_rng(seed)
    methods = make_methods()

    # Generate event start times — well-separated, in detection phase only
    detection_start = n_calibration
    detection_end = n_calibration + n_detection
    # Reserve detection_end - event_duration_frames as latest possible start
    available_starts = np.arange(
        detection_start + event_duration_frames,
        detection_end - event_duration_frames,
    )
    # Sample event starts with min spacing (event_duration + 2 frames buffer)
    spacing = event_duration_frames + 2
    n_blocks = len(available_starts) // spacing
    if n_blocks < n_events:
        n_events = n_blocks
        print(f"  Reduced n_events to {n_events} (insufficient room)")
    chosen_blocks = rng.choice(np.arange(n_blocks), size=n_events, replace=False)
    event_starts = sorted([
        available_starts[block * spacing + rng.integers(0, spacing)]
        for block in chosen_blocks
    ])
    # Build event-frame set
    event_frames = set()
    for es in event_starts:
        for k in range(event_duration_frames):
            event_frames.add(es + k)
    # Build event-window set (event_start - 1, event_end + 1) for TP scoring
    event_windows = []
    for es in event_starts:
        ee = es + event_duration_frames - 1
        event_windows.append((es - 1, ee + 1))

    n_total = n_calibration + n_detection
    n_event_frames = len(event_frames)
    n_nonevent_detection_frames = n_detection - n_event_frames

    # Track all detector triggers in detection phase
    triggers_per_method = {name: [] for name in methods}
    cal_skipped = {name: 0 for name in methods}

    # Track event-phase: each frame in an event has phase ∈ [0, dur-1]
    # to give consistent sinusoid across frames
    event_freq_per_event = {
        es: rng.uniform(1.5, 4.5) for es in event_starts
    }
    event_phase_per_event = {
        es: rng.uniform(0, 2 * np.pi) for es in event_starts
    }

    for frame_idx in range(n_total):
        is_event_frame = frame_idx in event_frames
        if is_event_frame:
            # Find which event this frame belongs to (use event_starts list)
            es = max(s for s in event_starts if s <= frame_idx)
            samples = generate_event_frame(
                rng, noise_power=1.0, event_snr_db=event_snr_db,
                event_freq_hz=event_freq_per_event[es],
                event_phase=event_phase_per_event[es],
            )
        else:
            samples = generate_noise_frame(rng, noise_power=1.0)

        for name, method in methods.items():
            trigger, strength = method.process_frame(samples, 1.0)
            stats = method.get_stats()
            if not stats.get("calibrated", True):
                cal_skipped[name] += 1
                continue
            if trigger:
                triggers_per_method[name].append(frame_idx)

    # Per-method scoring
    print()
    print(f"  {'Method':<14} {'Events_det':>10} {'DetRate':>8} {'FrmFAR':>8} "
          f"{'Triggers':>9} {'F1*':>6}")
    print("  " + "-" * 72)
    counters_out = {}
    for name in methods:
        triggers = triggers_per_method[name]
        # Per-event detection: did at least one trigger fall in each event window?
        events_detected = 0
        for (lo, hi) in event_windows:
            if any(lo <= t <= hi for t in triggers):
                events_detected += 1
        # Frame-level FAR on non-event frames in detection phase
        non_event_triggers = sum(
            1 for t in triggers
            if t not in event_frames and t >= n_calibration
        )
        frame_far = (non_event_triggers / n_nonevent_detection_frames
                     if n_nonevent_detection_frames > 0 else 0.0)
        det_rate = events_detected / n_events if n_events > 0 else 0.0
        # F1 approximation (event-level precision/recall)
        prec_events = (events_detected /
                       max(1, events_detected + non_event_triggers))
        f1_events = (2 * prec_events * det_rate / (prec_events + det_rate)
                     if (prec_events + det_rate) > 0 else 0.0)
        print(f"  {name:<14} {events_detected:>10} {det_rate:>8.3f} "
              f"{frame_far:>8.3f} {len(triggers):>9} {f1_events:>6.3f}")
        counters_out[name] = {
            "events_detected": events_detected,
            "det_rate": det_rate,
            "frame_far": frame_far,
            "n_triggers": len(triggers),
        }
    print()
    print(f"  *F1* uses event-level precision/recall, approximate.")
    return counters_out


# =============================================================================
# TEST 3 -- SNR sweep, build detection-vs-SNR curves
# =============================================================================

def test_snr_sweep(
    snr_values_db: list,
    n_calibration: int = 200,
    n_detection: int = 600,
    n_events: int = 30,
    event_duration_frames: int = 4,
    seed: int = 42,
):
    """Run the multi-frame event scenario across SNR values.
    Reports event-level detection rate and frame-level FAR per method per SNR.
    """
    print("=" * 75)
    print(f"TEST 3: SNR sweep")
    print(f"  Calibration: {n_calibration} frames, Detection: {n_detection} "
          f"frames, Events per SNR: {n_events} × {event_duration_frames} frames")
    print(f"  SNR values (dB): {snr_values_db}")
    print("=" * 75)

    method_names = ["lipski_fft", "ca_cfar", "os_cfar", "cusum"]
    results_by_snr = {}

    for snr in snr_values_db:
        rng = np.random.default_rng(seed)
        methods = make_methods()

        # Same event scheduling as test_events
        detection_start = n_calibration
        detection_end = n_calibration + n_detection
        available_starts = np.arange(
            detection_start + event_duration_frames,
            detection_end - event_duration_frames,
        )
        spacing = event_duration_frames + 2
        n_blocks = len(available_starts) // spacing
        n_evts = min(n_events, n_blocks)
        chosen_blocks = rng.choice(np.arange(n_blocks), size=n_evts, replace=False)
        event_starts = sorted([
            available_starts[block * spacing + rng.integers(0, spacing)]
            for block in chosen_blocks
        ])
        event_frames = set()
        for es in event_starts:
            for k in range(event_duration_frames):
                event_frames.add(es + k)
        event_windows = [(es - 1, es + event_duration_frames) for es in event_starts]
        event_freq_per_event = {es: rng.uniform(1.5, 4.5) for es in event_starts}
        event_phase_per_event = {es: rng.uniform(0, 2 * np.pi) for es in event_starts}

        n_total = n_calibration + n_detection
        n_nonevent_detection_frames = n_detection - len(event_frames)

        triggers_per_method = {name: [] for name in methods}

        for frame_idx in range(n_total):
            is_event_frame = frame_idx in event_frames
            if is_event_frame:
                es = max(s for s in event_starts if s <= frame_idx)
                samples = generate_event_frame(
                    rng, noise_power=1.0, event_snr_db=snr,
                    event_freq_hz=event_freq_per_event[es],
                    event_phase=event_phase_per_event[es],
                )
            else:
                samples = generate_noise_frame(rng, noise_power=1.0)

            for name, method in methods.items():
                trigger, _ = method.process_frame(samples, 1.0)
                stats = method.get_stats()
                if not stats.get("calibrated", True):
                    continue
                if trigger:
                    triggers_per_method[name].append(frame_idx)

        scored = {}
        for name in method_names:
            triggers = triggers_per_method[name]
            events_detected = 0
            for (lo, hi) in event_windows:
                if any(lo <= t <= hi for t in triggers):
                    events_detected += 1
            non_event_triggers = sum(
                1 for t in triggers
                if t not in event_frames and t >= n_calibration
            )
            scored[name] = {
                "det": events_detected / n_evts if n_evts > 0 else 0.0,
                "far": (non_event_triggers / n_nonevent_detection_frames
                        if n_nonevent_detection_frames > 0 else 0.0),
            }
        results_by_snr[snr] = scored

    print()
    print(f"  {'SNR (dB)':<10}", end="")
    for name in method_names:
        print(f" {name + ' DR':>15}", end="")
    print()
    print("  " + "-" * (10 + 16 * len(method_names)))

    for snr in snr_values_db:
        print(f"  {snr:<10.1f}", end="")
        for name in method_names:
            print(f" {results_by_snr[snr][name]['det']:>15.3f}", end="")
        print()

    print()
    print(f"  {'SNR (dB)':<10}", end="")
    for name in method_names:
        print(f" {name + ' FAR':>15}", end="")
    print()
    print("  " + "-" * (10 + 16 * len(method_names)))

    for snr in snr_values_db:
        print(f"  {snr:<10.1f}", end="")
        for name in method_names:
            print(f" {results_by_snr[snr][name]['far']:>15.3f}", end="")
        print()
    print()
    return results_by_snr


# =============================================================================
# MAIN
# =============================================================================

if __name__ == "__main__":
    print()
    print("Classical Comparators Verification Harness")
    print(f"Frame: {FRAME_SIZE} samples @ {SAMPLE_RATE_HZ} Hz "
          f"= {FRAME_DURATION_S:.2f} sec")
    print()

    # Test 1: noise-only -- empirical false-alarm rate
    test_noise_only(n_frames=1500, seed=42)

    # Test 2a: events at canonical 18 dB SNR (matches simulator default)
    # Multi-frame events to exercise CUSUM's pulse onset/offset
    test_events(
        n_calibration=200,
        n_detection=1500,
        n_events=80,
        event_duration_frames=4,
        event_snr_db=18.0,
        seed=42,
    )

    # Test 2b: events at lower SNR (12 dB)
    test_events(
        n_calibration=200,
        n_detection=1500,
        n_events=80,
        event_duration_frames=4,
        event_snr_db=12.0,
        seed=42,
    )

    # Test 3: SNR sweep
    test_snr_sweep(
        snr_values_db=[6, 9, 12, 15, 18, 21, 24],
        n_calibration=200,
        n_detection=600,
        n_events=30,
        event_duration_frames=4,
        seed=42,
    )

    print("=" * 75)
    print("Verification complete.")
    print("=" * 75)