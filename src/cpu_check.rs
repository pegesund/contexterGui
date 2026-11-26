//! One-time CPU speed check at first startup.
//!
//! Local Whisper speech recognition needs a reasonably fast CPU. On weak
//! school PCs (2-core Pentium/Celeron, no AVX2) even the smallest model runs
//! at ~0.1x realtime, so voice input takes minutes. Those machines use the
//! cloud transcription mode instead; this check decides which machines are
//! "weak" once, and the result is stored in settings.

use std::time::Instant;

pub struct CpuCheck {
    pub weak: bool,
    pub detail: String,
}

/// Below this many logical cores local STT is too slow.
const MIN_CORES: usize = 4;
/// Single-thread multiply-add throughput (millions of ops per second) below
/// which the CPU is treated as weak — a safety net; missing AVX2 is the main
/// signal. Measured: Pentium 4405U ~520 (also no AVX2), Apple M-series ~1580;
/// the loop scales with clock speed (~0.25 x MHz on x86).
const MIN_MOPS: f64 = 400.0;

pub fn run() -> CpuCheck {
    let cores = std::thread::available_parallelism().map(|n| n.get()).unwrap_or(1);
    let avx2 = has_avx2();
    let mops = bench_mops();
    let weak = cores < MIN_CORES || !avx2 || mops < MIN_MOPS;
    CpuCheck {
        weak,
        detail: format!("cores={} avx2={} mops={:.0} -> {}", cores, avx2, mops,
            if weak { "weak" } else { "ok" }),
    }
}

#[cfg(target_arch = "x86_64")]
fn has_avx2() -> bool {
    std::arch::is_x86_feature_detected!("avx2")
}

/// Non-x86 (Apple Silicon) CPUs have their own fast SIMD; Whisper is fine.
#[cfg(not(target_arch = "x86_64"))]
fn has_avx2() -> bool {
    true
}

/// ~100 ms single-thread multiply-add loop; returns millions of ops/second.
fn bench_mops() -> f64 {
    const N: usize = 4096;
    let a: Vec<f32> = (0..N).map(|i| (i % 7) as f32 * 0.5).collect();
    let b: Vec<f32> = (0..N).map(|i| (i % 5) as f32 * 0.25).collect();
    let mut acc = 0.0f32;
    let mut ops: u64 = 0;
    let start = Instant::now();
    while start.elapsed().as_millis() < 100 {
        for i in 0..N {
            acc = acc * 0.999 + a[i] * b[i];
        }
        ops += N as u64 * 2;
    }
    std::hint::black_box(acc);
    ops as f64 / start.elapsed().as_secs_f64() / 1e6
}
