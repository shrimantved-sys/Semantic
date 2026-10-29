import argparse
import asyncio
import json
import os
import statistics
import time
from typing import Optional
import httpx

SERVER_URL = "http://127.0.0.1:8004"
DEFAULT_OUTPUT_FILE = "benchmark_results.json"

DEFAULT_WORKLOAD = [
    "Artificial intelligence is",
    "Python programming allows developers to quickly build web apps",
    "Data science techniques help",
    "Neural networks can learn complex pattern representations",
    "Artificial intelligence is",
    "Python programming allows developers to quickly build web apps",
    "Data science techniques help",
    "Neural networks can learn complex pattern representations",
    "Artificial intelligence is",
    "Python programming allows developers to quickly build web apps",
    "Data science techniques help",
    "Neural networks can learn complex pattern representations"
]


async def send_benchmark_request(
    client: httpx.AsyncClient,
    prompt: str,
    enable_cache: bool,
    enable_semantic_cache: bool = True
) -> dict:
    url = f"{SERVER_URL}/generate"
    params = {
        "enable_cache": "true" if enable_cache else "false",
        "enable_semantic_cache": "true" if enable_semantic_cache else "false"
    }
    start = time.time()
    res = await client.post(url, params=params, json={"prompt": prompt, "max_new_tokens": 15}, timeout=60.0)
    latency_ms = (time.time() - start) * 1000.0

    if res.status_code != 200:
        raise RuntimeError(f"HTTP {res.status_code}: {res.text}")

    data = res.json()
    data["client_latency_ms"] = latency_ms
    return data


async def run_benchmark_mode(
    mode_name: str,
    enable_cache: bool,
    enable_semantic_cache: bool = True,
    prompts: Optional[list[str]] = None
) -> tuple[dict, list[dict]]:
    workload = prompts or DEFAULT_WORKLOAD
    print(f"\nRunning {mode_name} ({len(workload)} requests)...")

    async with httpx.AsyncClient() as client:
        await client.post(f"{SERVER_URL}/clear_cache")
        start_time = time.time()

        if enable_cache and len(workload) >= 4:
            split = len(workload) // 3 if len(workload) >= 6 else 2
            # Cold pass followed by warm pass
            p1 = [send_benchmark_request(client, p, enable_cache, enable_semantic_cache) for p in workload[:split]]
            res1 = await asyncio.gather(*p1)
            p2 = [send_benchmark_request(client, p, enable_cache, enable_semantic_cache) for p in workload[split:]]
            res2 = await asyncio.gather(*p2)
            results = res1 + res2
        else:
            tasks = [send_benchmark_request(client, p, enable_cache, enable_semantic_cache) for p in workload]
            results = await asyncio.gather(*tasks)

        duration = time.time() - start_time

    latencies = sorted(r["total_latency_ms"] for r in results)
    proc_times = [r["processing_time_ms"] for r in results]
    cache_hits = [1 if r["cache_hit"] else 0 for r in results]

    total_gen_toks = sum(r.get("generated_tokens", 0) for r in results)

    metrics = {
        "mode": mode_name,
        "cache_enabled": enable_cache,
        "semantic_cache_enabled": enable_semantic_cache,
        "requests_count": len(workload),
        "duration_sec": round(duration, 3),
        "rps": round(len(workload) / duration, 2),
        "tps": round(total_gen_toks / duration, 2),
        "p50_ms": round(statistics.median(latencies), 2),
        "p90_ms": round(latencies[int(len(latencies) * 0.90) - 1], 2),
        "p99_ms": round(latencies[-1], 2),
        "avg_proc_ms": round(sum(proc_times) / len(proc_times), 2),
        "hit_rate_pct": round((sum(cache_hits) / len(cache_hits)) * 100.0, 1),
        "prompt_tokens": sum(r.get("prompt_tokens", 0) for r in results),
        "gen_tokens": total_gen_toks,
        "pad_tokens": sum(r.get("padded_tokens", 0) for r in results),
        "total_tokens": sum(r.get("total_model_tokens", 0) for r in results)
    }

    print(f"  Duration  : {metrics['duration_sec']}s | Throughput: {metrics['rps']} req/s ({metrics['tps']} tok/s)")
    print(f"  Latency   : P50={metrics['p50_ms']}ms, P90={metrics['p90_ms']}ms, P99={metrics['p99_ms']}ms")
    print(f"  Cache Hit : {metrics['hit_rate_pct']}%")
    return metrics, results


def save_benchmark_results(data: dict, output_file: str) -> None:
    try:
        with open(output_file, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
        print(f"\n[Saved] Benchmark results successfully stored to '{output_file}'.")
    except Exception as e:
        print(f"\n[Error] Failed to save benchmark results to '{output_file}': {e}")


async def main():
    parser = argparse.ArgumentParser(description="Performance benchmark for Two-Tier Cache LLM Inference Server")
    parser.add_argument("--prompt", type=str, help="Single custom prompt")
    parser.add_argument("--prompts", nargs="+", help="Multiple custom prompts")
    parser.add_argument("--output", "-o", type=str, default=DEFAULT_OUTPUT_FILE, help=f"Output file path for results (default: {DEFAULT_OUTPUT_FILE})")
    parser.add_argument("--no-save", action="store_true", help="Disable saving results to file")
    parser.add_argument("--interactive", action="store_true", help="Launch interactive prompt CLI")
    args = parser.parse_args()

    if args.interactive:
        import prompt_cli
        prompt_cli.main()
        return

    workload = [args.prompt] * 4 if args.prompt else args.prompts
    actual_workload = workload or DEFAULT_WORKLOAD

    print("=" * 80)
    print(" LLM SERVER TWO-TIER CACHE BENCHMARK SUITE")
    print("=" * 80)

    m1, r1 = await run_benchmark_mode("Mode 1 (Baseline: Direct LLM, No Cache)", False, False, workload)
    m2, r2 = await run_benchmark_mode("Mode 2 (Tier 1 Exact Match Cache)", True, False, workload)
    m3, r3 = await run_benchmark_mode("Mode 3 (Two-Tier Exact + Semantic Cache)", True, True, workload)

    print("\n" + "=" * 80)
    print(f"{'Metric':<30} | {'Mode 1 (No Cache)':<17} | {'Mode 2 (Tier 1)':<17} | {'Mode 3 (Two-Tier)':<17}")
    print("-" * 80)
    print(f"{'Duration':<30} | {m1['duration_sec']:<15} s | {m2['duration_sec']:<15} s | {m3['duration_sec']:<15} s")
    print(f"{'Throughput (req/s)':<30} | {m1['rps']:<17} | {m2['rps']:<17} | {m3['rps']:<17}")
    print(f"{'Throughput (tokens/s)':<30} | {m1['tps']:<17} | {m2['tps']:<17} | {m3['tps']:<17}")
    print(f"{'P50 Latency':<30} | {m1['p50_ms']:<14} ms | {m2['p50_ms']:<14} ms | {m3['p50_ms']:<14} ms")
    print(f"{'P90 Latency':<30} | {m1['p90_ms']:<14} ms | {m2['p90_ms']:<14} ms | {m3['p90_ms']:<14} ms")
    print(f"{'P99 Latency':<30} | {m1['p99_ms']:<14} ms | {m2['p99_ms']:<14} ms | {m3['p99_ms']:<14} ms")
    print(f"{'Avg Processing Time':<30} | {m1['avg_proc_ms']:<14} ms | {m2['avg_proc_ms']:<14} ms | {m3['avg_proc_ms']:<14} ms")
    print(f"{'Cache Hit Rate':<30} | {m1['hit_rate_pct']:<15} % | {m2['hit_rate_pct']:<15} % | {m3['hit_rate_pct']:<15} %")
    print("=" * 80 + "\n")

    benchmark_record = {
        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
        "server_url": SERVER_URL,
        "workload_size": len(actual_workload),
        "prompts": actual_workload,
        "summary": {
            "mode_1_baseline": m1,
            "mode_2_tier1_cache": m2,
            "mode_3_twotier_cache": m3
        },
        "speedup": {
            "tier1_cache_speedup": round(m2["rps"] / max(m1["rps"], 0.001), 2),
            "two_tier_cache_speedup": round(m3["rps"] / max(m1["rps"], 0.001), 2),
            "p50_latency_reduction_pct": round((1 - (m3["p50_ms"] / max(m1["p50_ms"], 0.001))) * 100, 1)
        },
        "detailed_results": {
            "mode_1": r1,
            "mode_2": r2,
            "mode_3": r3
        }
    }

    if not args.no_save:
        save_benchmark_results(benchmark_record, args.output)


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except httpx.ConnectError:
        print(f"Error: Could not connect to server at {SERVER_URL}. Please start server.py first.")
