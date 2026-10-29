#!/usr/bin/env python3
"""
CLI runner for demonstrating Two-Tier Cache (Tier 1 Exact + Tier 2 Semantic)
against prompts loaded from a text file.
"""

import argparse
import asyncio
import json
import os
import sys
import time
from typing import Optional
import httpx

DEFAULT_SERVER_URL = "http://127.0.0.1:8004"
DEFAULT_PROMPTS_FILE = "semantic_prompts.txt"
DEFAULT_OUTPUT_FILE = "semantic_prompt_results.json"


def load_prompts(filepath: str) -> list[str]:
    if not os.path.exists(filepath):
        print(f"Error: Prompt file '{filepath}' not found.")
        sys.exit(1)

    with open(filepath, "r", encoding="utf-8") as f:
        prompts = [line.strip() for line in f if line.strip() and not line.strip().startswith("#")]

    if not prompts:
        print(f"Error: No valid prompts found in '{filepath}'.")
        sys.exit(1)

    return prompts


async def run_prompts(
    server_url: str,
    prompts: list[str],
    threshold: float,
    max_tokens: int,
    output_file: Optional[str] = DEFAULT_OUTPUT_FILE
):
    print(f"\nRunning {len(prompts)} prompts against {server_url} (Threshold: {threshold}, Max Tokens: {max_tokens})\n" + "-" * 70)

    async with httpx.AsyncClient(timeout=60.0) as client:
        try:
            status_resp = await client.get(f"{server_url}/api/status")
            if status_resp.status_code != 200:
                print(f"Server returned HTTP {status_resp.status_code}")
                return
        except Exception:
            print(f"Could not connect to {server_url}. Ensure server.py is running.")
            return

        # Start with a clean cache
        await client.post(f"{server_url}/clear_cache")

        start_all = time.time()
        results = []
        for i, prompt in enumerate(prompts, 1):
            url = f"{server_url}/generate"
            params = {
                "enable_cache": "true",
                "enable_semantic_cache": "true",
                "similarity_threshold": threshold
            }
            start = time.time()
            resp = await client.post(url, params=params, json={"prompt": prompt, "max_new_tokens": max_tokens})
            latency_ms = (time.time() - start) * 1000.0

            if resp.status_code != 200:
                print(f"[{i}/{len(prompts)}] Error ({resp.status_code}): {resp.text}")
                continue

            data = resp.json()
            data["client_latency_ms"] = round(latency_ms, 2)
            results.append(data)

            match_type = data.get("match_type", "none")
            badge = {
                "exact": "[Tier 1 Exact Hit]",
                "semantic": f"[Tier 2 Semantic Hit (Score: {data.get('similarity_score')})]",
                "none": "[Cache Miss -> Inference]"
            }.get(match_type, "[Unknown]")

            print(f"[{i}/{len(prompts)}] \"{prompt}\"")
            print(f"  Result : {badge}")
            if match_type == "semantic":
                print(f"  Matched: \"{data.get('matched_prompt')}\"")
            print(f"  Latency: {data.get('total_latency_ms', 0):.2f} ms")
            print(f"  Output : \"{data.get('generated_text', '').strip()[:80]}...\"\n")

        total_duration = time.time() - start_all
        exact = sum(1 for r in results if r.get("match_type") == "exact")
        semantic = sum(1 for r in results if r.get("match_type") == "semantic")
        misses = sum(1 for r in results if r.get("match_type") == "none")
        avg_latency = (sum(r.get("total_latency_ms", 0) for r in results) / len(results)) if results else 0.0

        print("=" * 70)
        print("Execution Summary:")
        print(f"  Total Processed    : {len(results)}")
        print(f"  Exact Matches      : {exact}")
        print(f"  Semantic Matches   : {semantic}")
        print(f"  Full Misses        : {misses}")
        print(f"  Cache Hit Rate     : {((exact + semantic) / max(len(results), 1)) * 100:.1f}%")
        print(f"  Average Latency    : {avg_latency:.2f} ms")
        print("=" * 70 + "\n")

        results_record = {
            "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
            "server_url": server_url,
            "threshold": threshold,
            "max_tokens": max_tokens,
            "total_duration_sec": round(total_duration, 3),
            "summary": {
                "total_prompts": len(results),
                "exact_matches": exact,
                "semantic_matches": semantic,
                "cache_misses": misses,
                "cache_hit_rate_pct": round(((exact + semantic) / max(len(results), 1)) * 100.0, 1),
                "avg_latency_ms": round(avg_latency, 2)
            },
            "results": results
        }

        if output_file:
            try:
                with open(output_file, "w", encoding="utf-8") as f:
                    json.dump(results_record, f, indent=2)
                print(f"[Saved] Semantic prompt results successfully stored to '{output_file}'.\n")
            except Exception as e:
                print(f"[Error] Failed to save results to '{output_file}': {e}\n")


def main():
    parser = argparse.ArgumentParser(description="Two-Tier Semantic Cache Prompt Runner")
    parser.add_argument("--file", default=DEFAULT_PROMPTS_FILE, help=f"Prompts file (default: {DEFAULT_PROMPTS_FILE})")
    parser.add_argument("--threshold", type=float, default=0.88, help="Semantic similarity threshold (default: 0.88)")
    parser.add_argument("--max-tokens", type=int, default=25, help="Tokens to generate (default: 25)")
    parser.add_argument("--output", "-o", default=DEFAULT_OUTPUT_FILE, help=f"Path to store results JSON (default: {DEFAULT_OUTPUT_FILE})")
    parser.add_argument("--no-save", action="store_true", help="Disable saving results to file")
    parser.add_argument("--url", default=DEFAULT_SERVER_URL, help=f"Server URL (default: {DEFAULT_SERVER_URL})")

    args = parser.parse_args()
    prompts = load_prompts(args.file)
    out = None if args.no_save else args.output
    asyncio.run(run_prompts(args.url, prompts, args.threshold, args.max_tokens, output_file=out))


if __name__ == "__main__":
    main()
