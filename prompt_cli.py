#!/usr/bin/env python3
"""
Interactive CLI and prompt runner for the Two-Tier LLM Inference Server.
"""

import argparse
import asyncio
import json
import sys
import time
import httpx

DEFAULT_SERVER_URL = "http://127.0.0.1:8004"


async def send_prompt(
    client: httpx.AsyncClient,
    server_url: str,
    prompt: str,
    max_new_tokens: int = 20,
    enable_cache: bool = True,
    enable_semantic_cache: bool = True,
    similarity_threshold: float = 0.90,
    cache_ttl: float = 3600.0
) -> dict:
    params = {
        "enable_cache": "true" if enable_cache else "false",
        "enable_semantic_cache": "true" if enable_semantic_cache else "false",
        "similarity_threshold": similarity_threshold,
        "cache_ttl": cache_ttl
    }
    url = f"{server_url}/generate"
    payload = {"prompt": prompt, "max_new_tokens": max_new_tokens}

    start = time.time()
    resp = await client.post(url, params=params, json=payload)
    latency_ms = (time.time() - start) * 1000.0

    if resp.status_code != 200:
        print(f"\nError [{resp.status_code}]: {resp.text}")
        return {}

    data = resp.json()
    data["client_latency_ms"] = round(latency_ms, 2)
    return data


def print_result_card(res: dict) -> None:
    if not res:
        return

    match_type = res.get("match_type", "none")
    status_label = {
        "exact": "Tier 1 Exact Match Hit",
        "semantic": f"Tier 2 Semantic Hit (Score: {res.get('similarity_score')})",
        "none": "Cache Miss (Model Inference)"
    }.get(match_type, "Unknown")

    print("\n" + "=" * 70)
    print(f" Request ID  : {res.get('request_id', 'N/A')}")
    print(f" Prompt      : \"{res.get('prompt', '')}\"")
    print("-" * 70)
    print(f" Response    :\n {res.get('generated_text', '').strip()}")
    print("-" * 70)
    print(f" Cache Status: {status_label}")
    if match_type in ("exact", "semantic") and res.get("matched_prompt"):
        print(f" Matched Key : \"{res.get('matched_prompt')}\"")

    print(f" Latency     : {res.get('total_latency_ms', 0):.2f} ms (Server) | {res.get('client_latency_ms', 0):.2f} ms (Client)")
    print(f" Model Time  : {res.get('processing_time_ms', 0):.2f} ms")
    print(f" Tokens      : {res.get('prompt_tokens', 0)} prompt + {res.get('generated_tokens', 0)} generated = {res.get('total_model_tokens', 0)} total")
    print("=" * 70)


def save_cli_results(results: list[dict], output_file: str) -> None:
    try:
        with open(output_file, "w", encoding="utf-8") as f:
            json.dump(results, f, indent=2)
        print(f"\n[Saved] Query results successfully stored to '{output_file}'.")
    except Exception as e:
        print(f"\n[Error] Failed to save results to '{output_file}': {e}")


async def interactive_session(args):
    print("\n" + "=" * 70)
    print(" Interactive Inference Shell (Type 'quit' or 'exit' to stop)")
    print(f" Target: {args.url} | Tokens: {args.max_tokens} | Cache: {'ON' if not args.no_cache else 'OFF'}")
    print("=" * 70 + "\n")

    session_results = []
    async with httpx.AsyncClient(timeout=60.0) as client:
        while True:
            try:
                user_input = input("Prompt > ").strip()
                if not user_input:
                    continue
                if user_input.lower() in ("exit", "quit"):
                    print("Goodbye!")
                    break

                res = await send_prompt(
                    client=client,
                    server_url=args.url,
                    prompt=user_input,
                    max_new_tokens=args.max_tokens,
                    enable_cache=not args.no_cache,
                    enable_semantic_cache=not args.no_semantic_cache,
                    similarity_threshold=args.threshold,
                    cache_ttl=args.ttl
                )
                if res:
                    session_results.append(res)
                print_result_card(res)
            except (KeyboardInterrupt, EOFError):
                print("\nSession ended.")
                break

    if args.output and session_results:
        save_cli_results(session_results, args.output)


async def run_prompt_file(args, prompts: list[str]):
    results = []
    async with httpx.AsyncClient(timeout=60.0) as client:
        if args.concurrent:
            tasks = [
                send_prompt(
                    client=client,
                    server_url=args.url,
                    prompt=p,
                    max_new_tokens=args.max_tokens,
                    enable_cache=not args.no_cache,
                    enable_semantic_cache=not args.no_semantic_cache,
                    similarity_threshold=args.threshold,
                    cache_ttl=args.ttl
                )
                for p in prompts
            ]
            results = await asyncio.gather(*tasks)
            for res in results:
                print_result_card(res)
        else:
            for i, p in enumerate(prompts, 1):
                print(f"Processing [{i}/{len(prompts)}]: {p}")
                res = await send_prompt(
                    client=client,
                    server_url=args.url,
                    prompt=p,
                    max_new_tokens=args.max_tokens,
                    enable_cache=not args.no_cache,
                    enable_semantic_cache=not args.no_semantic_cache,
                    similarity_threshold=args.threshold,
                    cache_ttl=args.ttl
                )
                if res:
                    results.append(res)
                print_result_card(res)

    if args.output and results:
        save_cli_results(results, args.output)


def main():
    parser = argparse.ArgumentParser(description="Inference CLI for Two-Tier Cache Server")
    parser.add_argument("--prompt", type=str, help="Single prompt to execute")
    parser.add_argument("--file", type=str, help="File with prompt lines")
    parser.add_argument("--output", "-o", type=str, help="Output file path to save results JSON")
    parser.add_argument("--concurrent", action="store_true", help="Send file prompts concurrently")
    parser.add_argument("--max-tokens", type=int, default=20, help="Max tokens to generate (default: 20)")
    parser.add_argument("--no-cache", action="store_true", help="Disable server cache")
    parser.add_argument("--no-semantic-cache", action="store_true", help="Disable semantic Tier 2 cache")
    parser.add_argument("--threshold", type=float, default=0.90, help="Similarity threshold (default: 0.90)")
    parser.add_argument("--ttl", type=float, default=3600.0, help="Cache TTL in seconds (default: 3600.0)")
    parser.add_argument("--url", default=DEFAULT_SERVER_URL, help=f"Server URL (default: {DEFAULT_SERVER_URL})")

    args = parser.parse_args()

    if args.prompt:
        async def single_run():
            async with httpx.AsyncClient(timeout=60.0) as client:
                res = await send_prompt(
                    client=client,
                    server_url=args.url,
                    prompt=args.prompt,
                    max_new_tokens=args.max_tokens,
                    enable_cache=not args.no_cache,
                    enable_semantic_cache=not args.no_semantic_cache,
                    similarity_threshold=args.threshold,
                    cache_ttl=args.ttl
                )
                print_result_card(res)
                if args.output and res:
                    save_cli_results([res], args.output)
        asyncio.run(single_run())
    elif args.file:
        try:
            with open(args.file, "r", encoding="utf-8") as f:
                prompts = [line.strip() for line in f if line.strip() and not line.strip().startswith("#")]
            asyncio.run(run_prompt_file(args, prompts))
        except Exception as e:
            print(f"Error reading file '{args.file}': {e}")
    else:
        asyncio.run(interactive_session(args))


if __name__ == "__main__":
    main()

