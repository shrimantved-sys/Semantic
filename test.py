import asyncio
import json
import os
import time
import httpx

SERVER_URL = "http://127.0.0.1:8004"
LOG_FILE = "telemetry_metrics.jsonl"


async def run_tests():
    print(f"Connecting to test server at {SERVER_URL}...")
    if os.path.exists(LOG_FILE):
        os.remove(LOG_FILE)

    async with httpx.AsyncClient(timeout=30.0) as client:
        # 1. Health & Playground UI
        print("[1/9] Verifying server status and playground UI...")
        r = await client.get(f"{SERVER_URL}/api/status")
        assert r.status_code == 200, f"Status check failed: {r.status_code}"
        assert r.json().get("status") == "ok"

        r_ui = await client.get(f"{SERVER_URL}/playground")
        assert r_ui.status_code == 200, f"Playground check failed: {r_ui.status_code}"

        # 2. Querying with empty cache
        print("[2/9] Testing cold query with empty cache...")
        await client.post(f"{SERVER_URL}/clear_cache")
        prompt_q1 = "What is the boiling point of water at sea level?"
        r_empty = await client.post(
            f"{SERVER_URL}/generate",
            json={"prompt": prompt_q1, "max_new_tokens": 15}
        )
        assert r_empty.status_code == 200
        d_empty = r_empty.json()
        assert d_empty["cache_hit"] is False
        assert d_empty["match_type"] == "none"

        # 3. Tier 1: Exact Match Cache Hit
        print("[3/9] Testing Tier 1 exact-match hit...")
        r_exact = await client.post(
            f"{SERVER_URL}/generate",
            json={"prompt": prompt_q1, "max_new_tokens": 15}
        )
        assert r_exact.status_code == 200
        d_exact = r_exact.json()
        assert d_exact["cache_hit"] is True
        assert d_exact["match_type"] == "exact"
        assert d_exact["similarity_score"] == 1.0
        assert d_exact["matched_prompt"] == prompt_q1
        assert d_exact["generated_text"] == d_empty["generated_text"]

        # 4. Tier 2: Semantic Match Cache Hit
        print("[4/9] Testing Tier 2 semantic similarity hit...")
        paraphrase = "At what temperature does water boil at sea level?"
        r_sem = await client.post(
            f"{SERVER_URL}/generate?similarity_threshold=0.88",
            json={"prompt": paraphrase, "max_new_tokens": 15}
        )
        assert r_sem.status_code == 200
        d_sem = r_sem.json()
        assert d_sem["cache_hit"] is True
        assert d_sem["match_type"] == "semantic"
        assert d_sem["similarity_score"] is not None and d_sem["similarity_score"] >= 0.88
        assert d_sem["matched_prompt"] == prompt_q1
        assert d_sem["generated_text"] == d_empty["generated_text"]

        # 5. Miss on distinct query
        print("[5/9] Testing full cache miss on distinct question...")
        r_miss = await client.post(
            f"{SERVER_URL}/generate?similarity_threshold=0.90",
            json={"prompt": "Who wrote Romeo and Juliet?", "max_new_tokens": 15}
        )
        assert r_miss.status_code == 200
        d_miss = r_miss.json()
        assert d_miss["cache_hit"] is False
        assert d_miss["match_type"] == "none"
        assert d_miss["similarity_score"] is None

        # 6. TTL eviction
        print("[6/9] Testing TTL expiration and auto-eviction...")
        ttl_prompt = "Short lived prompt query"
        await client.post(
            f"{SERVER_URL}/generate?cache_ttl=1.0",
            json={"prompt": ttl_prompt, "max_new_tokens": 10}
        )
        r_hit = await client.post(
            f"{SERVER_URL}/generate?cache_ttl=1.0",
            json={"prompt": ttl_prompt, "max_new_tokens": 10}
        )
        assert r_hit.json()["cache_hit"] is True

        await asyncio.sleep(1.2)
        r_expired = await client.post(
            f"{SERVER_URL}/generate?cache_ttl=1.0",
            json={"prompt": ttl_prompt, "max_new_tokens": 10}
        )
        assert r_expired.json()["cache_hit"] is False

        # 7. Manual cache clear (single and all)
        print("[7/9] Testing manual cache invalidation...")
        target_prompt = "Who wrote Romeo and Juliet?"
        r_clear = await client.post(f"{SERVER_URL}/clear_cache?prompt={target_prompt}")
        assert r_clear.status_code == 200

        r_check = await client.post(
            f"{SERVER_URL}/generate",
            json={"prompt": target_prompt, "max_new_tokens": 15}
        )
        assert r_check.json()["cache_hit"] is False

        r_clear_all = await client.post(f"{SERVER_URL}/clear_cache")
        assert r_clear_all.json()["remaining_entries"] == 0

        # 8. Malformed input error handling
        print("[8/9] Testing malformed input validation...")
        r_bad = await client.post(
            f"{SERVER_URL}/generate",
            json={"prompt": "   ", "max_new_tokens": 15}
        )
        assert r_bad.status_code == 400

        # 9. Stored Result API Endpoints
        print("[9/10] Testing stored result APIs...")
        r_bench = await client.get(f"{SERVER_URL}/api/benchmark_results")
        assert r_bench.status_code in (200, 404)
        if r_bench.status_code == 200:
            assert "summary" in r_bench.json()

        r_eval = await client.get(f"{SERVER_URL}/api/evaluation_results")
        assert r_eval.status_code in (200, 404)
        if r_eval.status_code == 200:
            assert "threshold_sweep" in r_eval.json()

        r_prompt = await client.get(f"{SERVER_URL}/api/prompt_results")
        assert r_prompt.status_code in (200, 404)
        if r_prompt.status_code == 200:
            assert "results" in r_prompt.json()

    # 10. Verify telemetry log format
    print("[10/10] Verifying telemetry metrics file...")
    assert os.path.exists(LOG_FILE), f"Log file '{LOG_FILE}' was not created."
    with open(LOG_FILE, "r", encoding="utf-8") as f:
        records = [json.loads(line.strip()) for line in f if line.strip()]

    required_keys = [
        "request_id", "timestamp", "queue_time_ms", "processing_time_ms",
        "total_latency_ms", "batch_size", "cache_hit", "match_type", "similarity_score", "matched_prompt"
    ]
    for i, log in enumerate(records, 1):
        for k in required_keys:
            assert k in log, f"Missing key '{k}' in record #{i}"

    print(f"\nAll tests passed successfully ({len(records)} telemetry records verified).")


if __name__ == "__main__":
    try:
        asyncio.run(run_tests())
    except httpx.ConnectError:
        print(f"Error: Could not connect to server at {SERVER_URL}.")
        print("Please ensure server.py is running before executing tests.")
