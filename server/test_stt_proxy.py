"""Tests for stt_proxy with a fake speech backend (no network)."""
import asyncio
import os
import struct
import time

os.environ["SPELL_PROXY_NO_APP"] = "1"

import pytest
from fastapi.testclient import TestClient

import stt_proxy


def wav(samples=16000, rate=16000, channels=1, bits=16):
    data = b"\x01\x00" * samples
    fmt = struct.pack("<HHIIHH", 1, channels, rate, rate * channels * bits // 8, channels * bits // 8, bits)
    return (b"RIFF" + struct.pack("<I", 36 + len(data)) + b"WAVE"
            + b"fmt " + struct.pack("<I", 16) + fmt
            + b"data" + struct.pack("<I", len(data)) + data)


def make_client(rate=30, recognizer=None):
    cfg = stt_proxy.Config({"SPELL_APP_KEY": "k1", "SPELL_GOOGLE_API_KEY": "g", "SPELL_RATE_PER_MIN": str(rate)})
    calls = []

    async def fake(w, pcm, lang):
        calls.append((len(pcm), lang))
        return "hei verden"

    app = stt_proxy.create_app(cfg, recognizer or fake)
    return TestClient(app), calls


def post(client, body, key="k1", lang="nb"):
    return client.post("/transcribe", content=body,
                       headers={"Content-Type": "audio/wav", "X-Spell-Key": key, "X-Spell-Lang": lang})


def test_transcribes_valid_wav():
    c, calls = make_client()
    r = post(c, wav())
    assert r.status_code == 200 and r.json() == {"text": "hei verden"}
    assert calls == [(32000, "nb-NO")]


def test_language_mapping():
    c, calls = make_client()
    for code, want in [("no", "nb-NO"), ("nn", "nn-NO"), ("en", "en-US"), ("nb-NO", "nb-NO")]:
        assert post(c, wav(), lang=code).status_code == 200
        assert calls[-1][1] == want
    assert post(c, wav(), lang="de").status_code == 400


def test_rejects_bad_app_key():
    c, calls = make_client()
    assert post(c, wav(), key="wrong").status_code == 403
    assert post(c, wav(), key="").status_code == 403
    assert calls == []


def test_rejects_wrong_audio_format():
    c, _ = make_client()
    assert post(c, b"not audio at all" * 10).status_code == 400
    assert post(c, wav(rate=48000)).status_code == 400
    assert post(c, wav(channels=2)).status_code == 400


def test_rejects_too_long_audio():
    c, _ = make_client()
    assert post(c, wav(samples=16000 * 75)).status_code == 413


def test_empty_audio_returns_empty_text_without_provider_call():
    c, calls = make_client()
    r = post(c, wav(samples=0))
    assert r.json() == {"text": ""} and calls == []


def test_rate_limit_per_ip():
    c, _ = make_client(rate=2)
    assert post(c, wav()).status_code == 200
    assert post(c, wav()).status_code == 200
    assert post(c, wav()).status_code == 429


def test_provider_error_is_not_echoed():
    async def boom(w, pcm, lang):
        raise RuntimeError("secret detail key=abc")
    c, _ = make_client(recognizer=boom)
    r = post(c, wav())
    assert r.status_code == 502 and "secret" not in r.text


def test_refuses_to_start_without_app_key_or_credentials():
    with pytest.raises(RuntimeError):
        stt_proxy.Config({"SPELL_GOOGLE_API_KEY": "g"}).check()
    with pytest.raises(RuntimeError):
        stt_proxy.Config({"SPELL_APP_KEY": "k"}).check()
    with pytest.raises(RuntimeError):
        stt_proxy.Config({"SPELL_APP_KEY": "k", "GOOGLE_APPLICATION_CREDENTIALS": "/x"}).check()


def test_health_reports_backend():
    c, _ = make_client()
    assert c.get("/health").json() == {"ok": True, "backend": "google-v1"}


def test_v1_sends_key_in_header_not_url(monkeypatch):
    seen = {}

    class R:
        status_code = 200
        def json(self):
            return {"results": [{"alternatives": [{"transcript": " hei "}]}]}

    class FakeClient:
        async def post(self, url, json, headers):
            seen.update(url=url, headers=headers, json=json)
            return R()

    monkeypatch.setattr(stt_proxy, "http", lambda: FakeClient())
    assert asyncio.run(stt_proxy.recognize_v1("APIKEY", b"\x00\x00" * 10, "nb-NO")) == "hei"
    assert "APIKEY" not in seen["url"] and seen["headers"]["X-Goog-Api-Key"] == "APIKEY"
    assert seen["json"]["config"]["languageCode"] == "nb-NO"


def test_requests_run_concurrently_not_one_at_a_time():
    """50 requests that each wait 1 s on 'Google' must finish in ~1 s, not 50 s."""
    import httpx as hx

    async def slow(w, pcm, lang):
        await asyncio.sleep(1.0)
        return "ok"

    cfg = stt_proxy.Config({"SPELL_APP_KEY": "k1", "SPELL_GOOGLE_API_KEY": "g", "SPELL_RATE_PER_MIN": "1000"})
    app = stt_proxy.create_app(cfg, slow)

    async def run():
        transport = hx.ASGITransport(app=app)
        async with hx.AsyncClient(transport=transport, base_url="http://t") as c:
            start = time.monotonic()
            rs = await asyncio.gather(*[
                c.post("/transcribe", content=wav(), headers={"X-Spell-Key": "k1", "X-Spell-Lang": "nb"})
                for _ in range(50)])
            return time.monotonic() - start, [r.status_code for r in rs]

    elapsed, codes = asyncio.run(run())
    assert codes == [200] * 50
    assert elapsed < 5, f"took {elapsed:.1f}s: requests are being serialised"


def test_overload_returns_503_instead_of_queueing():
    import httpx as hx

    async def slow(w, pcm, lang):
        await asyncio.sleep(0.5)
        return "ok"

    cfg = stt_proxy.Config({"SPELL_APP_KEY": "k1", "SPELL_GOOGLE_API_KEY": "g",
                            "SPELL_RATE_PER_MIN": "1000", "SPELL_MAX_CONCURRENT": "5"})
    app = stt_proxy.create_app(cfg, slow)

    async def run():
        transport = hx.ASGITransport(app=app)
        async with hx.AsyncClient(transport=transport, base_url="http://t") as c:
            rs = await asyncio.gather(*[
                c.post("/transcribe", content=wav(), headers={"X-Spell-Key": "k1"})
                for _ in range(20)])
            return [r.status_code for r in rs]

    codes = asyncio.run(run())
    assert codes.count(200) >= 5 and 503 in codes and set(codes) <= {200, 503}
