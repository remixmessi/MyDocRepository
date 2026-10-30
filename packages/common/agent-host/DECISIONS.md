# A0 runtime and model decisions (2026-10-28)

Use the small MyDoc loop and host-native fetch, with no additional third-party
runtime dependency. The existing repository Yarn lock remains authoritative;
there is no implicit latest SDK resolution, copied source or upstream patch.

pi-ai/pi-agent-core are candidates, not adopted dependencies. The inspected
PI-Desktop version was 1.3.0 with application-specific patches; OpenClaw's core
was private and depended on its package graph. Importing those applications would
add unrelated contracts and maintenance obligations. Their license compatibility
is not asserted and no source is redistributed. If adoption is reconsidered,
inspect the exact release/license/dependencies/patches before installing it.

The selected dependency set adds no third-party license obligations. Portable
contracts/core bundle without Node dependencies. Host owns fetch, configured
endpoint, credentials, capabilities and cancellation. Native SSE tests cover
fragmented JSON/UTF-8/CRLF, completion, truncation, redacted HTTP errors and usage.
Capabilities are explicit configuration (default text/tools, 128000 context), not
inferred from model names; unsupported tools are rejected before networking.

Usage is accepted only as validated nonnegative token totals, once per request.
Run usageReports distinguishes missing vendor usage from a measured zero; sums
are only for reporting requests. The core total-token limit stops before tools
when a response exceeds it; it cannot reclaim tokens already consumed remotely.
Round/deadline/character/context limits remain effective without vendor usage.

Budget defaults are engineering guardrails, not measured throughput: 20 rounds,
10 minutes, 4 concurrent reads, 8000 result characters, 100000 prompt/output
characters, 128000 cumulative reported tokens. Context uses a conservative UTF-8
byte estimate plus overhead unless a model-specific estimator is injected;
compaction triggers at 80%, is bounded to 128000 input / 8000 summary characters,
and retains complete tool exchanges. Result pages retain at most 1000000
characters per run and expose at most 8000 per page. Hosts must choose actual
input headroom from the model window/output allowance before A1 deployment.

A0 acceptance is local protocol/kernel verification, not supplier/platform
certification. A1 owns persistent queues, real providers, secrets, authenticated
HTTP/IPC, resource authorization and process recovery. A0 exposes RunStore/IdPort,
recovery generation fencing and result paging contracts without opening SQLite.
