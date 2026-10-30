# @mydoc/agent-host

Initial in-memory admission facade. `start` allows one active run per session;
`cancel` matches both session and current run ID. Completion/failure releases the
session slot. Callers supply trusted identities and configured ports.

This is not an authenticated service or a persistent scheduler. A1 will add account
scope, durable queues, leases, subprocess RPC, replay, storage and shutdown guards.
Do not expose this facade directly to untrusted HTTP/IPC clients.


## Compatible model adapter

`CompatibleProvider` accepts an exact configured chat-completions endpoint, model,
maxOutputTokens and optional apiKey. It uses host-native fetch; no SDK, retry,
redirect or fallback is added. HTTP is restricted to explicit loopback addresses;
HTTPS endpoints must come from trusted host configuration, never tool arguments.
The host still owns endpoint allowlists, secrets storage and workspace policy.

The adapter maps tool schemas, assistant calls and results, parses bounded SSE
streams including UTF-8/CRLF fragmentation, and releases complete calls only after
an accepted finish reason and [DONE]. Truncation, length limits, invalid JSON,
multiple choices and unsupported finish reasons fail closed. Vendor HTTP error
bodies are not forwarded. Summary messages map to labeled user data, not system
instructions. Validated usage and explicit tool/window/output capabilities are supported. Missing
usage is represented by zero usageReports, not measured zero consumption. Additional
provider dialects are not included.

Tests use a real local HTTP server and synthetic model chunks. This is transport
verification, not certification of an external supplier or local model product.


## SQLite adapter (A1 foundation)

Import `SQLiteAgentStore` from `@mydoc/agent-host/sqlite-store`. Inject an already
checked SQLiteProjectDatabase and trusted account scope. Migration 5 owns DDL;
there is no request-time table creation. Claim a queued run then inject
store.events(lease) into runAgent. The adapter commits journal/message/call/result
and status together, exposes session cursors, fences old database/run generations
and isolates account scope. It is not yet wired into the production HTTP/IPC host.
Only the authoritative runtime owner calls recoverActive after worker loss;
opening another connection must not interrupt a live run. Persistent quotas,
compaction/cleanup, process coordination and authenticated APIs remain follow-up.


## Durable controlled runtime

`./durable-host` exports DurableAgentHost: enqueue persists first, runNext claims
one queued run and journals via SQLiteAgentStore. configure receives the lease;
domain commit services must validate that lease immediately before their writes.
Provider/tool callbacks are fenced before and after awaited operations. recover is
an explicit exclusive-owner operation, not automatic on each database connection.
close stops admission, aborts active runs and waits for their completion; queued
work remains persisted. Session cursor replay is supplied by the store.

`./child-runtime` exports ChildAgentRuntime. Supply an application-owned bundled
worker-child.ts entry (esbuild Node CJS); never accept a path from a remote caller.
The worker runs the portable loop, while pull-based model streaming, tools,
admission, interactions, sources, compaction, result pages and event writes use
host RPC. Model keys and database handles stay in the host. The child inherits no
ambient environment except NODE_ENV, and no parent execArgv. This is trusted-code
fault isolation, not an OS security sandbox. A cancellation grace deadline kills
unresponsive workers; an independent watchdog also bounds a hung startup/runtime.
A custom synchronous estimator cannot cross IPC and is explicitly rejected.

Worker exit/failure is journaled as interrupted, unresolved tool calls receive
unconfirmed-result messages, and generation advances; writes are never blindly
retried. This module is integrated with the persistent store and tested with real
Node processes, but production Web/Electron process assembly and authenticated
HTTP/IPC endpoints are still pending. No database migration is added here.
