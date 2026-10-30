# @mydoc/agent-core

Platform-independent `runAgent` with injected provider, tool gateway, event sink
and clock. A run streams complete calls, records the assistant message before
execution, executes bounded contiguous read batches with write barriers and feeds paired results
back to the model. The default read concurrency is four. Admission, including human
waits, runs before acquiring execution slots. An interaction answer must pass a
second authoritative gateway admission; it is never itself a permission grant.

The provider adapter owns incremental wire JSON parsing and yields `tool` only
when arguments are complete, then `done` after a valid response. Missing completion
never executes the batch. Duplicate call IDs are rejected across the run.

Cancellation/deadlines race uncooperative promises and forward AbortSignal. A tool
already started may have committed: its outcome is recorded as unconfirmed, and
future recovery must inspect the domain operation receipt. Cancellation cannot
undo external effects or kill arbitrary adapter code. Gateways must fence current
run authority immediately before commits. Journal failure rejects the run without
attempting more recovery writes; a durable host must reconcile interrupted runs.

Limits enforce rounds, wall time, output characters, call count, input size and
model-visible result characters. ContextManager preserves the original user goal,
keeps assistant/tool exchanges intact and projects summaries without rewriting the
journal. The default estimate counts serialized UTF-8 bytes plus fixed overhead;
it is a conservative approximation, not measured provider token usage. Hosts can
inject model-specific estimation and a bounded summarizer. Summary calls are
cancellable and their input/output size is capped. Hard-limit overflow stops before
invoking the provider. Pagination references, policy/source-aware prompt assembly,
provider SDKs and durable recovery remain follow-up work.

No database, shell, network implementation or UI is included. Event seq is local
to a run; a durable session event log must assign its own session cursor in A1.


## Workspace and source context

Optional `sourceContext` is a trusted host-owned ContextAccessPort. It supplies
workspace transfer policy, allowed tools and bounded excerpts, and authorizes each
source dependency before requests (including after asynchronous compaction).
Policies become system messages; excerpts are labeled untrusted user data with
stable per-run citation IDs and source references. The entire prefix counts toward
the input budget. Sources from other workspaces, policy revision changes or any
revoked dependency stop the run, protecting references retained in summaries.
The host must update policy revision for every policy change and validate source
hashes/revisions when constructing snapshots. This port is not yet wired to
CloudServer or SQLite; gateway authorization remains mandatory for tool execution.


A0 acceptance now includes lifecycle validators and recovery generation rules,
reported token usage budgets, and bounded result paging. RunResult.resultPages
retains run-local full results behind the continuation reference in truncated tool
messages. Inject ResultPages for host-owned storage. Missing vendor usage is
explicit via usageReports; default limits remain guardrails rather than measured
model throughput. See agent-host/DECISIONS.md for scope and dependency decisions.
