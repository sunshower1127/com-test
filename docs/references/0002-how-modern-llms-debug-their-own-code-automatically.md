# How modern LLMs debug their own code automatically

**Every major LLM provider has converged on the same core architecture: a server-side tool-use loop where the model generates code, a sandboxed runtime executes it, and the full error output feeds back to the model for autonomous retry.** The differences lie in sandbox technology, retry limits, and context engineering. Claude allows **10 iterations** per API call (extendable via `pause_turn`), Gemini caps at **5 retries**, and OpenAI publishes no hard limit. Academic research confirms this loop works specifically because code execution provides *external feedback* — a critical distinction from pure self-correction, which degrades model performance. For a developer building an Electron app, the most practical path is Pyodide (WebAssembly) for in-process sandboxing combined with a structured generate→execute→reflect→retry loop capped at 3–5 iterations.

---

## The universal loop: generate, execute, observe, retry

All three major platforms — Claude, Gemini, and ChatGPT — implement what is essentially the same architecture. The model calls a "code execution" tool, a sandboxed environment runs the code, and structured results (stdout, stderr, exit code) return to the model as context. The model then decides: produce a final answer, or write corrected code and execute again. This entire cycle happens **server-side within a single API call**, requiring no developer intervention for the retry logic.

The pattern expressed in pseudocode is remarkably simple:

```
while retries < max_retries:
    code = LLM.generate(task, previous_errors)
    result = sandbox.execute(code)
    if result.exit_code == 0:
        return result.stdout
    retries += 1
    previous_errors.append(result.stderr)
return failure
```

What differentiates production implementations from this skeleton is the richness of error context (full tracebacks vs. binary pass/fail), the sophistication of reflection prompts, and safety rails like repeated-error detection and no-op patch rejection. Academic research consistently shows that **richer feedback produces better repairs** — LDB's block-level variable tracking achieves **98.2%** on HumanEval versus simpler approaches in the low 80s.

---

## Claude, Gemini, and OpenAI: three approaches to the same problem

### Claude's code execution (Anthropic)

Claude runs code in **Ubuntu 24.04 Linux containers** with 1 CPU, 5 GiB RAM, and 5 GiB disk. The sandbox is completely network-isolated in the API version. When code fails, Claude receives the full `stdout`, `stderr`, and `return_code` as structured content blocks. The server-side sampling loop permits **up to 10 tool-call iterations per API request**. When that limit is reached, the API returns `stop_reason: "pause_turn"`, and the developer can send the response back to continue — effectively allowing unlimited iteration under developer control.

Claude's sub-tools include `bash_code_execution` (shell commands) and `text_editor_code_execution` (file operations). Pre-installed libraries cover the standard data science stack: pandas, numpy, scipy, scikit-learn, matplotlib, seaborn, plus extensive file-processing packages. Containers persist for **30 days** and can be reused across API calls by passing the container ID, maintaining file state between requests. A more advanced feature — **programmatic tool calling** — lets Claude write Python scripts that orchestrate external tool calls from within the sandbox, pausing execution when external results are needed.

### Gemini's code execution (Google)

Gemini uses **Google's gVisor** (a user-space kernel intercepting system calls) in Google's Runtime Environment. The key constraints are tighter: **30 seconds maximum** per execution and a hard cap of **5 automatic retries** per API call. Error results come back as structured `codeExecutionResult` parts with an explicit `outcome` field (`OUTCOME_OK` or `OUTCOME_FAILED`) plus the full output including tracebacks.

All retry iterations happen within a single `generateContent` response. The model interleaves `text`, `executableCode`, and `codeExecutionResult` parts — you can see the entire debugging chain in one API response. Gemini supports Python only for execution, charges no additional fee beyond standard token billing (unlike OpenAI and Anthropic), and does not expose a persistent container concept to developers.

### OpenAI's Code Interpreter

OpenAI's sandbox runs in a **fully sandboxed virtual machine** with configurable memory tiers from **1 GB to 64 GB**. Containers expire after **20 minutes of inactivity** but maintain full state (variables, files, imported modules) within a session. OpenAI's documentation explicitly states the model "can keep rewriting and running that code until it succeeds," but publishes **no hard retry limit** — the model autonomously decides when to stop.

The Responses API returns `code_interpreter_call` items capturing the full execution trace. Generated files persist in the container at `/mnt/data/` and can be downloaded via the container files API. OpenAI charges per container tier beyond standard token costs.

| Feature | Claude | Gemini | OpenAI |
|---|---|---|---|
| **Sandbox** | Linux container (Ubuntu) | gVisor in GRTE | Sandboxed VM |
| **Retry limit** | 10 per call (extendable) | 5 per call (hard) | No published limit |
| **Execution timeout** | 5 min container billing | 30 seconds | ~120 seconds |
| **Memory** | 5 GiB | Not configurable | 1–64 GB tiers |
| **Container persistence** | 30 days | None (ephemeral) | 20-min idle timeout |
| **Error context returned** | stdout + stderr + return_code | outcome enum + output | stdout + stderr + exit code |
| **Network access** | None (API) / limited allowlist (claude.ai) | None | None |
| **Extra cost** | $0.05/hr after free tier | None | Per-session/tier |

---

## What the research papers reveal about self-debugging

The academic literature on LLM self-debugging has matured rapidly since 2023, with several landmark papers establishing both the promise and limits of the approach.

**Self-Debugging** (Chen et al., ICLR 2024) introduced three feedback strategies. "Simple feedback" (binary pass/fail) helps somewhat. "Unit test feedback" (execution results with error messages) helps more. But the most interesting variant — **"code explanation" feedback** — has the model explain its own code line-by-line in natural language, a "rubber duck debugging" approach that works *without any execution feedback at all*. On complex SQL queries, Self-Debugging improved accuracy by **9 percentage points** and matched baselines generating 10× more candidate programs.

**Reflexion** (Shinn et al., NeurIPS 2023) added **episodic memory**. After each failed attempt, the model generates a verbal reflection ("I made an off-by-one error in the loop boundary because...") stored in a memory buffer. Subsequent attempts include these reflections as context. This approach achieved **91% pass@1 on HumanEval** — surpassing GPT-4's 80% at the time. The self-reflection component alone provided an **8% absolute boost** over memory without reflection.

**LDB** (Zhong et al., ACL 2024) pushed further by emulating human debugger behavior: segmenting programs into basic blocks via control flow graphs, tracking **intermediate variable values** at each block, and asking the LLM to verify block-by-block correctness. This fine-grained approach reached **98.2% on HumanEval** with GPT-4o seeds.

**Self-Refine** (Madaan et al., NeurIPS 2023) demonstrated a pure self-feedback loop using a single LLM as generator, critic, and refiner. Most gains came in the **first 2–3 iterations**, with diminishing returns after that — a finding consistent across nearly all papers and directly informing the 3–5 retry limits used in production systems.

Two critical counterpoints deserve attention. Olausson et al. (ICLR 2024) showed that when you account for the **computational cost** of repair attempts, the gains are "often modest, highly variable, and sometimes absent." And Huang et al. (ICLR 2024) demonstrated that **LLMs cannot self-correct reasoning without external feedback** — performance sometimes *degrades*. Code debugging works precisely because execution provides that external signal. This is the single most important insight from the literature: **the execution environment is not optional — it is the mechanism that makes self-correction viable**.

---

## Open-source tools and how they implement the loop

The open-source ecosystem offers several mature implementations of the auto-debugging pattern, each with distinct architectural choices.

**SWE-Agent** (Princeton) wraps LLM agents in Docker containers with shell access. Its successor, **mini-SWE-agent**, achieves over 74% on SWE-bench Verified in roughly 100 lines of code. The architecture is pure simplicity: the only tool is bash, each command runs via `subprocess.run`, and the entire history is a linear message list. Guardrails include automatic linting after edits — changes only apply if they don't introduce syntax errors.

**OpenHands** (65k+ GitHub stars) uses an **event-stream abstraction** where actions and observations form a perception-action loop. Each agent gets its own Docker container with an integrated Jupyter kernel for stateful Python execution. The **CodeAct agent** writes and executes code (Python/bash), observes results including full stdout/stderr, and iterates. Sub-agent delegation allows specialized collaborators to handle subtasks.

**Aider** (39k+ stars) implements an elegant **auto-lint + auto-test loop**. After every LLM edit: (1) run linter on edited files, (2) if lint errors found, feed back to LLM, (3) run test suite, (4) if tests fail, feed output back to LLM. Tree-sitter AST-aware context helps the model understand code structure. This lint-first approach catches many errors cheaply before full execution.

**LangGraph** provides the most developer-friendly framework for building custom loops. A three-node graph (Generate → Execute → Reflect) with conditional routing handles the retry logic declaratively. Built-in `RetryPolicy` supports configurable attempts and conditions. The `ToolNode` with `handle_tool_errors=True` automatically sends raw error text back to the LLM for self-correction.

**E2B** offers the sandbox primitive: Firecracker microVMs with ~150ms cold start, stateful Jupyter sessions, and SDKs for Python and TypeScript. E2B deliberately does *not* implement retry logic — it provides the execution layer, leaving the debugging loop to the developer. This separation of concerns makes it composable with any framework.

---

## Building this in Electron: practical architecture

For a Korean developer building an Electron-based app, the most practical sandbox strategy is **Pyodide (CPython compiled to WebAssembly)**. Pyodide runs in the renderer process with browser-grade isolation — no filesystem access, no network, no system calls. Cached Pyodide instances start in under **100ms** after initial load (which takes 2–3 seconds). Performance is **10–30% slower** than native Python, acceptable for most code execution tasks.

The recommended Electron architecture has three layers:

The **renderer process** (UI) sends code execution requests via Electron IPC to the **main process**, which manages a Pyodide sandbox instance. The sandbox executes code with `pyodide.runPythonAsync()`, captures results, and returns structured output (stdout, stderr, error) via IPC back to the renderer. The LLM integration (API calls to Claude, Gemini, or OpenAI) runs in the main process or a utility process.

For the self-debugging loop itself, implement these safety rails drawn from production systems: a **retry ceiling of 3–5 attempts** (matching the diminishing-returns finding from Self-Refine), **repeated-error detection** (stop if the same error appears twice consecutively), **no-op patch rejection** (stop if the LLM returns identical code), and a **token budget cap** as a hard ceiling. Feed the LLM the original task, the failing code, the full stderr/traceback, and summaries of previous attempts — this context engineering is where most of the quality differentiation happens.

For cases where Pyodide's library limitations are a problem (no C extensions, limited package set), fall back to **E2B's cloud sandboxes** ($0.05/hour per VM) or **Docker containers** spawned from the main process. The LangChain Sandbox library (now deprecated but architecturally instructive) demonstrated exactly this Pyodide-in-Deno pattern, and the source code remains a useful reference.

Korean developer resources on Velog and WikiDocs document the same canonical pattern — "생성 → 실행 → 오류 확인 → 수정" (Generate → Execute → Check Error → Fix) — with practical notes on how statically-typed languages produce compiler errors that LLMs can fix faster, and how logging corrections to a growing guideline file (similar to Aider's approach) improves results over time.

---

## Conclusion: what actually matters

The technical details across platforms and papers converge on a few non-obvious insights. First, **execution feedback is the mechanism, not the model's self-knowledge** — without running the code and feeding back real errors, self-correction often fails or degrades quality. Second, **most value comes from the first 2–3 retries**; going beyond 5 iterations rarely helps and often wastes tokens. Third, **the quality of error context matters more than the number of retries** — LDB's block-level variable tracking vastly outperforms simple pass/fail signals. Fourth, **the sandbox is a security boundary, not just a convenience** — the industry is moving from Docker (shared kernel) to Firecracker microVMs or gVisor for untrusted AI-generated code.

For an Electron implementation, the practical architecture is: Pyodide for local sandboxing, a structured 3–5 iteration loop with rich error context, and cloud sandbox fallback (E2B) for heavy workloads. The pattern is simple enough to implement in a weekend, but the context engineering — what exactly you feed back to the model on each failure — is what separates a demo from a production system.