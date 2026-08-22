"use strict";
(() => {
  var __getOwnPropNames = Object.getOwnPropertyNames;
  var __esm = (fn, res, err) => function __init() {
    if (err) throw err[0];
    try {
      return fn && (res = (0, fn[__getOwnPropNames(fn)[0]])(fn = 0)), res;
    } catch (e) {
      throw err = [e], e;
    }
  };
  var __commonJS = (cb, mod) => function __require() {
    try {
      return mod || (0, cb[__getOwnPropNames(cb)[0]])((mod = { exports: {} }).exports, mod), mod.exports;
    } catch (e) {
      throw mod = 0, e;
    }
  };

  // web-src/agent-core/types.ts
  var init_types = __esm({
    "web-src/agent-core/types.ts"() {
      "use strict";
    }
  });

  // web-src/agent-core/skill.ts
  var init_skill = __esm({
    "web-src/agent-core/skill.ts"() {
      "use strict";
    }
  });

  // web-src/agent-core/loop.ts
  function utf8Size(s) {
    let n = 0;
    for (let i = 0; i < s.length; i++) {
      const c = s.charCodeAt(i);
      n += c < 128 ? 1 : c < 2048 ? 2 : 3;
    }
    return n;
  }
  function messageSize(m) {
    if (m.role === "tool") {
      return m.results.reduce((n2, r) => n2 + utf8Size(r.output) + 40, 0);
    }
    let n = utf8Size(m.text);
    if (m.role === "user" && m.images) {
      n += m.images.reduce((s, img) => s + img.base64.length, 0);
    }
    if (m.role === "assistant" && m.toolCalls) {
      for (const c of m.toolCalls) {
        try {
          n += utf8Size(JSON.stringify(c.input)) + 40;
        } catch {
          n += 40;
        }
      }
    }
    return n;
  }
  function historySize(messages) {
    return messages.reduce((n, m) => n + messageSize(m), 0);
  }
  function mechanicalDigest(dropped) {
    const lines = [];
    for (const m of dropped) {
      if (m.role === "user" && !m.text.startsWith(COMPACT_SUMMARY_PREFIX)) {
        lines.push(`- User: ${m.text.slice(0, 200)}`);
      } else if (m.role === "assistant" && m.text && !m.toolCalls?.length) {
        lines.push(`  Reply: ${m.text.slice(0, 200)}`);
      }
    }
    return lines.join("\n").slice(0, 4e3) || "(earlier conversation omitted)";
  }
  function sanitizeAgentPayload(payload) {
    return payload.replace(/\b(?:sk-|AIza|ghp_|secret_)[A-Za-z0-9_-]{16,}/g, "[REDACTED_API_KEY]").replace(/([a-z][a-z0-9+.-]*:\/\/[^\s:@/]+):[^\s@/]+@/gi, "$1:[REDACTED_CREDENTIALS]@").replace(
      /(password|passwd|secret_key|private_key)(\s*[:=]\s*)["'][^"']+["']/gi,
      '$1$2"[REDACTED_SECURE_TOKEN]"'
    );
  }
  var COMPACT_MAX_BYTES, COMPACT_KEEP_RECENT_BYTES, SUMMARIZE_TOOL_OUTPUT_MAX, SUMMARIZE_TIMEOUT_MS, STALE_TOOL_KEEP_RECENT, STALE_TOOL_OUTPUT_MAX, MAX_INPUT_PARSE_RETRIES, TURN_LIMIT_NOTE, COMPLETED_VIA_TOOLS_TEXT, SUMMARIZE_SYSTEM, COMPACT_SUMMARY_PREFIX, COMPACT_SUMMARY_HEADER, COMPACT_SUMMARY_ACK, AgentLoop;
  var init_loop = __esm({
    "web-src/agent-core/loop.ts"() {
      "use strict";
      COMPACT_MAX_BYTES = 256 * 1024;
      COMPACT_KEEP_RECENT_BYTES = 96 * 1024;
      SUMMARIZE_TOOL_OUTPUT_MAX = 2e3;
      SUMMARIZE_TIMEOUT_MS = 3e4;
      STALE_TOOL_KEEP_RECENT = 2;
      STALE_TOOL_OUTPUT_MAX = 1e3;
      MAX_INPUT_PARSE_RETRIES = 3;
      TURN_LIMIT_NOTE = "[System] The tool-call turn limit for this request has been reached; no more tools may be called this turn. Answer directly from the information already gathered; if the task is unfinished, briefly state what is done and what remains.";
      COMPLETED_VIA_TOOLS_TEXT = "(completed tool actions; no text reply)";
      SUMMARIZE_SYSTEM = `You are a conversation compressor. Compress this editing session between the user and the AI assistant into a concise summary so later turns can continue with context. Keep: the user's goals and key instructions, completed changes (which files/pages/elements were modified), important facts and data, and outstanding items. For specific figures/statistics, mark their provenance: figures from the user or from tool results (e.g. web_search) keep their source; figures the assistant produced without a source must be marked "(unverified)" so later turns do not treat them as established facts. Omit: pleasantries, tool-call details, and intermediate trial and error. Use a bullet list of at most 400 words. Write the summary in the same language as the conversation. Output only the summary body, with no preamble.`;
      COMPACT_SUMMARY_PREFIX = "[Summary of earlier conversation";
      COMPACT_SUMMARY_HEADER = "[Summary of earlier conversation (auto-compacted)]";
      COMPACT_SUMMARY_ACK = "Understood, continuing from the progress so far.";
      AgentLoop = class {
        options;
        history = [];
        handle = null;
        running = false;
        cancelled = false;
        turns = 0;
        /** Finalizing turn after hitting the turn limit: no tools, let the model answer from what it has read */
        finalizing = false;
        mutationSeen = false;
        inputParseFails = 0;
        turnStopReason = null;
        turnText = "";
        toolCalls = [];
        /** user message of the in-flight run; a failed run rolls it (and everything after) back out of history */
        runUserMsg = null;
        /** invalidates stale transport callbacks after cancel/reset */
        generation = 0;
        /** per-run abort: aborted on cancel(); long tools (e.g. generate_deck) use it to break internal loops */
        abortController = null;
        constructor(options) {
          this.options = options;
        }
        get busy() {
          return this.running;
        }
        get messages() {
          return this.history;
        }
        /**
         * Seed the conversation with restored history (e.g. transcript reloaded from
         * disk when a document reopens), so follow-up instructions keep their context.
         * No-op unless the loop is idle with an empty history.
         * Old messages over the compaction budget fold into a mechanical digest
         * (no LLM request on restore, guaranteeing zero latency).
         */
        restore(messages) {
          if (this.running || this.history.length > 0 || messages.length === 0) return;
          const normalized = messages.map(
            (m) => m.role === "assistant" && !m.text ? { ...m, text: COMPLETED_VIA_TOOLS_TEXT } : m
          );
          this.history = normalized.filter(
            (m, i) => m.role !== "user" || normalized[i + 1] && normalized[i + 1].role !== "user"
          );
          if (this.history.length === 0) return;
          if (this.compactionEnabled()) {
            const { maxBytes, keepRecentBytes } = this.compactBudget();
            if (historySize(this.history) > maxBytes) {
              const cut = this.findCompactCut(keepRecentBytes);
              if (cut > 0) {
                const digest = mechanicalDigest(this.history.slice(0, cut));
                this.history = [
                  { role: "user", text: `${COMPACT_SUMMARY_HEADER}
${digest}` },
                  { role: "assistant", text: COMPACT_SUMMARY_ACK },
                  ...this.history.slice(cut)
                ];
              }
            }
          }
          this.trimHistory();
        }
        /** images: inline attachments for this user turn (vision input; see AgentImage) */
        run(instruction, images) {
          if (this.running || !instruction) return;
          this.running = true;
          this.cancelled = false;
          this.turns = 0;
          this.finalizing = false;
          this.mutationSeen = false;
          this.inputParseFails = 0;
          this.abortController = new AbortController();
          const context = this.options.skill.buildContext?.() ?? "";
          const format = this.options.formatUserMessage ?? ((instr, ctx) => ctx ? `${instr}

${ctx}` : instr);
          const userMsg = {
            role: "user",
            text: format(instruction, context),
            ...images?.length ? { images } : {}
          };
          void this.beginRun(userMsg);
        }
        /** Compact (if needed), push the user message, then start the turn. Compaction failure doesn't block the run. */
        async beginRun(userMsg) {
          const generation = this.generation;
          try {
            await this.maybeCompact();
          } catch {
          }
          if (generation !== this.generation) return;
          if (this.cancelled) {
            this.running = false;
            this.options.events?.onDone?.({ text: "", cancelled: true, turnLimit: false });
            return;
          }
          while (this.history.at(-1)?.role === "user") this.history.pop();
          this.trimHistory();
          if (userMsg.role === "user") {
            userMsg = { ...userMsg, text: sanitizeAgentPayload(userMsg.text) };
          }
          this.runUserMsg = userMsg;
          this.history.push(userMsg);
          this.startTurn();
        }
        /**
         * A run failed: remove its user message and every message after it, so the
         * failed instruction can't be silently re-executed by the next run.
         */
        rollbackFailedRun() {
          const msg = this.runUserMsg;
          this.runUserMsg = null;
          if (!msg) return;
          const i = this.history.lastIndexOf(msg);
          if (i >= 0) this.history.splice(i);
        }
        // ── Context compaction: fold old conversation into a summary, keep recent messages verbatim ──
        compactionEnabled() {
          return this.options.compaction !== false;
        }
        compactBudget() {
          const opt = this.options.compaction === false ? void 0 : this.options.compaction;
          return {
            maxBytes: opt?.maxBytes ?? COMPACT_MAX_BYTES,
            keepRecentBytes: opt?.keepRecentBytes ?? COMPACT_KEEP_RECENT_BYTES
          };
        }
        /**
         * Find the compaction cut at a user boundary: accumulate from the tail up to keepRecentBytes.
         * Returns the start index of the kept segment; if no suitable boundary exists,
         * fall back to keeping the last user turn.
         */
        findCompactCut(keepRecentBytes) {
          let kept = 0;
          let cut = -1;
          for (let i = this.history.length - 1; i >= 0; i--) {
            kept += messageSize(this.history[i]);
            if (kept > keepRecentBytes && cut >= 0) break;
            if (this.history[i].role === "user") cut = i;
          }
          if (cut < 0) {
            for (let i = this.history.length - 1; i >= 0; i--) {
              if (this.history[i].role === "user") return i;
            }
          }
          return cut;
        }
        async maybeCompact() {
          if (!this.compactionEnabled()) return;
          const { maxBytes, keepRecentBytes } = this.compactBudget();
          if (historySize(this.history) <= maxBytes) return;
          const cut = this.findCompactCut(keepRecentBytes);
          if (cut <= 0) return;
          const dropped = this.history.slice(0, cut);
          const opt = this.options.compaction === false ? void 0 : this.options.compaction;
          let summary = null;
          if (!opt?.disableLlmSummary) summary = await this.summarizeViaLlm(dropped);
          if (!summary) summary = mechanicalDigest(dropped);
          this.history = [
            { role: "user", text: `${COMPACT_SUMMARY_HEADER}
${summary}` },
            { role: "assistant", text: COMPACT_SUMMARY_ACK },
            ...this.history.slice(cut)
          ];
        }
        /** Hand the folded conversation to the model for a summary; returns null on failure/timeout (falls back to the mechanical digest). */
        summarizeViaLlm(dropped) {
          const slim = dropped.map((m) => {
            if (m.role === "tool") {
              return {
                role: "tool",
                results: m.results.map((r) => ({
                  ...r,
                  output: r.output.slice(0, SUMMARIZE_TOOL_OUTPUT_MAX)
                }))
              };
            }
            if (m.role === "user" && m.images?.length) return { role: "user", text: m.text };
            return m;
          });
          return new Promise((resolve) => {
            let text = "";
            let settled = false;
            const finish = (v) => {
              if (settled) return;
              settled = true;
              clearTimeout(timer);
              resolve(v);
            };
            const timer = setTimeout(() => finish(null), SUMMARIZE_TIMEOUT_MS);
            try {
              this.handle = this.options.transport.stream(
                {
                  system: SUMMARIZE_SYSTEM,
                  messages: [
                    ...slim,
                    { role: "user", text: "Compress the conversation above as instructed." }
                  ],
                  tools: []
                },
                {
                  onDelta: (t) => {
                    text += t;
                  },
                  onToolCall: () => {
                  },
                  onDone: () => finish(text.trim() || null),
                  onError: () => finish(null)
                }
              );
            } catch {
              finish(null);
            }
          });
        }
        /**
         * When over budget mid-run (between tool turns), truncate stale tool outputs:
         * keep structure (tool_use/tool_result pairs intact), cut content only,
         * and keep the most recent N verbatim.
         */
        squashStaleToolOutputs() {
          if (!this.compactionEnabled()) return;
          const { maxBytes } = this.compactBudget();
          if (historySize(this.history) <= maxBytes) return;
          let recent = 0;
          for (let i = this.history.length - 1; i >= 0; i--) {
            const m = this.history[i];
            if (m.role !== "tool") continue;
            recent++;
            if (recent <= STALE_TOOL_KEEP_RECENT) continue;
            m.results = m.results.map(
              (r) => r.output.length > STALE_TOOL_OUTPUT_MAX ? {
                ...r,
                output: `${r.output.slice(0, STALE_TOOL_OUTPUT_MAX)}
\u2026(output truncated: too long)`
              } : r
            );
          }
        }
        cancel() {
          if (!this.running) return;
          this.cancelled = true;
          this.abortController?.abort();
          this.handle?.cancel();
        }
        /** drop the conversation (e.g. when a different document is opened) */
        reset() {
          this.generation++;
          this.abortController?.abort();
          this.handle?.cancel();
          this.handle = null;
          this.running = false;
          this.cancelled = false;
          this.history = [];
          this.runUserMsg = null;
        }
        /** Runs at run boundaries only (restore / before a new user message): a long run's tail is all assistant/tool messages, and cutting mid-run would empty the request. */
        trimHistory() {
          const max = this.options.maxHistory ?? 40;
          if (this.history.length <= max) return;
          let i = this.history.length - max;
          while (i < this.history.length && this.history[i].role !== "user") i++;
          if (i >= this.history.length) return;
          const next = this.history.slice(i);
          if (this.runUserMsg && !next.includes(this.runUserMsg)) return;
          this.history = next;
        }
        startTurn() {
          const generation = this.generation;
          this.turnText = "";
          this.toolCalls = [];
          this.turnStopReason = null;
          let settled = false;
          this.handle = this.options.transport.stream(
            {
              system: this.options.skill.systemPrompt + (this.options.systemSuffix?.() ?? ""),
              messages: [...this.history],
              tools: this.finalizing ? [] : this.options.skill.tools
            },
            {
              onDelta: (text) => {
                if (generation !== this.generation || settled) return;
                this.turnText += text;
                this.options.events?.onText?.(this.turnText);
              },
              onToolCall: (call) => {
                if (generation !== this.generation || settled) return;
                this.toolCalls.push(call);
              },
              onStopReason: (reason) => {
                if (generation !== this.generation || settled) return;
                this.turnStopReason = reason;
              },
              onDone: () => {
                if (generation !== this.generation || settled) return;
                settled = true;
                void this.finishTurn();
              },
              onError: (error) => {
                if (generation !== this.generation || settled) return;
                settled = true;
                this.running = false;
                this.rollbackFailedRun();
                this.options.events?.onError?.(error);
              }
            }
          );
        }
        async finishTurn() {
          const { events, skill, captureSnapshot } = this.options;
          const toolCalls = this.toolCalls;
          if (toolCalls.length === 0 || this.cancelled || this.finalizing) {
            this.history.push({ role: "assistant", text: this.turnText || COMPLETED_VIA_TOOLS_TEXT });
            this.running = false;
            this.runUserMsg = null;
            events?.onDone?.({
              text: this.turnText,
              cancelled: this.cancelled,
              turnLimit: this.finalizing,
              // set only when true so exact-shape consumers/tests stay unaffected
              ...this.turnStopReason === "max_tokens" && !this.cancelled ? { truncated: true } : {}
            });
            return;
          }
          this.history.push({ role: "assistant", text: this.turnText, toolCalls });
          const generation = this.generation;
          const results = [];
          for (const call of toolCalls) {
            if (this.cancelled) {
              results.push({
                id: call.id,
                name: call.name,
                output: "(the user stopped the run; this tool was not executed)",
                isError: true
              });
              continue;
            }
            if (call.truncated || call.inputError) {
              this.inputParseFails++;
              const output = call.truncated ? "Tool arguments were cut off by the output length limit; the tool was not executed. Split this operation into several smaller tool calls (less content per call) and try again." : `Tool input JSON failed to parse; the tool was not executed: ${call.inputError}
Fix the arguments (make sure quotes inside strings are escaped) and call again.`;
              results.push({ id: call.id, name: call.name, output, isError: true });
              events?.onToolExecuted?.({
                call,
                execution: { output, isError: true, summary: call.name }
              });
              continue;
            }
            this.inputParseFails = 0;
            events?.onToolStart?.(call);
            const snapshot = !this.mutationSeen ? captureSnapshot?.() : void 0;
            let execution;
            try {
              execution = await skill.executeTool(call, this.abortController?.signal);
            } catch (e) {
              execution = {
                output: e instanceof Error ? e.message : String(e),
                isError: true,
                summary: call.name
              };
            }
            if (generation !== this.generation) return;
            const firstMutation = !!execution.mutated && !this.mutationSeen;
            if (execution.mutated) this.mutationSeen = true;
            results.push({
              id: call.id,
              name: call.name,
              output: execution.output,
              isError: execution.isError
            });
            events?.onToolExecuted?.({
              call,
              execution,
              snapshotBefore: firstMutation ? snapshot : void 0
            });
          }
          this.history.push({ role: "tool", results });
          if (this.cancelled) {
            this.running = false;
            this.runUserMsg = null;
            events?.onDone?.({ text: this.turnText, cancelled: true, turnLimit: false });
            return;
          }
          if (this.inputParseFails >= MAX_INPUT_PARSE_RETRIES) {
            this.running = false;
            this.rollbackFailedRun();
            events?.onError?.(
              `Tool input was unusable (unparseable or truncated) ${MAX_INPUT_PARSE_RETRIES} times in a row; retries stopped, please send the request again`
            );
            return;
          }
          this.turns++;
          if (this.turns >= (this.options.maxTurns ?? 8)) {
            this.finalizing = true;
            this.history.push({ role: "user", text: TURN_LIMIT_NOTE });
          }
          this.squashStaleToolOutputs();
          events?.onTurnEnd?.();
          this.startTurn();
        }
      };
    }
  });

  // web-src/agent-core/index.ts
  var init_agent_core = __esm({
    "web-src/agent-core/index.ts"() {
      "use strict";
      init_types();
      init_skill();
      init_loop();
    }
  });

  // web-src/ai-provider/types.ts
  var init_types2 = __esm({
    "web-src/ai-provider/types.ts"() {
      "use strict";
    }
  });

  // web-src/ai-provider/fetch.ts
  async function aiFetch(url, init) {
    try {
      return await fetch(url, init);
    } catch (primaryError) {
      const signal = init.signal;
      if (!rescueFetch || signal?.aborted) throw primaryError;
      console.warn("[ai-provider] fetch failed, retrying via rescue fetch:", String(primaryError));
      try {
        return await rescueFetch(url, init);
      } catch {
        throw primaryError;
      }
    }
  }
  var rescueFetch;
  var init_fetch = __esm({
    "web-src/ai-provider/fetch.ts"() {
      "use strict";
      rescueFetch = null;
    }
  });

  // web-src/ai-provider/http-error.ts
  function httpBodyDetail(body) {
    const head = body.trimStart().slice(0, 30).toLowerCase();
    const isHtml = ["<!doctype", "<html", "<head", "<body"].some((tag) => head.startsWith(tag));
    if (isHtml) {
      return "the service returned a web page instead of an API response (likely a temporary network or gateway block) \u2014 check your connection and retry";
    }
    return body.slice(0, 500);
  }
  var init_http_error = __esm({
    "web-src/ai-provider/http-error.ts"() {
      "use strict";
    }
  });

  // web-src/ai-provider/providers.ts
  function gensparkAttributionHeaders(baseUrl) {
    return baseUrl?.startsWith("https://www.genspark.ai") ? { "X-Agent-Type": GENSPARK_AGENT_TYPE } : {};
  }
  var GENSPARK_AGENT_TYPE;
  var init_providers = __esm({
    "web-src/ai-provider/providers.ts"() {
      "use strict";
      GENSPARK_AGENT_TYPE = "genoffice";
    }
  });

  // web-src/ai-provider/watchdog.ts
  function createStreamWatchdog(parent, connectMs = AI_CONNECT_TIMEOUT_MS, idleMs = AI_IDLE_TIMEOUT_MS) {
    const controller = new AbortController();
    let timedOutAfter = 0;
    let timer;
    const arm = (ms) => {
      clearTimeout(timer);
      timer = setTimeout(() => {
        timedOutAfter = ms;
        controller.abort();
      }, ms);
    };
    const onParentAbort = () => controller.abort();
    if (parent?.aborted) controller.abort();
    else parent?.addEventListener("abort", onParentAbort, { once: true });
    arm(connectMs);
    return {
      signal: controller.signal,
      touch: () => arm(idleMs),
      async guard(run) {
        try {
          return await run();
        } catch (e) {
          if (timedOutAfter > 0) throw new AiTimeoutError(timedOutAfter);
          throw e;
        } finally {
          clearTimeout(timer);
          parent?.removeEventListener("abort", onParentAbort);
        }
      }
    };
  }
  var AI_CONNECT_TIMEOUT_MS, AI_IDLE_TIMEOUT_MS, AiTimeoutError;
  var init_watchdog = __esm({
    "web-src/ai-provider/watchdog.ts"() {
      "use strict";
      AI_CONNECT_TIMEOUT_MS = 6e4;
      AI_IDLE_TIMEOUT_MS = 18e4;
      AiTimeoutError = class extends Error {
        constructor(ms) {
          super(`AI request timed out: no data received from the network for ${Math.round(ms / 1e3)}s`);
          this.name = "AiTimeoutError";
        }
      };
    }
  });

  // web-src/ai-provider/stream.ts
  async function* sseLines(body, onBytes) {
    const decoder = new TextDecoder();
    let buffer = "";
    const stream = body;
    const reader = stream.getReader();
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      onBytes?.();
      buffer += decoder.decode(value, { stream: true });
      const lines = buffer.split("\n");
      buffer = lines.pop() ?? "";
      for (const line of lines) yield line;
    }
    if (buffer) yield buffer;
  }
  function repairUnescapedQuotes(json) {
    let out = "";
    let inStr = false;
    for (let i = 0; i < json.length; i++) {
      const c = json[i];
      if (!inStr) {
        if (c === '"') inStr = true;
        out += c;
        continue;
      }
      if (c === "\\") {
        out += c + (json[++i] ?? "");
        continue;
      }
      if (c === '"') {
        let j = i + 1;
        while (j < json.length && " \n\r	".includes(json[j])) j++;
        const next = json[j];
        if (next === void 0 || ",}]:".includes(next)) {
          inStr = false;
          out += c;
        } else {
          out += '\\"';
        }
        continue;
      }
      out += c;
    }
    return out;
  }
  function sseErrorText(error, fallback) {
    if (typeof error === "string" && error) return error;
    if (error && typeof error === "object") {
      const message = error.message;
      if (typeof message === "string" && message) return message;
      try {
        return JSON.stringify(error);
      } catch {
      }
    }
    return fallback;
  }
  async function jsonBodyInsteadOfSse(response) {
    const contentType = response.headers.get("content-type") ?? "";
    return contentType.includes("application/json") ? await response.text() : null;
  }
  function creditsNoticeText(value) {
    if (typeof value === "string") {
      const t = value.toLowerCase();
      const credits = t.includes("genspark.ai/pricing") || t.includes("credit") && (t.includes("exhausted") || t.includes("insufficient"));
      return credits ? value : null;
    }
    if (Array.isArray(value) || value && typeof value === "object") {
      for (const v of Object.values(value)) {
        const hit = creditsNoticeText(v);
        if (hit) return hit;
      }
    }
    return null;
  }
  function throwIfCreditsNotice(bodyText) {
    let parsed;
    try {
      parsed = JSON.parse(bodyText);
    } catch {
      return;
    }
    const notice = creditsNoticeText(parsed);
    if (notice) throw new AiCreditsError(notice);
  }
  function parseToolInput(json) {
    if (!json.trim()) return { input: {} };
    try {
      return { input: JSON.parse(json) };
    } catch (e) {
      try {
        return { input: JSON.parse(repairUnescapedQuotes(json)) };
      } catch {
        const msg = e instanceof Error ? e.message : String(e);
        return { input: {}, error: `${msg}; raw: ${json.slice(0, 500)}` };
      }
    }
  }
  function openAiMessages(system, messages) {
    const out = [{ role: "system", content: system }];
    for (const m of messages) {
      if (m.role === "user") {
        if (!m.images?.length) {
          out.push({ role: "user", content: m.text });
        } else {
          out.push({
            role: "user",
            content: [
              ...m.text ? [{ type: "text", text: m.text }] : [],
              ...m.images.map((img) => ({
                type: "image_url",
                image_url: { url: `data:${img.mime};base64,${img.base64}` }
              }))
            ]
          });
        }
      } else if (m.role === "assistant") {
        const hasTools = !!(m.toolCalls && m.toolCalls.length > 0);
        out.push({
          role: "assistant",
          content: m.text || (hasTools ? null : "(no content)"),
          ...hasTools ? {
            tool_calls: m.toolCalls.map((call) => ({
              id: call.id,
              type: "function",
              function: { name: call.name, arguments: JSON.stringify(call.input) }
            }))
          } : {}
        });
      } else {
        for (const r of m.results) {
          out.push({ role: "tool", tool_call_id: r.id, content: r.output });
        }
      }
    }
    return out;
  }
  function emitOpenAiJsonMessage(bodyText, cb) {
    let msg;
    try {
      msg = JSON.parse(bodyText);
    } catch {
      throw new Error(`The model returned an unparseable JSON body: ${httpBodyDetail(bodyText)}`);
    }
    if (msg.error) throw new Error(sseErrorText(msg.error, "Model error"));
    const choice = msg.choices?.[0];
    let emitted = false;
    if (choice?.message?.content) {
      emitted = true;
      cb.onDelta(choice.message.content);
    }
    const toolCalls = [];
    for (const tc of choice?.message?.tool_calls ?? []) {
      if (!tc.function?.name) continue;
      emitted = true;
      const { input, error } = parseToolInput(tc.function.arguments ?? "");
      toolCalls.push({
        id: tc.id ?? crypto.randomUUID(),
        name: tc.function.name,
        input,
        inputError: error
      });
    }
    const lastTool = toolCalls.at(-1);
    if (choice?.finish_reason === "length" && lastTool) lastTool.truncated = true;
    for (const call of toolCalls) cb.onToolCall(call);
    if (!emitted) throw new Error(`The model returned no content: ${httpBodyDetail(bodyText)}`);
    if (choice?.finish_reason === "length") cb.onStopReason?.("max_tokens");
  }
  async function streamOpenAiCompatible(baseUrl, config, system, messages, tools, maxTokens, cb) {
    const wd = createStreamWatchdog(cb.signal);
    return wd.guard(
      () => openAiCompatibleTurn(baseUrl, config, system, messages, tools, maxTokens, cb, wd)
    );
  }
  async function openAiCompatibleTurn(baseUrl, config, system, messages, tools, maxTokens, cb, wd) {
    const onBytes = () => {
      wd.touch();
      cb.onActivity?.();
    };
    const response = await aiFetch(`${baseUrl.replace(/\/$/, "")}/chat/completions`, {
      method: "POST",
      signal: wd.signal,
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${config.apiKey}`,
        ...gensparkAttributionHeaders(baseUrl)
      },
      body: JSON.stringify({
        model: config.model,
        max_tokens: maxTokens,
        messages: openAiMessages(system, messages),
        ...tools.length > 0 ? {
          tools: tools.map((t) => ({
            type: "function",
            function: { name: t.name, description: t.description, parameters: t.inputSchema }
          }))
        } : {},
        temperature: 0.3,
        stream: true
      })
    });
    onBytes();
    if (!response.ok || !response.body) {
      throw new Error(`HTTP ${response.status}: ${httpBodyDetail(await response.text())}`);
    }
    const jsonBody = await jsonBodyInsteadOfSse(response);
    if (jsonBody !== null) {
      throwIfCreditsNotice(jsonBody);
      return emitOpenAiJsonMessage(jsonBody, cb);
    }
    const pendingTools = /* @__PURE__ */ new Map();
    let stopReason;
    let abnormalFinish;
    let sawFinish = false;
    let emitted = false;
    const flushTools = () => {
      const entries = [...pendingTools.entries()].sort(([a], [b]) => a - b);
      const lastIndex = entries.at(-1)?.[0];
      for (const [index, pending] of entries) {
        if (pending.name) {
          const { input, error } = parseToolInput(pending.json);
          emitted = true;
          cb.onToolCall({
            id: pending.id,
            name: pending.name,
            input,
            inputError: error,
            // a 'length' finish cuts off the last streaming tool's arguments
            ...stopReason === "max_tokens" && index === lastIndex ? { truncated: true } : {}
          });
        }
      }
      pendingTools.clear();
    };
    for await (const line of sseLines(response.body, onBytes)) {
      if (!line.startsWith("data:")) continue;
      const payload = line.slice(5).trim();
      if (!payload) continue;
      if (payload === "[DONE]") break;
      const event = JSON.parse(payload);
      if (event.error) throw new Error(sseErrorText(event.error, "Model stream error"));
      const choice = event.choices?.[0];
      if (!choice) continue;
      if (choice.delta?.content) {
        emitted = true;
        cb.onDelta(choice.delta.content);
      }
      for (const tc of choice.delta?.tool_calls ?? []) {
        const pending = pendingTools.get(tc.index) ?? {
          id: tc.id ?? crypto.randomUUID(),
          name: "",
          json: ""
        };
        if (tc.id) pending.id = tc.id;
        if (tc.function?.name) pending.name += tc.function.name;
        if (tc.function?.arguments) pending.json += tc.function.arguments;
        pendingTools.set(tc.index, pending);
      }
      if (choice.finish_reason) {
        sawFinish = true;
        if (choice.finish_reason === "length") stopReason = "max_tokens";
        else if (choice.finish_reason !== "stop" && choice.finish_reason !== "tool_calls") {
          abnormalFinish = choice.finish_reason;
        }
        flushTools();
      }
    }
    flushTools();
    if (!emitted && abnormalFinish) {
      throw new Error(`The model returned no content (finish_reason=${abnormalFinish})`);
    }
    if (!emitted && !sawFinish) {
      throw new Error("The model returned no content (empty stream)");
    }
    if (stopReason) cb.onStopReason?.(stopReason);
  }
  var AiCreditsError;
  var init_stream = __esm({
    "web-src/ai-provider/stream.ts"() {
      "use strict";
      init_fetch();
      init_http_error();
      init_providers();
      init_watchdog();
      AiCreditsError = class extends Error {
        constructor(notice) {
          super(notice);
          this.name = "AiCreditsError";
        }
      };
    }
  });

  // web-src/ai-provider/index.ts
  var init_ai_provider = __esm({
    "web-src/ai-provider/index.ts"() {
      "use strict";
      init_types2();
      init_stream();
      init_providers();
      init_fetch();
      init_watchdog();
      init_http_error();
    }
  });

  // web-src/entry.ts
  var require_entry = __commonJS({
    "web-src/entry.ts"() {
      init_agent_core();
      init_ai_provider();
      var pendingToolCalls = /* @__PURE__ */ new Map();
      chrome.webview.addEventListener("message", (ev) => {
        const data = ev.data;
        if (!data || data.kind !== "tool-result") return;
        const resolve = pendingToolCalls.get(data.requestId);
        if (!resolve) return;
        pendingToolCalls.delete(data.requestId);
        resolve({
          output: data.output,
          isError: data.isError,
          mutated: data.mutated,
          summary: data.summary
        });
      });
      function callDotNetTool(toolName, input2) {
        const requestId = crypto.randomUUID();
        return new Promise((resolve) => {
          pendingToolCalls.set(requestId, resolve);
          const msg = { kind: "tool-call", requestId, toolName, input: input2 };
          chrome.webview.postMessage(msg);
        });
      }
      var PROVIDER_CONFIG = {
        apiKey: "test",
        model: "test-model"
      };
      var BASE_URL = "http://127.0.0.1:9000/v1";
      var MAX_TOKENS = 1024;
      function makeTransport() {
        return {
          stream(request, callbacks) {
            const controller = new AbortController();
            const t0 = performance.now();
            let chunkIndex = 0;
            streamOpenAiCompatible(
              BASE_URL,
              PROVIDER_CONFIG,
              request.system,
              request.messages,
              request.tools,
              MAX_TOKENS,
              {
                onDelta: (text) => {
                  chunkIndex++;
                  appendLine(`  [chunk ${chunkIndex} @ +${Math.round(performance.now() - t0)}ms] ${JSON.stringify(text)}`);
                  callbacks.onDelta(text);
                },
                onToolCall: callbacks.onToolCall,
                onStopReason: callbacks.onStopReason,
                signal: controller.signal
              }
            ).then(() => callbacks.onDone()).catch((e) => callbacks.onError(e instanceof Error ? e.message : String(e)));
            return { cancel: () => controller.abort() };
          }
        };
      }
      var wordSkill = {
        id: "spike3-word-tools",
        systemPrompt: "You are a test assistant running inside a VSTO Word add-in spike (spike 3: real COM tool execution). You can read the document, insert text, and create/edit a native Word chart. Use the tools when asked to.",
        tools: [
          {
            name: "get_document_context",
            description: "Reads the active Word document's paragraph/word count and a text preview.",
            inputSchema: { type: "object", properties: {} }
          },
          {
            name: "insert_content",
            description: "Inserts a paragraph of text at the end of the active Word document.",
            inputSchema: {
              type: "object",
              properties: { text: { type: "string" } },
              required: ["text"]
            }
          },
          {
            name: "edit_chart",
            description: "Creates (if none exists) or edits a native Word chart: sets its title and its first series values.",
            inputSchema: {
              type: "object",
              properties: {
                title: { type: "string" },
                values: { type: "array", items: { type: "number" } }
              },
              required: ["title", "values"]
            }
          }
        ],
        executeTool: (call) => callDotNetTool(call.name, call.input)
      };
      var loop = new AgentLoop({
        transport: makeTransport(),
        skill: wordSkill,
        events: {
          onText: (text) => setAssistantBubble(text),
          onToolStart: (call) => {
            appendLine(`  [tool call] ${call.name}(${JSON.stringify(call.input)})`);
          },
          onToolExecuted: (event) => {
            appendLine(
              `  [tool result] ${event.call.name} -> ${event.execution.isError ? "ERROR: " : ""}${event.execution.output}`
            );
          },
          onTurnEnd: () => appendLine("  [turn end - back to model]"),
          onDone: (result) => {
            setAssistantBubble(result.text || "(no text)");
            setBusy(false);
          },
          onError: (error) => {
            appendLine(`[error] ${error}`);
            setBusy(false);
          }
        }
      });
      var transcript = document.getElementById("transcript");
      var input = document.getElementById("input");
      var sendBtn = document.getElementById("sendBtn");
      var assistantBubble = null;
      function appendLine(text) {
        const div = document.createElement("div");
        div.className = "line";
        div.textContent = text;
        transcript.appendChild(div);
        transcript.scrollTop = transcript.scrollHeight;
      }
      function setAssistantBubble(text) {
        if (!assistantBubble) {
          assistantBubble = document.createElement("div");
          assistantBubble.className = "line assistant";
          transcript.appendChild(assistantBubble);
        }
        assistantBubble.textContent = "assistant: " + text;
        transcript.scrollTop = transcript.scrollHeight;
      }
      function setBusy(busy) {
        sendBtn.disabled = busy;
        input.disabled = busy;
      }
      function send() {
        const text = input.value.trim();
        if (!text || loop.busy) return;
        appendLine("user: " + text);
        input.value = "";
        assistantBubble = null;
        setBusy(true);
        loop.run(text);
      }
      sendBtn.addEventListener("click", send);
      input.addEventListener("keydown", (e) => {
        if (e.key === "Enter") send();
      });
      appendLine("[spike 2 ready] talking to " + BASE_URL);
    }
  });
  require_entry();
})();
//# sourceMappingURL=bundle.js.map
