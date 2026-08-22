# Mock Server Editing-Mode Test Harness Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Turn the throwaway OpenAI-compatible mock server at `C:\dev\openai-mock-server` into a harness that lets a developer manually verify the 4 editing modes (Read Only / Comment Only / Track Changes / Full Autonomy) across all 19 real tools in Word/Excel/PowerPoint — both "does the client only offer allowed tools" and "does the server-side C# gate in `WordTools.cs`/`ExcelTools.cs`/`PowerPointTools.cs` actually reject a tool call that shouldn't have been possible."

**Architecture:** No new services or files. Two additions to the existing single-file FastAPI app (`openai_server/main.py`): (1) a complete per-tool demo-arguments table covering all 19 real tools instead of the current 3, so the server's existing "call everything you were offered" behavior no longer sends nonsense arguments that crash the C# tool executors; (2) a `FORCE_TOOL:<toolName>` trigger parsed from the latest user chat message, which — when present — makes the mock call that exact tool regardless of what was actually offered in `request.tools`, so a developer can deliberately try to call a tool the client-side filtering would never have offered (e.g. `insert_content` while in Read Only mode) and confirm the C#-side gate blocks it independently. This exercises the "defense-in-depth" server-side gate that per-app `WordTools.cs`/`ExcelTools.cs`/`PowerPointTools.cs` already implement but had no way to be deliberately tripped from outside.

**Tech Stack:** Python 3.10+, FastAPI, pytest (existing `tests/test_main.py` conventions), `fastapi.testclient.TestClient`.

**Spec:** No separate spec document — the tool argument shapes are read directly from the real, already-implemented tool schemas in `officeoffice`'s `WordAiAddIn/web-src/entry.ts`, `ExcelAiAddIn/web-src/entry.ts`, `PowerPointAiAddIn/web-src/entry.ts` (client-side JSON schemas) and their corresponding C# executors (`WordTools.cs`, `ExcelTools.cs`, `PowerPointTools.cs`), which this plan treats as the source of truth.

## Global Constraints

- All work happens in `C:\dev\openai-mock-server` — a separate git repository from `officeoffice`. Do not touch anything under `C:\dev\officeoffice`.
- This is a throwaway local dev tool, not shipped code — no need for production hardening (auth, rate limiting, etc.) beyond what already exists.
- Preserve all existing behavior for requests that don't use `FORCE_TOOL:` — the default "call every offered tool with matching demo args" behavior must keep working exactly as before for tools it already covered, and must now also work correctly for the 16 tools it previously mishandled.
- Preserve the existing CORS, streaming, and Unicode-safety fixes already in `main.py` (`CORSMiddleware`, `sys.stdout.reconfigure`, `await asyncio.sleep`) — do not revert or bypass them.
- New Python code targets the same Python version already in use (3.10+ per `pyproject.toml`) — no need to support older versions.

---

### Task 1: Complete the per-tool demo-arguments table for all 19 real tools

**Files:**
- Modify: `openai_server/main.py:47-51` (the `_DEMO_TOOL_ARGS` dict)
- Test: `tests/test_main.py`

**Interfaces:**
- Consumes: nothing new — `_DEMO_TOOL_ARGS` is already looked up by tool name inside `chat_completions()`'s existing tool-calling branch (`main.py:60-63`).
- Produces: a complete dict covering all 19 real tool names below, each mapping to a JSON-serializable dict matching that tool's actual required-fields schema. Task 2 reuses this same dict for the `FORCE_TOOL:` path.

The 19 real tool names and their exact required-argument shapes, read directly from the current codebase:

**Word (`WordAiAddIn/web-src/entry.ts`, `ALL_WORD_TOOLS`):**
- `get_document_context` — `{}`
- `insert_content` — `{"text": string}`
- `edit_chart` — `{"title": string, "values": number[]}`
- `read_blocks` — `{"startIndex": number, "endIndex": number}`
- `replace_blocks` — `{"startIndex": number, "endIndex": number, "text": string}`
- `apply_commands` — `{"commands": [{"kind": string, ...}]}` (kinds: `set_bold`/`set_italic` need `startIndex`, `endIndex`, `value`; `set_heading` needs `index`, `level`; `find_replace` needs `find`, `replace`)
- `add_comment` — `{"anchorText": string, "commentText": string}`

**Excel (`ExcelAiAddIn/web-src/entry.ts`, and `ExcelTools.cs:113-145` for `propose_operations`'s sub-kinds):**
- `get_workbook_context` — `{}`
- `read_range` — `{"address": string}` (sheet optional)
- `read_cells` — `{"addresses": string[]}`
- `propose_operations` — `{"operations": [{"kind": string, ...}]}` (kinds: `set_cell` needs `address`, `value`; `set_formula` needs `address`, `formula`; `set_range`/`format_range`/`insert_rows`/`delete_rows`/`insert_cols`/`delete_cols`/`add_chart` have their own shapes — the demo only needs one representative op)

**PowerPoint (`PowerPointAiAddIn/web-src/entry.ts`):**
- `get_deck_context` — `{}`
- `read_slide` — `{"slideIndex": number}`
- `set_element_text` — `{"slideIndex": number, "shapeIndex": number, "text": string}`
- `set_element_style` — `{"slideIndex": number, "shapeIndex": number, "bold"?: boolean, "italic"?: boolean, "fontSize"?: number, "color"?: string}`
- `set_element_transform` — `{"slideIndex": number, "shapeIndex": number, "left"?: number, "top"?: number, "width"?: number, "height"?: number, "rotation"?: number}`
- `add_text_box` — `{"slideIndex": number, "left": number, "top": number, "width": number, "height": number, "text": string}`
- `add_shape` — `{"slideIndex": number, "shapeType": string, "left": number, "top": number, "width": number, "height": number}`
- `delete_element` — `{"slideIndex": number, "shapeIndex": number}`

Every demo argument set below uses `slideIndex`/`shapeIndex`/`startIndex`/`endIndex` of `0` and small placeholder content, matching the style already established by the 3 existing entries (`insert_content`'s "Hello from the AI agent running inside VSTO Word!"). Real documents/sheets/decks opened for manual testing should have at least one paragraph/cell/slide+shape so these indices resolve to something real rather than throwing an out-of-range error from the C# side (that's expected and fine — it's the caller's job to have a non-empty document open, not this mock's).

- [ ] **Step 1: Replace `_DEMO_TOOL_ARGS` with the complete table**

Replace `openai_server/main.py:47-51`:
```python
_DEMO_TOOL_ARGS = {
    "get_document_context": {},
    "insert_content": {"text": "Hello from the AI agent running inside VSTO Word!"},
    "edit_chart": {"title": "Quarterly Revenue", "values": [10, 25, 15, 30]},
}
```
with:
```python
_DEMO_TOOL_ARGS = {
    # Word
    "get_document_context": {},
    "insert_content": {"text": "Hello from the AI agent running inside VSTO Word!"},
    "edit_chart": {"title": "Quarterly Revenue", "values": [10, 25, 15, 30]},
    "read_blocks": {"startIndex": 0, "endIndex": 2},
    "replace_blocks": {"startIndex": 0, "endIndex": 0, "text": "Replaced by the AI agent."},
    "apply_commands": {
        "commands": [
            {"kind": "set_bold", "startIndex": 0, "endIndex": 0, "value": True},
        ]
    },
    "add_comment": {"anchorText": "the", "commentText": "AI-added comment for testing."},
    # Excel
    "get_workbook_context": {},
    "read_range": {"address": "A1:C10"},
    "read_cells": {"addresses": ["A1", "B2"]},
    "propose_operations": {
        "operations": [
            {"kind": "set_cell", "address": "A1", "value": "AI wrote this"},
        ]
    },
    # PowerPoint
    "get_deck_context": {},
    "read_slide": {"slideIndex": 0},
    "set_element_text": {"slideIndex": 0, "shapeIndex": 0, "text": "AI-updated text."},
    "set_element_style": {"slideIndex": 0, "shapeIndex": 0, "bold": True},
    "set_element_transform": {"slideIndex": 0, "shapeIndex": 0, "left": 100.0, "top": 100.0},
    "add_text_box": {
        "slideIndex": 0, "left": 50.0, "top": 50.0, "width": 200.0, "height": 80.0,
        "text": "New text box from the AI agent.",
    },
    "add_shape": {
        "slideIndex": 0, "shapeType": "rectangle",
        "left": 50.0, "top": 50.0, "width": 100.0, "height": 60.0,
    },
    "delete_element": {"slideIndex": 0, "shapeIndex": 0},
}
```

- [ ] **Step 2: Write a test verifying every tool name in a synthetic request gets a matching, non-fallback demo arg set**

Append to `tests/test_main.py`:
```python
from openai_server.main import _DEMO_TOOL_ARGS

ALL_19_TOOL_NAMES = [
    "get_document_context", "insert_content", "edit_chart", "read_blocks",
    "replace_blocks", "apply_commands", "add_comment",
    "get_workbook_context", "read_range", "read_cells", "propose_operations",
    "get_deck_context", "read_slide", "set_element_text", "set_element_style",
    "set_element_transform", "add_text_box", "add_shape", "delete_element",
]

def test_all_19_real_tools_have_demo_args():
    missing = [name for name in ALL_19_TOOL_NAMES if name not in _DEMO_TOOL_ARGS]
    assert missing == [], f"Missing demo args for: {missing}"

def _tool_schema(name):
    return {"type": "function", "function": {"name": name, "description": "", "parameters": {}}}

def test_chat_completions_calls_every_offered_tool_with_its_real_demo_args():
    response = client.post(
        "/v1/chat/completions",
        json={
            "model": "test-model",
            "messages": [{"role": "user", "content": "do everything"}],
            "tools": [_tool_schema(name) for name in ALL_19_TOOL_NAMES],
        },
    )
    assert response.status_code == 200
    tool_calls = response.json()["choices"][0]["message"]["tool_calls"]
    called_names = {tc["function"]["name"] for tc in tool_calls}
    assert called_names == set(ALL_19_TOOL_NAMES)
    for tc in tool_calls:
        import json as _json
        args = _json.loads(tc["function"]["arguments"])
        assert args == _DEMO_TOOL_ARGS[tc["function"]["name"]]
```

- [ ] **Step 3: Run the tests**

Run: `cd C:\dev\openai-mock-server && poetry run pytest tests/test_main.py -k "demo_args or every_offered_tool" -v`
Expected: both new tests PASS.

- [ ] **Step 4: Commit**

```bash
cd C:/dev/openai-mock-server
git add openai_server/main.py tests/test_main.py
git commit -m "feat: complete demo-args table for all 19 real Office add-in tools"
```

---

### Task 2: `FORCE_TOOL:<name>` trigger to test the server-side editing-mode gate

**Files:**
- Modify: `openai_server/main.py` (the `chat_completions` handler)
- Test: `tests/test_main.py`

**Interfaces:**
- Consumes: `_DEMO_TOOL_ARGS` from Task 1 (must already be complete before this task, since the whole point is being able to force-call any of the 19 real tools, not just the original 3).
- Produces: when the latest message with `role == "user"` in `request.messages` has content containing the substring `FORCE_TOOL:<toolName>` (case-sensitive, `<toolName>` is one contiguous token of word characters/underscores immediately following the colon, e.g. `FORCE_TOOL:insert_content`), the response returns a single `tool_calls` entry for exactly that tool name with its demo args from `_DEMO_TOOL_ARGS` (or `{}` if the name isn't in the table, so a typo'd/nonexistent name still round-trips visibly as an "Unknown tool" error from the real C# side rather than crashing the mock) — regardless of whether that tool appears in `request.tools` at all. This lets a developer type e.g. `FORCE_TOOL:insert_content` into the chat box while the add-in is in Read Only mode (which would never normally offer `insert_content` to the model) and see whether `WordTools.Execute`'s server-side gate rejects it with `"Blocked: editing mode is Read Only."` — proving the C# gate works independently of client-side filtering.
- This check must run BEFORE the existing `if request.tools and not already_ran_tools:` branch, and must short-circuit it (a forced call always wins, whether or not `tools` were offered) — but must NOT trigger on the second turn (the turn where the tool result comes back), since by then `already_ran_tools` is true and the trigger phrase is still sitting in the earlier user message. Scope the check to only the LAST message in `request.messages`, and only when that last message's role is `user` (i.e. only fires on the turn immediately following the user typing the trigger phrase, not on every subsequent turn of the same conversation).

- [ ] **Step 1: Write the failing test**

Append to `tests/test_main.py`:
```python
import re as _re

def test_force_tool_trigger_calls_named_tool_even_if_not_offered():
    response = client.post(
        "/v1/chat/completions",
        json={
            "model": "test-model",
            "messages": [{"role": "user", "content": "please FORCE_TOOL:insert_content now"}],
            # Deliberately do NOT offer insert_content - simulates Read Only
            # mode, where the client would only ever offer read-only tools.
            "tools": [_tool_schema("get_document_context")],
        },
    )
    assert response.status_code == 200
    tool_calls = response.json()["choices"][0]["message"]["tool_calls"]
    assert len(tool_calls) == 1
    assert tool_calls[0]["function"]["name"] == "insert_content"
    import json as _json
    assert _json.loads(tool_calls[0]["function"]["arguments"]) == _DEMO_TOOL_ARGS["insert_content"]

def test_force_tool_trigger_does_not_fire_on_a_later_turn():
    response = client.post(
        "/v1/chat/completions",
        json={
            "model": "test-model",
            "messages": [
                {"role": "user", "content": "please FORCE_TOOL:insert_content now"},
                {"role": "assistant", "content": None, "tool_calls": [
                    {"id": "call_1", "type": "function",
                     "function": {"name": "insert_content", "arguments": "{}"}},
                ]},
                {"role": "tool", "content": "Inserted.", "tool_call_id": "call_1"},
            ],
            "tools": [_tool_schema("get_document_context")],
        },
    )
    assert response.status_code == 200
    # Second turn: no forced tool call left to re-trigger, and no new
    # user message asking for tools, so this should fall through to the
    # plain text-response path (no tool_calls key at all).
    assert "tool_calls" not in response.json()["choices"][0]["message"] or \
        response.json()["choices"][0]["message"].get("tool_calls") is None
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `cd C:\dev\openai-mock-server && poetry run pytest tests/test_main.py -k force_tool -v`
Expected: FAIL — `insert_content` is not currently offered, so today's code calls only `get_document_context`.

- [ ] **Step 3: Implement the trigger**

In `openai_server/main.py`, add near the top of the file (after the existing `_DEMO_TOOL_ARGS` dict from Task 1):
```python
import re

_FORCE_TOOL_PATTERN = re.compile(r"FORCE_TOOL:(\w+)")


def _find_forced_tool_name(request: ChatCompletionRequest):
    """Manual test hook: if the LAST message is from the user and contains
    FORCE_TOOL:<name>, the mock calls exactly that tool regardless of what
    was offered in request.tools - lets a developer deliberately try to
    call a tool the client-side editing-mode filtering would never offer,
    to verify the server-side gate in WordTools.cs/ExcelTools.cs/
    PowerPointTools.cs rejects it independently."""
    if not request.messages:
        return None
    last = request.messages[-1]
    if last.role != "user" or not last.content:
        return None
    match = _FORCE_TOOL_PATTERN.search(last.content)
    return match.group(1) if match else None
```

Then in `chat_completions`, replace the branch condition at `main.py:59` (currently `if request.tools and not already_ran_tools:`) — read the current file first to get the exact surrounding lines, then restructure so the forced-tool check runs first and reuses the same response-building code:

```python
@app.post("/v1/chat/completions")
async def chat_completions(request: ChatCompletionRequest):
    print(f"Request to /v1/chat/completions: {request.dict()}")
    already_ran_tools = any(m.role == "tool" for m in request.messages)
    forced_tool_name = None if already_ran_tools else _find_forced_tool_name(request)

    if forced_tool_name or (request.tools and not already_ran_tools):
        if forced_tool_name:
            tool_names = [forced_tool_name]
        else:
            tool_names = [tool['function']['name'] for tool in request.tools]

        tool_calls = []
        for name in tool_names:
            args = _DEMO_TOOL_ARGS.get(name, {})
            tool_calls.append(ToolCall(
                id=f"call_{name}",
                type="function",
                function=Function(
                    name=name,
                    arguments=json.dumps(args)
                )
            ))
        message = ChatCompletionMessageWithToolCalls(
            role="assistant",
            tool_calls=tool_calls
        )
        choice = ChatCompletionResponseChoiceWithToolCalls(
            index=0,
            message=message,
            finish_reason="tool_calls"
        )
        response = ChatCompletionResponseWithToolCalls(
            id="chatcmpl-123",
            object="chat.completion",
            created=int(time.time()),
            model=request.model,
            choices=[choice],
            usage={"prompt_tokens": 0, "completion_tokens": 0, "total_tokens": 0},
            tool_choice="auto"
        )
        return response

    # ... rest of the function (streaming/non-streaming plain text response) unchanged
```

Note the fallback for an unknown/mistyped forced tool name changed from `{"location": "Boston, MA"}` to `{}` — an empty-args call to a real tool name will surface as a real, informative C# exception (e.g. a `KeyNotFoundException` from a missing required `JsonElement` property) when routed through the real add-in, which is more useful for testing than a silently-wrong fake location string. This also applies to the pre-existing default path now (any tool name not in `_DEMO_TOOL_ARGS`, forced or not) — update the fallback in the same place it already exists.

- [ ] **Step 4: Run the tests and confirm they pass**

Run: `cd C:\dev\openai-mock-server && poetry run pytest tests/test_main.py -v`
Expected: all tests PASS, including the two new ones and Task 1's two new ones. (Pre-existing failures unrelated to this change, if any, are out of scope — do not fix unrelated pre-existing test breakage as part of this task.)

- [ ] **Step 5: Commit**

```bash
cd C:/dev/openai-mock-server
git add openai_server/main.py tests/test_main.py
git commit -m "feat: FORCE_TOOL trigger to test server-side editing-mode gate directly"
```

---

### Task 3: Testing guide documenting the mode-limits procedure

**Files:**
- Create: `docs/mode-testing.md`

**Interfaces:**
- Consumes: nothing (documentation only).
- Produces: a short, concrete runbook a developer can follow without re-deriving the tool names or the trigger syntax from source.

- [ ] **Step 1: Write the guide**

Create `docs/mode-testing.md`:
```markdown
# Testing editing-mode limits against the mock server

This mock (`openai_server/main.py`) supports two things useful for verifying
the 4 editing modes (Read Only / Comment Only / Track Changes / Full
Autonomy) in the Word/Excel/PowerPoint add-ins:

## 1. Normal flow: verify the client only offers allowed tools

Start the mock (`poetry run uvicorn openai_server.main:app --reload --port 9000`
or `run_server.bat`), open the add-in, switch to a mode, and send any
message. The mock calls every tool it was offered in that turn's
`request.tools` - so the "Running N tools..." group shown in the chat UI
tells you exactly which tools the client considered allowed for the current
mode:

| Mode | Tools that should get called |
|---|---|
| Read Only | Word: `get_document_context`, `read_blocks`. Excel: `get_workbook_context`, `read_range`, `read_cells`. PowerPoint: `get_deck_context`, `read_slide`. |
| Comment Only | Read Only's set, plus Word's `add_comment` (Excel/PowerPoint have no comment-equivalent tool yet - same set as Read Only). |
| Track Changes / Full Autonomy | The app's full tool list. |

If a tool outside that table gets called for a given mode, the client-side
filtering in that app's `web-src/entry.ts` (`toolsForMode()`) has a bug.

## 2. Adversarial flow: verify the server-side gate independently

Type a message containing `FORCE_TOOL:<toolName>` (e.g.
`FORCE_TOOL:insert_content`) while the add-in is in a mode that should
block that tool. The mock will call exactly that tool regardless of what
was offered - simulating a hallucinating or malicious model that ignores
the offered tool list. Real tool names (see `_DEMO_TOOL_ARGS` in
`openai_server/main.py` for the full list across all three apps).

Expected result: the tool step in the chat UI shows an error, and its
output text is the blocked message from the relevant `*Tools.cs`'s
`Execute()` method (e.g. `"Blocked: editing mode is Read Only."`) - proving
the gate in `WordTools.cs`/`ExcelTools.cs`/`PowerPointTools.cs` rejects the
call on the C# side, not just via client-side omission.

If instead the tool actually runs and mutates the document/sheet/deck, the
server-side gate for that app has a bug - the whole point of "defense in
depth" is that this must never happen regardless of what the client offers.
```

- [ ] **Step 2: Commit**

```bash
cd C:/dev/openai-mock-server
git add docs/mode-testing.md
git commit -m "docs: add editing-mode testing guide for the mock server"
```
