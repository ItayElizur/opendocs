# PP-17: Sheet-Scoped Defined Names — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Source item:** `docs/superpowers/plans/2026-08-23-ai-addin-product-plan.md` § PP-17 (P2).

**Goal:** `add_defined_name` and `delete_defined_name` accept an optional sheet scope, so the model can create and remove the sheet-scoped names that `read_sheet_features` already reports back to it.

**Architecture:** The asymmetry is exact and small. `ReadSheetFeatures` enumerates **both** collections and labels them (`ExcelAiAddIn/ExcelTools.cs:293-300`):

```csharp
foreach (Excel.Name n in sheet.Names)      // "Defined name (sheet-scoped): ..."
foreach (Excel.Name n in ActiveWorkbook.Names)  // "Defined name (workbook-scoped): ..."
```

while `AddDefinedName` (`:1124-1129`) and `DeleteDefinedName` (`:1131-1135`) only ever touch `ActiveWorkbook.Names`. So the model can *see* a sheet-scoped name, is told what it is called, and has no operation that can create or delete one — including no way to delete a sheet-scoped name it just discovered.

Excel's object model handles both through the same `Names` collection API; a worksheet has its own `Worksheet.Names` whose `Add` takes the same `Name`/`RefersTo` arguments. So the change is: resolve the target collection from an optional `scope`/`sheet` parameter, then use it.

Two subtleties that make the naive version wrong:
- **Deletion ambiguity.** A workbook can hold both a workbook-scoped `Sales` and a sheet-scoped `Sales` on Sheet2. `Names.Item(name)` on the workbook collection may resolve either depending on context. Deletion must be explicit about which one it targets.
- **`RefersTo` qualification.** `AddDefinedName` builds `"=" + reference` (`:1128`). An unqualified `A1:B10` in a sheet-scoped name resolves relative to that sheet — usually what the caller wants — but in a workbook-scoped name it resolves against whatever sheet is active at evaluation time, which is a latent wrong-answer bug that exists today.

**Tech Stack:** C# 7.3 / .NET Framework 4.8, `Microsoft.Office.Interop.Excel`; TypeScript for the schema.

## Global Constraints

- C# 7.3 / .NET Framework 4.8 only — no `using`-declarations, no target-typed `new()`, no switch expressions.
- Default behavior with no new parameter stays **workbook-scoped**, exactly as today.
- Do not change `read_sheet_features` — it is already correct and is the reason this gap is visible.
- No automated tests for COM executor methods (project convention). Verification is build + Task 4's manual matrix.
- Rebuild bundle + MSBuild after the `entry.ts` change.
- If PP-5 has landed, express the new parameters through its `EXCEL_OPS` table.

---

### Task 1: `add_defined_name` accepts a scope

**Files:**
- Modify: `ExcelAiAddIn/ExcelTools.cs`

- [ ] **Step 1:** Rewrite `AddDefinedName` (`:1124-1129`):

```csharp
private static void AddDefinedName(JsonElement op)
{
    string name = op.GetProperty("name").GetString();
    string reference = op.GetProperty("ref").GetString();
    bool sheetScoped = op.TryGetProperty("scope", out var sc) && sc.ValueKind == JsonValueKind.String
                       && sc.GetString() == "sheet";

    string refersTo = reference.StartsWith("=") ? reference : "=" + reference;

    if (sheetScoped)
    {
        Excel.Worksheet sheet = Sheet(op);   // honors the existing optional "sheet" property
        sheet.Names.Add(name, refersTo);
    }
    else
    {
        Globals.ThisAddIn.Application.ActiveWorkbook.Names.Add(name, refersTo);
    }
}
```

Reusing `Sheet(op)` (`:111-119`) means `sheet` keeps the meaning it has in every other operation in this file — no second sheet-resolution convention.

- [ ] **Step 2:** The `reference.StartsWith("=")` guard is a fix in its own right: today `"=A1:B10"` becomes `"==A1:B10"` and throws an opaque COM error.
- [ ] **Step 3: Qualify unqualified references.** When the reference contains no `!`, prefix it with the target sheet's name so the name resolves deterministically instead of against whatever sheet happens to be active. Quote the sheet name if it contains spaces (`'My Sheet'!$A$1:$B$10`). Do this for both scopes — the workbook-scoped case has the same latent bug today.
- [ ] **Step 4: Validate the name.** Excel rejects names that collide with a cell address (`A1`), start with a digit, or contain spaces — with an unhelpful COM error. Pre-check and throw a specific message.
- [ ] **Step 5: Existing-name behavior.** `Names.Add` with a name that already exists *replaces* it silently. Decide and document: accept an optional `overwrite` (default `false`), and without it throw if the name already exists in the target scope. Silent replacement of a defined name in a financial model is a genuinely damaging silent-success.

**Verification:** `MSBuild ExcelAiAddIn/ExcelAiAddIn.csproj -t:Build -p:Configuration=Debug`; a sheet-scoped name created here appears in Excel's Name Manager with the sheet in its Scope column.

---

### Task 2: `delete_defined_name` accepts a scope

**Files:**
- Modify: `ExcelAiAddIn/ExcelTools.cs`

- [ ] **Step 1:** Rewrite `DeleteDefinedName` (`:1131-1135`) to resolve the same way: `sheet.Names.Item(name).Delete()` for `scope: 'sheet'`, workbook collection otherwise.
- [ ] **Step 2: Missing name → specific error.** `Names.Item(name)` on an absent name throws a bare COM error. Catch it and report `"delete_defined_name: no <scope>-scoped name '<name>' found."` — including which scope was searched, so the model can retry with the other one.
- [ ] **Step 3: Ambiguity.** With no `scope`, only the workbook collection is searched (today's behavior). If the name is not found there but *does* exist sheet-scoped somewhere, say so in the error: "not found workbook-scoped; a sheet-scoped name with this name exists on Sheet2 — pass scope:'sheet' and sheet:'Sheet2'". That one sentence turns a dead end into a self-correcting next turn.

**Verification:** build; deleting a sheet-scoped name works; deleting a nonexistent one produces the actionable message.

---

### Task 3: Schema

**Files:**
- Modify: `ExcelAiAddIn/web-src/entry.ts`

- [ ] **Step 1:** Update both operations in `propose_operations`'s Data group (`ExcelAiAddIn/web-src/entry.ts:218`), currently `"add_defined_name" (name, ref), "delete_defined_name" (name)`:

```
"add_defined_name" (sheet?, scope?:"workbook"|"sheet" (default workbook; "sheet" scopes it to the sheet named by sheet? or the active sheet), name, ref, overwrite?),
"delete_defined_name" (sheet?, scope?:"workbook"|"sheet", name)
```

- [ ] **Step 2:** `scope` gets a real `enum`.
- [ ] **Step 3:** Note in the description that `read_sheet_features` reports both scopes and labels each — so the model knows how to discover what exists before acting.
- [ ] **Step 4:** Document the `ref` convention: an unqualified range is auto-qualified to the target sheet; pass a fully qualified `Sheet1!$A$1:$B$10` to be explicit.

**Verification:** bundle rebuilds; "create a sheet-scoped name Rates for B2:B20 on the Assumptions sheet" produces a correct single operation.

---

### Task 4: Manual verification matrix

- [ ] `add_defined_name({name, ref})` with no scope → workbook-scoped, exactly as before (no regression).
- [ ] `add_defined_name({name, ref, scope:'sheet'})` → Name Manager shows the sheet in Scope.
- [ ] `add_defined_name({name, ref, scope:'sheet', sheet:'Assumptions'})` → scoped to that sheet, not the active one.
- [ ] `ref: '=A1:B10'` (leading `=`) → works rather than producing `==`.
- [ ] Unqualified `ref` → resolves to the intended sheet even after switching the active sheet.
- [ ] Duplicate name without `overwrite` → specific error; with `overwrite: true` → replaced.
- [ ] Invalid name (`A1`, `1x`, `has space`) → specific error naming the rule.
- [ ] `delete_defined_name({name, scope:'sheet', sheet:'Assumptions'})` → removes the sheet-scoped one, leaves a same-named workbook-scoped one intact.
- [ ] Deleting a nonexistent name → error that names the searched scope and points at the other scope if the name exists there.
- [ ] `read_sheet_features` before and after each of the above reflects the change — the discovery loop the model actually uses.
