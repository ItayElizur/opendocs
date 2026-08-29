import { AgentLoop, type AgentSkill } from '@genoffice/agent-core'
import { streamForProvider, type AiProviderConfig } from '@genoffice/ai-provider'
import { mountChatUI, type EditingMode, type SelectionExtent, type ToolDisplayEntry } from '@officeai/chat-ui'
import {
  callDotNetTool,
  initBridge,
  persistMessage,
  postCollapse,
  postMode,
  postNewChatDivider,
  postTlsBypass,
  requestDocSettings,
  requestHistory,
  saveDocSettings,
  type RawSelectionPayload,
} from './bridge'
import { getSettings, makeTransport, setSettings } from './settings'

export interface ToolDisplayInfo {
  label: { en: string; he: string }
  description: { en: string; he: string }
}

/**
 * FT-2 Task 4: what the user has selected, classified from the raw bridge
 * payload into the vocabulary each app's tools actually take - Word's
 * paragraph-index tools get content, Excel's A1-addressed tools get an
 * address, PowerPoint's slideIndex/shapeIndex tools get indices. The `range`
 * variant carries more than the plan's minimal sketch (entireColumns/
 * entireRows/effectiveAddress/effectiveCellCount) because describeSelection's
 * whole-column/whole-row wording needs them - there is nowhere else for that
 * data to live.
 */
export type SelectionContext =
  | { kind: 'none' }
  | {
      kind: 'text'
      preview: string
      fullText: string
      startBlockIndex: number
      endBlockIndex: number
      objectKind: 'table' | 'chart' | 'smartart' | null
      objectIndex: number
    }
  | {
      kind: 'range'
      sheet: string
      address: string
      cellCount: number
      multi: boolean
      areaCount: number
      entireColumns: boolean
      entireRows: boolean
      effectiveAddress: string | null
      effectiveCellCount: number
    }
  | { kind: 'slides'; slideIndexes: number[] }
  | { kind: 'shapes'; slideIndex: number; shapeIndexes: number[]; names: string[]; textPreview: string[] }
  | { kind: 'shapeText'; slideIndex: number; shapeIndex: number; text: string }
  | {
      kind: 'mail'
      count: number
      entryIds: string[]
      subject: string
      senderName: string
      folderName: string
      conversationTopic: string | null
    }

function toSelectionContext(raw: RawSelectionPayload): SelectionContext {
  if (!raw.hasSelection) return { kind: 'none' }
  if (raw.app === 'excel') {
    return {
      kind: 'range',
      sheet: raw.sheet ?? '',
      address: raw.address ?? '',
      cellCount: raw.cellCount ?? 0,
      multi: raw.multi ?? false,
      areaCount: raw.areaCount ?? 1,
      entireColumns: raw.entireColumns ?? false,
      entireRows: raw.entireRows ?? false,
      effectiveAddress: raw.effectiveAddress ?? null,
      effectiveCellCount: raw.effectiveCellCount ?? 0,
    }
  }
  if (raw.app === 'outlook') {
    const entryIds = raw.entryIds ?? []
    if (entryIds.length === 0) return { kind: 'none' }
    return {
      kind: 'mail',
      count: raw.count ?? entryIds.length,
      entryIds,
      subject: raw.subject ?? '',
      senderName: raw.senderName ?? '',
      folderName: raw.folderName ?? '',
      conversationTopic: raw.conversationTopic ?? null,
    }
  }
  if (raw.app === 'powerpoint') {
    if (raw.selKind === 'slides') return { kind: 'slides', slideIndexes: raw.slideIndexes ?? [] }
    if (raw.selKind === 'shapes') {
      return {
        kind: 'shapes',
        slideIndex: raw.slideIndex ?? 0,
        shapeIndexes: raw.shapeIndexes ?? [],
        names: raw.names ?? [],
        textPreview: raw.textPreview ?? [],
      }
    }
    if (raw.selKind === 'shapeText') {
      return { kind: 'shapeText', slideIndex: raw.slideIndex ?? 0, shapeIndex: raw.shapeIndex ?? 0, text: raw.text ?? '' }
    }
    return { kind: 'none' }
  }
  // Word: no `app` field, only hasSelection/preview/fullText(/startBlockIndex/endBlockIndex/objectKind/objectIndex).
  return {
    kind: 'text',
    preview: raw.preview ?? '',
    fullText: raw.fullText ?? '',
    startBlockIndex: raw.startBlockIndex ?? -1,
    endBlockIndex: raw.endBlockIndex ?? -1,
    objectKind: raw.objectKind ?? null,
    objectIndex: raw.objectIndex ?? -1,
  }
}

/**
 * FT-2 Task 4 Step 2: the per-turn sentence injected via buildContext() -
 * must state the addressing vocabulary explicitly (an A1 address or a
 * slideIndex/shapeIndex pair) so the selection is actionable, not merely
 * informative. Kind-keyed rather than per-app-configurable because each app
 * only ever produces its own subset of kinds (Word: text/none; Excel:
 * range/none; PowerPoint: slides/shapes/shapeText/none) - AddInConfig.
 * describeSelection lets an app override this if it ever needs to.
 */
function defaultDescribeSelection(ctx: SelectionContext): string {
  switch (ctx.kind) {
    case 'none':
      return ''
    case 'text':
      // Post-hoc addition (2026-08-24, user-reported: selecting a table/
      // chart/SmartArt "doesn't appear under selection") - these objects'
      // own selection.Text is empty or a placeholder character, so this
      // must be checked before the ctx.fullText emptiness check below, or
      // an object selection would always fall through to "no selection".
      if (ctx.objectKind) {
        const readTool = ctx.objectKind === 'table' ? 'read_table' : ctx.objectKind === 'chart' ? 'read_chart' : 'read_smartart'
        const indexField = ctx.objectKind === 'table' ? 'tableIndex' : ctx.objectKind === 'chart' ? 'chartIndex' : 'smartArtIndex'
        return (
          `The user has selected ${ctx.objectKind} ${ctx.objectIndex} (0-based) in the document. ` +
          `Use ${readTool} {${indexField}:${ctx.objectIndex}} to see its current content before editing it.`
        )
      }
      // Post-hoc fix (2026-08-24, user-reported): previously gave only the
      // text with no addressability, so a request to transform the
      // selection in place (e.g. "translate this paragraph") had no way to
      // target replace_blocks at exactly the selected paragraphs and fell
      // back to insert_content, appending a new paragraph instead of
      // replacing the original. Now states the 0-based paragraph range
      // explicitly, matching FT-2's addressable wording for Excel/PowerPoint.
      if (!ctx.fullText) return ''
      if (ctx.startBlockIndex < 0) return `Content selected by the user:\n${ctx.fullText}`
      return (
        `The user has selected paragraphs [${ctx.startBlockIndex}-${ctx.endBlockIndex}] (0-based, inclusive):\n${ctx.fullText}\n` +
        `To replace or transform this selection in place, use replace_blocks with startIndex:${ctx.startBlockIndex}, endIndex:${ctx.endBlockIndex} rather than insert_content.`
      )
    case 'range': {
      if (ctx.multi) {
        return (
          `The user has selected ${ctx.areaCount} separate areas on ${ctx.sheet} (${ctx.address}), ` +
          `${ctx.cellCount} cells total. Issue one tool call per area.`
        )
      }
      if (ctx.entireColumns || ctx.entireRows) {
        const unit = ctx.entireColumns ? 'columns' : 'rows'
        if (ctx.effectiveAddress) {
          return (
            `The user has selected all of ${unit} ${ctx.address} on ${ctx.sheet}. Only ${ctx.effectiveAddress} contains data ` +
            `(${ctx.effectiveCellCount} cells) - use that bounded range, not the full selection, which exceeds read_range's 2000-cell cap.`
          )
        }
        return `The user has selected all of ${unit} ${ctx.address} on ${ctx.sheet}, which is currently empty.`
      }
      if (ctx.cellCount === 1) {
        return `The user has selected the single cell ${ctx.sheet}!${ctx.address}.`
      }
      return (
        `The user has selected ${ctx.sheet}!${ctx.address} (${ctx.cellCount} cells). ` +
        `Use this address directly with the range tools; call read_range if you need the values.`
      )
    }
    case 'slides':
      return `The user has selected slide${ctx.slideIndexes.length > 1 ? 's' : ''} ${ctx.slideIndexes.join(', ')} (0-based).`
    case 'shapes': {
      const list = ctx.shapeIndexes.map((idx, i) => `${idx} ("${ctx.names[i] ?? ''}")`).join(', ')
      return (
        `The user has selected shape${ctx.shapeIndexes.length > 1 ? 's' : ''} ${list} on slide ${ctx.slideIndex}. ` +
        `These are 0-based indices in the form the tools take (slideIndex, shapeIndex).`
      )
    }
    case 'shapeText':
      return (
        `The user has selected text inside shape ${ctx.shapeIndex} on slide ${ctx.slideIndex}: "${ctx.text}" - ` +
        `the selection is a run within that shape, not the whole shape.`
      )
    case 'mail': {
      if (ctx.count === 1) {
        return (
          `The user has selected the email "${ctx.subject}"` +
          (ctx.senderName ? ` from ${ctx.senderName}` : '') +
          (ctx.folderName ? ` in the ${ctx.folderName} folder` : '') +
          `. Its message_id (Outlook EntryID) is ${ctx.entryIds[0]} - pass that straight to get_email, ` +
          `reply_email, reply_all_email, forward_email, move_email, get_attachment, etc.`
        )
      }
      return (
        `The user has selected ${ctx.count} emails` +
        (ctx.folderName ? ` in the ${ctx.folderName} folder` : '') +
        (ctx.conversationTopic ? ` (conversation: "${ctx.conversationTopic}")` : '') +
        `. Their message_ids (Outlook EntryIDs) are: ${ctx.entryIds.join(', ')}.`
      )
    }
    default:
      return ''
  }
}

function numberToColumnLetter(n: number): string {
  let s = ''
  while (n > 0) {
    const rem = (n - 1) % 26
    s = String.fromCharCode(65 + rem) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}

function columnLetterToNumber(col: string): number {
  let n = 0
  for (let i = 0; i < col.length; i++) n = n * 26 + (col.charCodeAt(i) - 64)
  return n
}

/**
 * FT-2 Task 5: classifies the raw payload into the UI's SelectionExtent -
 * chat-ui.ts owns rendering/localizing the words, this only picks which case
 * applies and extracts the numbers, per Task 5 Step 3 ("a label the shell
 * computes"). Returns null for "no selection" (reverts the pill to its
 * per-app whole-scope label).
 */
function toSelectionScopeUpdate(raw: RawSelectionPayload): { hasSelection: boolean; preview?: string; extent?: SelectionExtent } | null {
  if (!raw.hasSelection) return null
  if (raw.app === 'excel') {
    if (raw.multi) {
      return { hasSelection: true, extent: { kind: 'multiArea', areaCount: raw.areaCount ?? 2, cellCount: raw.cellCount ?? 0 } }
    }
    // Task 2b's `effectiveRows`/`effectiveCellCount` are only present when the
    // selection actually needed intersecting with UsedRange (entire column/
    // row, or >10k cells) - null here means "not computed", not "computed and
    // empty"; an empty intersection is `effectiveAddress === null` with
    // `effectiveRows` unset, both rendered as the empty case in chat-ui.ts.
    if (raw.entireColumns) {
      const dataRows = raw.effectiveAddress ? raw.effectiveRows ?? null : null
      if ((raw.cols ?? 1) === 1) {
        return { hasSelection: true, extent: { kind: 'wholeColumn', col: raw.firstCol ?? '', dataRows } }
      }
      const lastCol = raw.firstCol && raw.cols ? numberToColumnLetter(columnLetterToNumber(raw.firstCol) + raw.cols - 1) : (raw.firstCol ?? '')
      return {
        hasSelection: true,
        extent: { kind: 'wholeColumns', firstCol: raw.firstCol ?? '', lastCol, cols: raw.cols ?? 0, dataRows },
      }
    }
    if (raw.entireRows) {
      return { hasSelection: true, extent: { kind: 'wholeRow', row: raw.firstRow ?? 0 } }
    }
    if ((raw.cellCount ?? 0) === 1) {
      return { hasSelection: true, extent: { kind: 'cell', address: raw.address ?? '' } }
    }
    return { hasSelection: true, extent: { kind: 'range', address: raw.address ?? '', rows: raw.rows ?? 0, cols: raw.cols ?? 0 } }
  }
  if (raw.app === 'outlook') {
    const count = (raw.entryIds ?? []).length
    if (count === 0) return null
    return { hasSelection: true, extent: { kind: 'mailSelection', count, subject: raw.subject ?? '' } }
  }
  if (raw.app === 'powerpoint') {
    if (raw.selKind === 'slides') {
      return { hasSelection: true, extent: { kind: 'slides', count: (raw.slideIndexes ?? []).length } }
    }
    if (raw.selKind === 'shapes') {
      const names = raw.names ?? []
      return { hasSelection: true, extent: { kind: 'shapes', count: names.length, primaryName: names.length === 1 ? names[0] ?? null : null } }
    }
    if (raw.selKind === 'shapeText') {
      return { hasSelection: true, extent: { kind: 'shapeText', shapeName: null } }
    }
    return null
  }
  // Word (post-hoc addition, 2026-08-24): a table/chart/SmartArt selection
  // renders as a proper extent pill (e.g. "Table 2") instead of the
  // quoted-text form, which would otherwise show empty/placeholder text for
  // these object kinds - the exact "no pointer" gap the user reported.
  if (raw.objectKind) {
    return { hasSelection: true, extent: { kind: 'wordObject', objectKind: raw.objectKind, objectIndex: raw.objectIndex ?? 0 } }
  }
  // Word: the existing quoted-preview form, capped by the C# side already.
  return { hasSelection: true, preview: raw.preview ?? '' }
}

export interface AddInConfig {
  /** every tool this app implements */
  tools: AgentSkill['tools']
  /** FT-1 Task 5: localized settings-screen label/description, keyed by tool name. */
  toolDisplay: Record<string, ToolDisplayInfo>
  systemPrompt: string
  skillId: string
  starters: Array<{ en: string; he: string }>
  /** tool names available in Read only mode */
  readOnlyTools: string[]
  /** additionally available in Comment only mode (Word's add_comment; empty/absent elsewhere) */
  commentOnlyExtraTools?: string[]
  /** inject the user's current selection into per-turn context (Word, Excel, PowerPoint - FT-2) */
  useSelectionContext?: boolean
  /** FT-2 Task 4 Step 2: overrides the default per-kind context sentence, if an app ever needs different wording. */
  describeSelection?: (ctx: SelectionContext) => string
  /** FT-2 Task 5: the scope-hint pill's "no selection" wording - defaults to 'doc' (Word). */
  scopeUnit?: 'doc' | 'sheet' | 'deck' | 'mailbox'
  /**
   * Restricts the editing-mode menu to this subset, in this order. Defaults to
   * all four modes (Word/Excel/PowerPoint). Outlook passes
   * ['readOnly', 'fullAutonomy'] - Comment only / Track changes have no meaning
   * for mail.
   */
  availableModes?: EditingMode[]
}

/**
 * Boots one add-in's chat panel: WebView2 bridge, settings, transport,
 * chat-UI mount, and AgentLoop event plumbing. Everything here was
 * previously duplicated near-verbatim across WordAiAddIn/ExcelAiAddIn/
 * PowerPointAiAddIn's entry.ts (PP-0) - each app now supplies only what is
 * genuinely app-specific through `config`.
 */
export function startAddIn(config: AddInConfig): void {
  // Task 11 (Word): editing-mode control. Client-side filtering only (first
  // line of defense - smaller prompts, fewer wasted turns); the real
  // enforcement is server-side in each app's *Tools.Execute, which gates
  // mutating tool calls even if the model somehow requests one that wasn't
  // offered here.
  let editingMode: EditingMode = 'fullAutonomy'

  const readOnlySet = new Set(config.readOnlyTools)
  const commentOnlySet = new Set([...config.readOnlyTools, ...(config.commentOnlyExtraTools ?? [])])

  function availableForMode(): string[] {
    if (editingMode === 'readOnly') return config.tools.filter((t) => readOnlySet.has(t.name)).map((t) => t.name)
    if (editingMode === 'commentOnly') return config.tools.filter((t) => commentOnlySet.has(t.name)).map((t) => t.name)
    return config.tools.map((t) => t.name)
  }

  // FT-1 Task 4: `null` means "no override - use the scope's full default set".
  // A non-null Set is the user's live registration override, reset to null on
  // every mode change (Step 2) so a stale override can never survive a scope
  // switch that would make it stale or, worse, too permissive.
  let registeredTools: Set<string> | null = null

  function activeTools(): AgentSkill['tools'] {
    const available = new Set(availableForMode())
    // Intersect with `available` even when an override is present - Task 4
    // Step 3: this is what keeps an out-of-scope tool from ever reaching the
    // model even if the override set is stale or was tampered with.
    return config.tools.filter((t) => available.has(t.name) && (registeredTools === null || registeredTools.has(t.name)))
  }

  // FT-1 Task 5: build the settings-screen tool list from the app's raw
  // toolDisplay map, so entry.ts only needs to supply {label, description}
  // per tool name - one shared place decides the fallback-on-drift behavior
  // (Task 5 Step 4) instead of every app reimplementing it.
  const toolDisplayList: ToolDisplayEntry[] = config.tools.map((t) => {
    const entry = config.toolDisplay[t.name]
    if (!entry) {
      // eslint-disable-next-line no-console
      console.warn(`[app-shell] tool "${t.name}" has no toolDisplay entry - falling back to its raw schema name`)
      return { name: t.name, label: { en: t.name, he: t.name }, description: { en: '', he: '' } }
    }
    return { name: t.name, label: entry.label, description: entry.description }
  })

  // Task 12 (Word)/FT-2 Task 4: the current selection, updated live from the
  // .NET-side selection-change handler and read by buildContext() below.
  let latestSelectionContext: SelectionContext = { kind: 'none' }
  const describeSelectionFn = config.describeSelection ?? defaultDescribeSelection

  // FT-1 Task 8: the document guidelines message. `savedDocMessage` is
  // whatever is currently persisted on disk (kept in sync by the bridge's
  // doc-settings-loaded response and by a successful Save); `activeDocMessage`
  // is what actually gets injected into the system prompt this conversation -
  // frozen by beginConversation() at conversation-start boundaries only
  // (initial load, New chat), never read live per-turn, so editing the
  // guidelines mid-conversation cannot retroactively change a run in progress.
  let savedDocMessage = ''
  let activeDocMessage = ''
  function beginConversation(): void {
    activeDocMessage = savedDocMessage
  }

  const skill: AgentSkill = {
    id: config.skillId,
    systemPrompt: config.systemPrompt,
    // Live getter (not a fixed array): AgentLoop.startTurn() reads
    // this.options.skill.tools fresh every turn (see
    // shared/web-src/agent-core/loop.ts), so this recomputes the tool list
    // per-turn from the current editingMode without needing to rebuild the
    // whole skill object or touch agent-core.
    get tools() {
      return activeTools()
    },
    ...(config.useSelectionContext
      ? {
          // FT-2 Task 4 Step 3: read once per run (AgentLoop.run() calls
          // buildContext() once, not per-turn) - a selection that changes
          // mid-run must not leak into a later turn of the same run. This is
          // deliberately NOT systemSuffix (FT-1's document guidelines, which
          // IS read every turn).
          buildContext: () => describeSelectionFn(latestSelectionContext),
        }
      : {}),
    executeTool: (call) => callDotNetTool(call.name, call.input),
  }

  const root = document.getElementById('root')!
  const ui = mountChatUI(root, {
    starters: config.starters,
    tools: toolDisplayList,
    scopeUnit: config.scopeUnit,
    modes: config.availableModes,
    onToolRegistrationChange: (registered) => {
      registeredTools = new Set(registered)
    },
    onCollapseChange: (collapsed) => postCollapse(collapsed),
    initialSettings: (() => {
      const s = getSettings()
      const slot = s.ai.providers[s.ai.provider]
      return {
        provider: s.ai.provider,
        apiKey: slot.apiKey,
        model: slot.model,
        baseUrl: slot.baseUrl,
        skipTlsVerify: s.skipTlsVerify,
        providers: s.ai.providers,
      }
    })(),
    // Post-hoc addition (2026-08-24, user-requested): AgentLoop.cancel()
    // already existed (its own comment even anticipated "when the user
    // clicks stop") but was never reachable from any UI control until now -
    // wired here, same forward-reference-via-closure pattern onSend already
    // uses for `loop` below (declared further down this file).
    onStop: () => loop.cancel(),
    onSend: (text) => {
      // Post-hoc change (2026-08-24, user-requested): previously a no-op
      // while busy (the textarea used to be disabled too, so this was
      // unreachable anyway). Now the textarea stays enabled during a run,
      // so a send while busy queues the message instead of dropping it -
      // dispatched automatically once the current run finishes (onDone
      // below), whether it finished normally or was stopped. The user's
      // message is shown and persisted immediately (chronologically
      // accurate - they sent it now), only the actual model run is
      // deferred; `pendingQueuedText` is declared further down this file
      // (same forward-reference-via-closure pattern already used for `loop`).
      if (loop.busy) {
        pendingQueuedText = text
        ui.addUserMessage(text)
        persistMessage('user', text)
        return
      }
      ui.addUserMessage(text)
      // Arms for a new assistant turn - does NOT append a bubble (PP-2:
      // chat-ui.ts creates one lazily on the first text delta, so a turn's
      // tool-call group lands above the text it produced). Do not move this
      // call to "fix" a perceived missing bubble.
      ui.beginAssistantMessage()
      ui.setBusy(true)
      persistMessage('user', text)
      loop.run(text)
    },
    onNewChat: () => {
      postNewChatDivider()
      loop.reset()
      beginConversation()
      ui.resetToEmpty()
    },
    onModeChange: (mode: EditingMode) => {
      editingMode = mode
      // Task 4 Step 2: reset the override BEFORE pushing the new scope down,
      // so the UI never briefly shows the old (possibly narrower or wider)
      // registration against the new scope's tool set.
      registeredTools = null
      postMode(mode)
      ui.setToolScope(availableForMode(), availableForMode())
    },
    onSettingsSave: (settings) => {
      const current = getSettings()
      const providers = {
        ...current.ai.providers,
        [settings.provider]: {
          apiKey: settings.apiKey || current.ai.providers[settings.provider]?.apiKey || '',
          model: settings.model || current.ai.providers[settings.provider]?.model || '',
          baseUrl: settings.baseUrl || current.ai.providers[settings.provider]?.baseUrl,
        },
      }
      setSettings({ ai: { provider: settings.provider, providers }, skipTlsVerify: settings.skipTlsVerify })
      postTlsBypass(settings.skipTlsVerify)
      // Task 9: registration itself already took effect live via
      // onToolRegistrationChange above - settings.registeredTools is an echo,
      // not applied here again. The doc message, however, is Save-gated (Task
      // 8 Step 5): persist it, but do NOT call beginConversation() - it must
      // not retroactively affect the conversation already in progress.
      if (settings.docSystemMessage !== undefined) {
        savedDocMessage = settings.docSystemMessage
        saveDocSettings(settings.docSystemMessage)
      }
    },
    // Tries the form's CURRENT (possibly unsaved) values, not getSettings() -
    // otherwise "Test connection" would silently test the last-saved config
    // instead of what the user just typed. A one-word prompt, no tools, and a
    // tiny token budget keep this cheap; the first delta proves the round
    // trip works without waiting for a full reply.
    onSettingsTest: (settings) => {
      return new Promise<string>((resolve, reject) => {
        const controller = new AbortController()
        const config: AiProviderConfig = { apiKey: settings.apiKey, model: settings.model, baseUrl: settings.baseUrl }
        let gotDelta = false
        streamForProvider(
          settings.provider,
          config,
          '',
          [{ role: 'user', text: 'Say "ok".' }],
          [],
          16,
          {
            onDelta: () => {
              if (gotDelta) return
              gotDelta = true
              controller.abort()
              resolve('Connection OK.')
            },
            onToolCall: () => {},
            signal: controller.signal,
          },
        )
          .then(() => {
            if (!gotDelta) resolve('Connection OK.')
          })
          .catch((e: unknown) => {
            // An abort triggered by our own success path above rejects too -
            // do not surface that as a failure.
            if (gotDelta) return
            reject(e instanceof Error ? e : new Error(String(e)))
          })
      })
    },
  })

  // Apply the persisted TLS-bypass preference on load too, not just after a
  // future Save - otherwise a user who enabled it last session would silently
  // go back to strict verification every time they reopen the document.
  postTlsBypass(getSettings().skipTlsVerify)

  // Post-hoc addition (2026-08-24, user-requested): a message sent while a
  // run is already busy is queued here rather than dropped, and dispatched
  // in onDone below once the current run finishes (normally or via stop).
  let pendingQueuedText: string | null = null

  let currentToolGroup: ReturnType<typeof ui.beginToolGroup> | null = null
  // Post-hoc fix (2026-08-24, user-reported): loop.ts's "turn" is one model
  // response - for a model that chains several tool calls back-to-back with
  // no text between them, that is often exactly one tool call per turn, so
  // closing/nulling currentToolGroup on every onTurnEnd split a single
  // logical batch into a separate "Ran 1 tool" box per call instead of one
  // group incrementing to "Ran N tools". The group should only actually
  // close when text has genuinely streamed since it opened (so a LATER
  // block of tools still gets its own group below that text, preserving
  // chronological order) - tracked here since only bootstrap.ts sees both
  // onText and onTurnEnd.
  let textStreamedSinceGroup = false
  const activeSteps = new Map<string, ReturnType<ReturnType<typeof ui.beginToolGroup>['addStep']>>()
  function closeToolGroup(): void {
    currentToolGroup?.end()
    currentToolGroup = null
    textStreamedSinceGroup = false
  }

  // Sends a continuation instruction rather than re-running the user's
  // original prompt: the run may already have applied mutating tools to the
  // document, and replaying the prompt would apply them a second time.
  const CONTINUE_INSTRUCTION =
    'Your previous reply was cut off by the length limit. Continue exactly where it stopped. ' +
    'Do not repeat what you already wrote and do not re-apply any edits you already made.'

  function continueRun(): void {
    if (loop.busy) return
    ui.beginAssistantMessage()
    ui.setBusy(true)
    // Not persisted to ChatStore as a user message - it is machine-generated
    // plumbing, not something the user typed, and would be confusing in
    // restored history.
    loop.run(CONTINUE_INSTRUCTION)
  }

  // Post-hoc addition (2026-08-24, user-requested): dispatches a message
  // queued via onSend while the previous run was busy - the user bubble and
  // ChatStore persistence already happened at queue time, so this only
  // needs to actually start the model run. Called from onDone/onError below
  // regardless of how the prior run ended (finished normally, truncated, or
  // stopped via the stop button) - a queued message should still go out.
  function dispatchQueuedIfAny(): void {
    if (pendingQueuedText === null) return
    const queued = pendingQueuedText
    pendingQueuedText = null
    ui.beginAssistantMessage()
    ui.setBusy(true)
    loop.run(queued)
  }

  const loop = new AgentLoop({
    transport: makeTransport(),
    skill,
    // Task 8 Step 2/4: appended to the system prompt every turn, but the
    // value it reads (activeDocMessage) only changes at conversation-start
    // boundaries - see beginConversation() above. Labeled explicitly as
    // user-supplied so the model treats it as standing instructions from the
    // user, not as system policy.
    systemSuffix: () => (activeDocMessage ? '\n\nDocument guidelines from the user:\n' + activeDocMessage : ''),
    events: {
      onText: (text) => {
        textStreamedSinceGroup = true
        ui.updateAssistantMessage(text)
      },
      onToolStart: (call) => {
        // If text streamed into the current group's turn, this new tool
        // belongs in a FRESH group below that text - not appended into the
        // still-open group above it (which renders it out of chronological
        // order and leaves the midway reply unsealed, so a later turn's
        // streamed text overwrites it). Back-to-back tool-only turns (no text
        // between) still share one incrementing group.
        if (currentToolGroup && textStreamedSinceGroup) closeToolGroup()
        if (!currentToolGroup) currentToolGroup = ui.beginToolGroup()
        activeSteps.set(call.id, currentToolGroup.addStep(call.name, call.input))
      },
      onToolExecuted: (event) => {
        activeSteps.get(event.call.id)?.complete({
          output: event.execution.output,
          isError: event.execution.isError,
          mutated: event.execution.mutated,
        })
        activeSteps.delete(event.call.id)
      },
      onTurnEnd: () => {
        // Only close here if text genuinely streamed since the group
        // opened - otherwise this was a back-to-back tool-only turn and the
        // group should stay open so the next turn's tools keep incrementing
        // the same count instead of starting a new box.
        if (textStreamedSinceGroup) closeToolGroup()
        // Freeze any midway reply as its own sealed bubble and bring the
        // thinking indicator back for the model-reasoning gap before the next
        // turn (onTurnEnd only fires for turns that called tools, so there is
        // always a next turn).
        ui.endTurn()
      },
      onDone: (result) => {
        closeToolGroup()
        const hasText = result.text.length > 0
        const finalText = hasText ? result.text : ui.translate('emptyReply')
        ui.endAssistantMessage(finalText)
        // A truncated turn that DID stream partial text keeps that text (set
        // above) and gets the notice appended below it, rather than being
        // treated as empty - `result.truncated`/`result.turnLimit` are
        // orthogonal to whether any text came through.
        if (result.truncated) ui.showNotice('truncated', () => continueRun())
        else if (result.turnLimit) ui.showNotice('turnLimit')
        ui.setBusy(false)
        // Persist only the model's actual text (or the localized empty
        // marker) - never the notice, which is UI state about this run, not
        // conversation content.
        persistMessage('assistant', finalText)
        dispatchQueuedIfAny()
      },
      onError: (error) => {
        closeToolGroup()
        const placeholder = `[Error: ${error}]`
        ui.endAssistantMessage(placeholder)
        persistMessage('assistant', placeholder)
        ui.showError(error)
        ui.setBusy(false)
        dispatchQueuedIfAny()
      },
    },
  })

  initBridge({
    onHistoryLoaded: (messages) => {
      ui.showHistoric(messages)
      loop.restore(messages.map((m) => ({ role: m.role, text: m.text })))
    },
    onSelectionChanged: (raw) => {
      latestSelectionContext = toSelectionContext(raw)
      ui.setSelectionScope(toSelectionScopeUpdate(raw))
    },
    // Task 8: "initial load" freezes the very first conversation's doc
    // message the moment it arrives (the pane's history/mode restore is
    // otherwise async too, so there is no earlier well-defined point to do
    // this at).
    onDocSettingsLoaded: (systemMessage) => {
      savedDocMessage = systemMessage
      ui.setDocSystemMessage(systemMessage)
      beginConversation()
    },
  })

  // Task 4: push the initial (scope-default, no override) tool registration
  // down to the settings view before the user ever opens it.
  ui.setToolScope(availableForMode(), availableForMode())

  requestHistory()
  requestDocSettings()
}
