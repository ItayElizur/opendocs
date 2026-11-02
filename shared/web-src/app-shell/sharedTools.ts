// Tools defined once and added to the Word, Excel and PowerPoint entry.ts files as `...SHARED_TOOLS` /
// `...SHARED_TOOL_DISPLAY`. A new shared tool is one more entry in each of the two exports below; the apps do not change.
// Shared tools are not added to any app's readOnlyTools / commentOnlyExtraTools: unless a tool is read-only it is
// advertised from Track Changes up, which is what `save` needs.

const SAVE_TOOL = {
  name: 'save',
  description:
    'Saves the open file. With saveAs=false (the default) it saves in place and overwrites the current file; if the file has never been saved, ' +
    "the app's own Save As dialog opens instead. With saveAs=true it opens the Save As dialog so the user can pick a name, folder or format. " +
    'If the user wants a copy, a new file name, a different location or a different format, you MUST set saveAs=true. ' +
    'The dialog waits for the user, and the call returns only after they save or cancel (a cancel is reported, not an error). ' +
    'Call this only when the user asks to save.',
  inputSchema: {
    type: 'object',
    properties: {
      saveAs: {
        type: 'boolean',
        description:
          'Default false. Set to true whenever the user wants a copy, a new file name, a different location or a different format; false overwrites the current file.',
      },
    },
  },
}

const SAVE_TOOL_DISPLAY = {
  label: { en: 'Save', he: 'שמירה' },
  description: {
    en: 'Saves the file, or opens Save As when a copy is wanted.',
    he: 'שומר את הקובץ, או פותח את "שמירה בשם" כשצריך עותק.',
  },
}

export const SHARED_TOOLS = [SAVE_TOOL]

export const SHARED_TOOL_DISPLAY = { save: SAVE_TOOL_DISPLAY }
