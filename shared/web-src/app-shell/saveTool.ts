// Shared `save` tool schema + display info, imported by the Word, Excel and PowerPoint entry.ts files.
// Advertised from Track Changes up: it is deliberately in neither readOnlyTools nor commentOnlyExtraTools.
export const SAVE_TOOL = {
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

export const SAVE_TOOL_DISPLAY = {
  label: { en: 'Save', he: 'שמירה' },
  description: {
    en: 'Saves the file, or opens Save As when a copy is wanted.',
    he: 'שומר את הקובץ, או פותח את "שמירה בשם" כשצריך עותק.',
  },
}
