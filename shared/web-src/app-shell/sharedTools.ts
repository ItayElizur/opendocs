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

// The skill tools (skills.ts, OfficeAi.Shared/SkillCatalog.cs). bootstrap.ts adds them itself, only when the app has
// skills, and counts them read-only in every mode; they are in no app's entry.ts or tier list.
export const SKILL_TOOLS = [
  {
    name: 'load_skill',
    description:
      "Loads a skill's instructions by name. Skills are listed in the system prompt. Call it when the request matches a " +
      "skill's description, before doing the task, then follow the instructions it returns. Do not load a skill that is already loaded in this conversation.",
    inputSchema: {
      type: 'object',
      properties: { name: { type: 'string', description: 'The skill name, exactly as listed.' } },
      required: ['name'],
    },
  },
  {
    name: 'read_skill_file',
    description:
      "Reads one of a skill's other files (references, templates), listed when the skill loads. Text files are returned as text. " +
      "For a binary file (such as an image) it returns the file's full path, to pass to a tool that takes a file path.",
    inputSchema: {
      type: 'object',
      properties: {
        skill: { type: 'string', description: 'The skill name.' },
        path: { type: 'string', description: "The file's path inside the skill, as listed, e.g. references/style.md." },
      },
      required: ['skill', 'path'],
    },
  },
]

export const SKILL_TOOL_DISPLAY = {
  load_skill: {
    label: { en: 'Load skill', he: 'טעינת מיומנות' },
    description: { en: "Reads a skill's instructions.", he: 'קורא את ההוראות של מיומנות.' },
  },
  read_skill_file: {
    label: { en: 'Read skill file', he: 'קריאת קובץ של מיומנות' },
    description: { en: 'Reads a reference or template file that belongs to a skill.', he: 'קורא קובץ עזר או תבנית ששייכים למיומנות.' },
  },
}
