/**
 * A placeholder whose resolved value is this marker removes its whole paragraph
 * from the filled document (e.g. an Eligibility Criteria staffing line whose
 * count is entered as 0), instead of printing the line with a "0". Applied by
 * fillPlaceholdersInAllParts in core/docx-edit.ts.
 */
export const DROP_PARAGRAPH_MARKER = '[[DROP-LINE]]'
