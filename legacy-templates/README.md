# Legacy templates

Retired document templates, kept for reference. They are **not** bundled with the app (not in `electron-builder.yml`)
and nothing in `src/`, `core/` or `electron/` loads them. A few tests read them as fixtures.

| File | Was |
| --- | --- |
| `work-order-template.docx` | Work Order type 1 |
| `work-order-template-kompally.docx` | Work Order type 2 |
| `file-backer-template.docx` | File Backer type 1 |
| `file-backer-template-header2.docx` | File Backer type 2 |
| `civil-tender-template.docx` | Tender Document type 1 |
| `civil-tender-template-header2.docx` | Tender Document type 2 |
| `intimation-template.docx` | Intimation template 1 |
| `qc-intimation-completion-template.docx` | QC Intimation/Completion type 1 |
| `issue-notice-template.docx`, `issue-notice-ee-template.docx` | Notice template 1 (SE / EE) |
| `zonal-memo-ee-template.docx` | Memo to EE type 1 (SE office) |
| `zonal-concluding-agreement-template.docx`, `zonal-concluding-agreement-template-header2.docx` | Concluding Agreement type 1 / type 3 (SE office) |

To restore one: move it back into `resources/`, re-add it to `electron-builder.yml`, and re-add its entry in
`core/workOrderTemplateVariants.ts`.
