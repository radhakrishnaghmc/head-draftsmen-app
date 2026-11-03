import type { DocFieldGroup } from './components/DocFieldsEditor'

/** The DB Note's editable fields, grouped and ordered the way the 28 numbered points print — shared by the Agreement workspace and the DB Note board. */
export const DB_NOTE_FIELDS: DocFieldGroup[] = [
  {
    title: 'The work',
    fields: [
      { label: 'Wincode', title: 'Wincode' },
      { label: 'Corporation', title: 'Corporation' },
      { label: 'CNO', title: 'Circle No' },
      { label: 'Name of the work', title: '1. Name of the work', multiline: true },
      { label: 'Name of the agency', title: '2. Name of the agency' },
      { label: 'Serial of bill', title: '3. Serial of bill' }
    ]
  },
  {
    title: 'Amounts & agreement',
    fields: [
      { label: 'Estimate Amount', title: '4. Estimate amount' },
      { label: 'Administrative Sanction Amount', title: '5. Administrative sanction' },
      { label: 'ECV', title: '6. ECV' },
      { label: 'Tender Percentage', title: '7. Tender percentage' },
      { label: 'Contract Amount', title: '8. Agreement value' },
      { label: 'Agreement No and Date', title: '9. Agreement no. & date' },
      { label: 'Period of Contract', title: '10. Period of contract', periodUnit: true }
    ]
  },
  {
    title: 'Progress & measurements',
    fields: [
      { label: 'Date of Completion', title: '12. Date of completion' },
      { label: 'EOT Details', title: '13. Extension of time' },
      { label: 'AE Recording Dates', title: '14. Recording dates (AE)', multiline: true },
      { label: 'MB Details', title: '15. M.B. details' },
      { label: 'DyEE Check Dates', title: '16. Dy.EE check measurement', multiline: true },
      { label: 'EE Test Check', title: '17. EE test check' },
      { label: 'SE Super Check', title: '18. SE super check' }
    ]
  },
  {
    title: 'Deductions & reports',
    fields: [
      { label: 'Embedded Tax', title: '19. Embedded tax' },
      { label: 'Price Escalation', title: '20. Price escalation' },
      { label: 'EOT Penalty', title: '21. EOT penalty' },
      { label: 'DC ZC Inspection', title: '22. DC / ZC visual inspection' },
      { label: 'QC Officials Report', title: '23. QC officials report' },
      { label: 'QCC Third Party Report', title: '24. QCC third party report' },
      { label: 'ATR Details', title: '25. A.T.R. details' },
      { label: 'Seignorage Details', title: '26. Seignorage charges' },
      { label: 'One Time Penalty Details', title: '27. One time penalty / DMFT / SMET' }
    ]
  },
  {
    title: 'Recommended payment',
    fields: [{ label: 'Recommended Amount', title: '28. Recommended for payment (Rs.)' }]
  }
]
