// The standard BOQ column layout — fixed, matching the bundled BOQ template
// (resources/boq-template.xlsx), so shared between the renderer (which
// builds the plain preview table) and the main process (which fills the
// real template file).
export const BOQ_HEADERS: string[] = [
  'Estimate Quantity (only Figures)',
  'Item Detailed \nSpecification Description',
  'Work Type \neg. Earth Work, Electrical works.. etc\n( upto 200 Characters)',
  'Item Short Description \n( upto 100 Characters)',
  'APSS / Morth Cl. Number \n( upto 200 Characters)',
  'Rate (INR) \n( Upto 2 Decimals )',
  'UOM\n( upto 50 Characters)',
  'Amount (INR) \n( Upto 2 Decimals )'
]
