# Agent Brief: Grand Group FZC Procurement Document System

## 1. Purpose
Grand Group FZC is an intermediary supplier. It receives RFQs/tenders from end users, sources supplier quotes, issues Quotations (QT) and Proforma Invoices (PI) to customers, issues Purchase Orders (PO) to suppliers, and then issues Packing Lists (PL), Delivery Notes (DN) and Commercial Invoices (CI) for shipment, delivery and payment settlement.

The system is a set of Excel workbooks on the shared drive `\\Server\Procurement\`. It is built for two users working at the same time. Item data is typed once and every other document is generated from it. Document numbers are sequential and never duplicated.

## 2. Files
- `GG_Job_Template.xlsx`: the layout (tabs, formulas, input cells, HOW TO USE sheet). Intended final form is `GG_Job_Template.xltm` (macro-enabled template, read-only). The VBA is not yet imported into it, and the .xltm has not been saved.
- `Master_Registry.xlsx`: the only source of document numbers. Seeded with `PI-GG-042-2026`; the next number is 43. Sheets: REGISTRY, PREFIXES.
- `vba/modConfig.bas`, `vba/modRegistry.bas`, `vba/modIssue.bas`: VBA modules. Import with Alt+F11, File, Import File.
- `scripts/build_templates.py`: regenerates both workbooks. Edit the script, not the xlsx, when the layout changes.
- Two packaged copies exist: `procurement-doc-automator/` (network share, lock folder) and `procurement-m365/` (SharePoint/OneDrive, check-out lock). Use the m365 version for SharePoint.
- Colour variants: `procurement-styles/` (v1), `-v2/`, `-v3/` (latest). Same cell content as the template; only formatting differs.

## 3. Core rules
1. One entry point per job. Items are typed on PROFORMA INVOICE rows 17-36 (Requested Description, Offered Description, UoM, Qty). Supplier price, box count and box size are typed on PRICE CALC on the same rows.
2. Numbers are claimed from Master_Registry, never typed or calculated. Format `PREFIX-GG-SEQ-YEAR`, e.g. `PI-GG-043-2026`. One sequence is shared by all types. Prefixes: PI, QT, PO, CI, PL, DN.
3. On issue, the DOCUMENT NO. cell becomes a hard-coded value, not a formula.
4. The template is never edited. Users open a fresh copy and save it as a job file (.xlsm).
5. Issued files are values-only .xlsx, with no macros and no controls block.
6. Each tab can be issued once. Re-issue means a new job.
7. PRICE CALC is internal. It holds supplier cost. Never send it or the job file to a customer.

## 4. Workflow and tabs
| Stage | Tab | Action |
|---|---|---|
| RFQ and supplier quote | PROFORMA INVOICE + PRICE CALC | Enter descriptions, UoM, Qty; supplier price, boxes, box size, kg per box |
| Customer quote | QUOTATION | Issue; QT number claimed |
| Customer PO, PI | PROFORMA INVOICE | Issue; PI number claimed; values frozen |
| Our PO to supplier | SUPPLIER PO | Enter Unit Cost per line (linked from PRICE CALC); issue; PO number |
| Shipment | PACKING LIST | Enter cartons, net/gross kg per line, ship date; issue; PL number |
| Delivery | DELIVERY NOTE | Enter delivery date, transport ref., received-by; issue; DN number |
| Payment | COMMERCIAL INVOICE | Issue; CI number; prices, VAT and freight from PI |

Supplier PL and DN are received documents. Ours are generated. Do not copy supplier cartons or weights into ours without checking.

## 5. Column schemas (strict)
- Quotation, PI, CI (pricing tables): `No.` | `Requested Description` | `Offered Description` | `UoM` | `Qty` | `Unit Price` | `Total Price`
- Packing List and Delivery Note: `No.` | `Requested Description` | `Offered Description` | `UoM` | `Qty` | `Carton No.` | `Net Wt (KG)` | `Gross Wt (KG)`
- Supplier PO: first five columns, then `Unit Cost` | `Total Cost`
- Requested Description = the end user's wording. Offered Description = the supplier's wording, which is what we sell.

## 6. Linking and mapping
- Item rows 17-36 are fixed on every tab. Never sort or insert rows on one tab only.
- Linked cells use `=IF(src="","",src)` so blanks show as blank, not 0.
- PI Unit Price is a formula from PRICE CALC; it is never typed.
- The TO (end user) is chosen from a dropdown on the PI (cell C8 in the template), fed by the end-user list on CONTROLS. Choosing OTHER lets the user type a name.
- Currency, incoterm and destination are on the PI controls block (K:L) and drive labels on the other tabs.

## 7. Pricing, VAT and freight (current template)
- Unit price = supplier cost x (1 + %). Example: 100 becomes 110, 115, 118 or 120 at 10%, 15%, 18%, 20%. This is markup on cost, not margin division. The user confirmed this.
- Scenarios 10/15/18/20% are shown side by side on PRICE CALC. E3 selects the % that feeds PI Unit Price (col R) and the PI freight line.
- Freight = the HIGHER of (total CBM x rate per CBM) or (total gross kg x rate per kg). CBM = boxes x L x W x H (cm) / 1,000,000.
- Documentation charge (default 10%) is added to freight first, then the selected % is added on top.
- SUBTOTAL = sum of Qty x Unit Price. VAT (default 5%, PI controls L5) is shown as an amount on the goods subtotal only.
- Freight is always a separate line after VAT. Not spread into unit prices.
- GRAND TOTAL = SUBTOTAL + VAT + FREIGHT.
- Set VAT to 0% in L5 for zero-rated jobs.

## 8. Concurrency design
- The registry is the only shared write target. Job files are per user.
- Layer 1, lock: a lock folder (network share version: `MkDir` on `Master_Registry.lock` with `owner.txt`; m365 version: SharePoint check-out). A lock older than 180 s is treated as abandoned.
- Layer 2, open-for-write retry: `Workbooks.Open` retried while the registry is read-only.
- Retry: 60 attempts, about 0.5 s plus jitter.
- Claim order: lock, open registry, read MAX(sequence), duplicate check, append row, save, close, unlock. The number is returned only after the save succeeds.
- If a number is reserved but the file is not saved, the number is marked VOID in Status and a new number is taken. Gaps are acceptable; duplicates are not.

## 9. VBA standards
- Windows desktop VBA only. No Office Scripts or SharePoint-only features in the logic.
- `Option Explicit` in every module. Constants in modConfig (retries, lock age, row bounds).
- Explicit error handling in each public procedure, with a `Done:` block that restores DisplayAlerts, ScreenUpdating and locks.
- Read server paths from named ranges on CONTROLS (`ctl_RegistryPath`, `ctl_IssuedFolder`). Never hard-code them.
- Validate before claiming a number.
- Use `Sleep` (kernel32, `#If VBA7`) for lock waits, not `Application.Wait`.
- Modules: modConfig (constants, CtlValue, WaitWithJitter); modRegistry (ClaimDocumentNumber and lock helpers); modIssue (IssueActiveSheet, IssueDocument, validation, ISSUE LOG).
- Status: the VBA has never been compiled or run in Excel. Test it before live use.

## 10. Setup (admin, once)
1. Create `\\Server\Procurement\Registry\`, `Issued\`, `Jobs\`, `Templates\`. Users need Read and Create; restrict delete on the registry.
2. Place `Master_Registry.xlsx` in `Registry\` and keep it closed.
3. Open the template, import the three .bas modules.
4. On each of the six document tabs (not PRICE CALC), add a button assigned to `IssueActiveSheet`, labelled "Issue this document".
5. Check CONTROLS paths.
6. Save As `Excel Macro-Enabled Template (*.xltm)` into `Templates\`, set read-only.
7. Test with a dummy customer: issue QT, then PI; the registry should show 43 and 44.

## 11. Troubleshooting
- "Registry is busy": another user is claiming. Wait. If it persists for minutes, check for a stale lock with an old owner.txt; delete only after confirming nobody is issuing.
- "Registry open read-only": someone has it open. Ask them to close it.
- "Cannot issue yet: row N": fix the named row.
- Issue failed after reservation: mark the number VOID and re-issue. Never reuse.
- Phantom zeros: the link formula lost its IF wrapper.
- Wrong totals: check items are on rows 17-36 and freight inputs are numbers, not text.

## 12. Known inconsistencies and open items (read before acting)
- SKILL.md (procurement-doc-automator) still describes an OLDER pricing model: margin basis cost/(1-m), PRICE CALC column V, PI!L7 for VAT, and an end-user list of 8 names. The current template uses cost x (1+%), PRICE CALC column R, PI!L5 for VAT, and a 9-name list. Treat HOW TO USE and the template as authoritative until SKILL.md is updated.
- End-user list (confirmed by the user): AIAA (Asmara International Airport), Massawa Port, MOT (Ministry of Transport), MOA (Ministry of Agriculture), MOE (Ministry of Education), Pharmecore, ASEL, MOEM (Ministry of Energy and Mines), OTHER (type name). The current template's CONTROLS sheet still says "Pharmaco" and must be corrected to "Pharmecore" in build_templates.py, then the workbooks regenerated.
- The .xltm template has not been saved. The VBA has not been compiled in Excel. The SharePoint deployment is documented but untested.
- The lock folder needs create-folder rights; clock skew over 180 s can break a healthy lock; sheet protection has no password.
- Not in scope: no email, no tax/duty/FX, no sending of documents.

## 13. Glossary
- QT: Quotation. PI: Proforma Invoice. PO: Supplier Purchase Order. PL: Packing List. DN: Delivery Note. CI: Commercial Invoice.
- CBM: cubic metres. Gross kg: total shipping weight. VOID: a reserved number that was not used.
- Job file: a user's working copy (.xlsm). Issued file: values-only .xlsx sent out.
