# Grand Group FZC Procurement Document System — Complete Guide & Google NotebookLM Video Tutorial Storyboard

> **Status:** Updated with all latest capabilities — Custom Markup % (`0–500%`), Custom VAT % (`0–100%`), Flexible 5-Position Extra Order Columns (add & move anytime), Offline Supplier Quote Importer (`.xlsx` / `.pdf` / Pasted Text), Editable Per-Job Remarks & Dual Signatories, 3 Executive Excel Themes (`v1`, `v2`, `v3`), Master Registry Sequential Locking, and Dual-Platform Offline Builds (Windows `.html` + iPad/iOS `.zip` PWA).

---

## 1. Purpose & Business Utility

The **Grand Group FZC Procurement Document System** is a **100% offline-first commercial procurement workstation** built for international trading, industrial supply, aviation spares, and chemical/pharmaceutical procurement teams.

### Problems It Eliminates
1. **Duplicate Data Entry:** Typing the same 20 line items across Quotation, Proforma Invoice, Supplier PO, Packing List, Delivery Note, and Commercial Invoice wastes hours and creates quantity/specification mismatches.
2. **Accidental Margin Leaks:** Sending internal costing spreadsheets to customers exposes supplier unit costs and profit margins.
3. **Rigid Templates:** Real procurement jobs need flexible markups (not just 4 fixed percentages), flexible VAT rates (5% UAE, 0% export, or custom regional rates), custom order columns (`Part No.`, `CAS No.`, `HS Code`, `Batch / Expiry`) positioned wherever the client expects them, and job-specific remarks/signatories.
4. **Inconsistent Document Numbers:** Manual numbering leads to duplicate invoice numbers or missing audit trails.
5. **Connectivity & Device Barriers:** Procurement officers work across Windows office PCs and iPads/iPhones in warehouses or transit, often without reliable internet access.

---

## 2. Core System Architecture & Latest Feature Set

| Module / Capability | What It Does | Key Controls & Rules |
| :--- | :--- | :--- |
| **Proforma Invoice (`PI`)** | Single-entry master sheet (Excel Rows 17–36). Stores both End-User Requested Description and Supplier Offered Description. | End-User dropdown (`Pharmecore` included + `OTHER`), Contact/ATTN, Address, Currency, Incoterm, **Custom VAT % (`PI!L5`)**, Editable Remarks & Left/Right Signature Labels, **Flexible Extra Columns (5 positions)**. |
| **Price Calc (`Internal`)** | Confidential internal costing & logistics margin engine. Never included in customer exports. | **Active Markup % (`Cell E3`)** with quick presets (`10%`, `15%`, `18%`, `20%`) **+ Custom % box (`0–500%`, decimals supported)**, live Active Custom Markup summary banner, Box `L/W/H (cm)` → CBM, `Kg/Box` → Gross KG, `Rate/CBM`, `Rate/KG`, `Doc Charge %`. |
| **Supplier Quote Importer** | Reads **any-format supplier `.xlsx`**, text-layer **`.pdf`**, or **pasted text** 100% offline. | Auto-captures supplier name, quote ref, currency, incoterm, payment terms, lead time, Part/CAS codes, descriptions, UoM, qty, and unit costs. Choose custom column placement and load into active job or new job. |
| **Quotation (`QT`)** | Formal customer commercial offer auto-populated from `PI` & `PRICE CALC`. | Zero duplicate entry; mirrors custom columns, custom markup, VAT, and freight. |
| **Supplier PO (`PO`)** | Supplier-facing Purchase Order with strict cost-only firewall. | Shows strictly **Supplier Unit Cost** & **Total Order Cost** — zero customer markup exposed. |
| **Packing List (`PL`)** | Shipment packaging & weight schedule. | Ship Date, Port of Loading, Destination, Carton No., independently verified Net & Gross Weights (KG), automatic Shipment Totals. |
| **Delivery Note (`DN`)** | Handover & warehouse receipt document. | Delivery Date, Transport/BL/Truck Ref., Received By, fully editable job Remarks and Left/Right Signature blocks (clear any field to omit). |
| **Commercial Invoice (`CI`)** | Final commercial billing document with bank wire details. | One-click **Issue** validates required fields, claims the next sequential number from Master Registry, and freezes financial totals. |
| **Master Registry** | Sequential audit register (`PREFIX-GG-SEQ-YEAR`), seeded at `PI-GG-042-2026` (next `#043`). | Strict sequential claiming, **VOID** retirement with mandatory audit reason (retired numbers are never reused), offline `Master_Registry.xlsx` export & sync. |
| **Excel Studio & Import** | Styling, branding, live-formula workbooks, backups, and offline app downloads. | 3 Executive Excel Themes (**v1 Navy**, **v2 Emerald**, **v3 Charcoal & Amber**), editable Company Profile/Logo/End-User dropdown, **Backup/Restore JSON** (for Laptop ↔ iPad transfer), **Windows `.html`**, and **iPad/iOS `.zip` PWA**. |

---

## 3. Confirmed Financial & Commercial Formulas

```text
Unit Selling Price (PRICE CALC Col R → PI Unit Price):
  = ROUND(Supplier_Unit_Cost × (1 + Active_Markup_E3), 2)

Active Markup % (Cell E3):
  Quick presets (10%, 15%, 18%, 20%) OR any custom figure from 0% to 500% (e.g., 12.5%)

Line Volume (CBM):
  = (Boxes × L_cm × W_cm × H_cm) / 1,000,000

Base Freight:
  = MAX(Total_CBM × Rate_per_CBM, Total_Gross_KG × Rate_per_KG)

Customer Freight (PRICE CALC R41 → PI H40):
  = ROUND((Base_Freight × (1 + Doc_Fee_%)) × (1 + Active_Markup_E3), 2)

VAT (PI!L5 — default 5%, 0% zero-rated, or any custom % 0–100%):
  = Goods_Subtotal × VAT_Rate    [Strictly applied to Goods ONLY — never on Freight]

Grand Total:
  = Goods_Subtotal + VAT + Customer_Freight
```

---

## 4. End-to-End Start-to-Finish Demo Walkthrough (7 Steps)

1. **Create a Job or Upload a Supplier Quote:** Start fresh on **Proforma Invoice**, or tap **Open .xlsx / .pdf** to import any supplier Excel or PDF quote directly into the review modal.
2. **Configure Controls, Custom VAT & Extra Columns:** Select End User (or `OTHER`), fill Contact & Address, set Currency/Incoterm/Custom VAT %, and add extra columns (e.g., `Part No.` or `CAS No.`) using the **Place:** selector (`After No.`, `After Requested Desc`, `After Offered Desc`, `After Qty`, `End of Table`). Move columns anytime via their inline chip dropdown.
3. **Calculate Margins & Freight (`Price Calc`):** Enter supplier costs and carton dimensions. Compare 10% / 15% / 18% / 20% scenarios or type **any custom percentage** (e.g., `12.5%`) and inspect live Gross Profit and Customer Freight.
4. **Review & Issue Customer Documents (`QT` / `PI`):** Toggle **A4 Print & Excel Preview**, tailor per-job Remarks & Signatures (or leave blank), click **Issue** to lock the sequential document number, and export **Values-Only `.xlsx`** or **Print / PDF**.
5. **Issue Cost-Only Supplier PO (`PO`):** Open **Supplier PO** — raw supplier unit costs populate automatically with zero margin leakage.
6. **Complete Logistics (`PL` & `DN`):** Enter carton numbers, verified net/gross weights, ship/delivery dates, and transport references on **Packing List** and **Delivery Note**.
7. **Issue Commercial Invoice (`CI`), Audit Registry & Backup:** Click **Issue CI** to claim `#043` in **Master Registry**, and download a **Backup JSON** in **Excel Studio** whenever transferring workspaces between Windows PC and iPad.

---

## 5. Complete 13-Screen Visual Workflow Storyboard (For Video Tutorial)

### Scene 01 — Master Job Entry, End-User Selection & Custom VAT %
![Proforma Invoice Entry & Controls](screenshots/01_pi_entry_controls.png)
- **Purpose:** Single-entry master controls (`PI` Cell `C8` & Block `K:L`). Set End User, Contact/ATTN, Address, Currency, Incoterm, Destination, preset `5%`/`0%` or Custom VAT %, and editable job Remarks & Signatories.
- **Voiceover Script:** *"Welcome to the Grand Group FZC Procurement Document System. Instead of typing the same order into six spreadsheets, every job starts once here on the Proforma Invoice entry sheet. Choose the client, currency, incoterm, and exact VAT percentage—preset or custom—and set job-specific remarks and signature names."*

### Scene 02 — Flexible Extra Order Columns (Choose & Move Position Anytime)
![Single-Entry Item Table with Flexible Custom Columns](screenshots/02_pi_items_custom_col.png)
- **Purpose:** Add up to 4 custom columns (`Part No.`, `CAS No.`, `Brand / Mfr`, `HS Code`, `Batch / Expiry`, or custom) into any of 5 table positions — and move existing columns anytime with zero formula impact.
- **Voiceover Script:** *"Different orders need different specifications. With one tap, add Part Number for spare parts or CAS Number for chemicals, and choose exactly where the column sits in the table. Need to move it later? Switch its position dropdown and all six documents and Excel exports update automatically."*

### Scene 03 — Internal Costing (`PRICE CALC`) with Custom Markup % & Freight Engine
![Internal Price Calc with 12.5% Custom Markup](screenshots/03_price_calc_custom_markup.png)
- **Purpose:** Confidential internal costing tab (never exported to customers). Combines quick presets (`10%`, `15%`, `18%`, `20%`) with an inline **Custom %** box (`12.5%` shown), live Active Custom Markup totals banner, and `MAX(CBM, KG)` freight + doc fee + markup engine.
- **Voiceover Script:** *"Switch to the confidential Price Calc tab, highlighted in red so it is never accidentally shared with customers. Enter supplier costs and box dimensions. Compare 10, 15, 18, and 20 percent scenarios side by side—or type any custom percentage like 12.5 percent. Column R feeds selling prices straight back to your Proforma Invoice and freight."*

### Scene 04 — Offline Supplier Quote Importer (`.xlsx` / `.pdf` / Pasted Text)
![Supplier Quote Importer Modal](screenshots/12_supplier_quote_importer.png)
- **Purpose:** Captures descriptions, Part/CAS numbers, UoM, quantities, unit costs, currency, and payment/delivery terms from any supplier format 100% offline.
- **Voiceover Script:** *"Suppliers rarely use our exact spreadsheet format. Drop in any supplier Excel or PDF quote—or paste copied lines—and the offline importer captures descriptions, part numbers, quantities, unit costs, and payment terms. Review the rows, pick where the Part Number column goes, and send everything into Price Calc in one click."*

### Scene 05 — Proforma Invoice (`PI`) Executive A4 Preview & Values-Only `.xlsx`
![Proforma Invoice A4 Executive Preview](screenshots/04_pi_a4_preview.png)
- **Purpose:** Client-ready A4 printout and styled Values-Only Excel export showing Grand Group logo, custom columns, Goods Subtotal, VAT on goods only, separate Freight line, Grand Total, bank wire details, and editable signature blocks.
- **Voiceover Script:** *"Switching to A4 Print Preview reveals an executive, client-ready Proforma Invoice. Notice our Part Number column neatly aligned, Goods Subtotal, VAT applied strictly to goods, freight shown as a clean separate line after VAT, bank details, and customizable signature lines."*

### Scene 06 — Quotation (`QT`) Auto-Linked Commercial Offer
![Quotation A4 Preview](screenshots/05_quotation_preview.png)
- **Purpose:** Automatically populated formal customer quotation sharing the exact same single-entry data and custom column layout.
- **Voiceover Script:** *"Moving down the trade workflow, the Quotation tab is already completely populated—same part numbers, same selling prices, same professional layout, with zero duplicate typing."*

### Scene 07 — Supplier Purchase Order (`PO`) Cost-Only Firewall
![Supplier Purchase Order A4 Preview](screenshots/06_supplier_po_preview.png)
- **Purpose:** Displays strictly raw Supplier Unit Costs and Total Order Cost — guaranteeing customer margins remain confidential.
- **Voiceover Script:** *"When the customer approves the order, open Supplier PO. The system automatically swaps selling prices for raw supplier unit costs, ensuring your margin stays strictly confidential."*

### Scene 08 — Packing List (`PL`) Carton & Independent Weight Verification
![Packing List A4 Preview](screenshots/07_packing_list_preview.png)
- **Purpose:** Logistics schedule tracking Carton Numbers, Net Weight (KG), and Gross Weight (KG) per row with automated Shipment Totals.
- **Voiceover Script:** *"At shipment stage, switch to Packing List. Enter carton numbers and independently verified net and gross weights. All item descriptions, part numbers, and quantities carry over automatically."*

### Scene 09 — Delivery Note (`DN`) with Editable Remarks & Dual Signatories
![Delivery Note A4 Preview](screenshots/08_delivery_note_preview.png)
- **Purpose:** Handover and warehouse receipt document with Transport/BL Ref., customizable remarks, and editable Left/Right signatory lines.
- **Voiceover Script:** *"For final handover, Delivery Note pairs shipment weights and transport references with fully editable remarks and dual signature blocks. Need a specific signer's name—or want to leave it blank? Edit it per job in seconds."*

### Scene 10 — Issuing & Freezing Commercial Invoice (`CI`) — Claiming `#043`
![Commercial Invoice Issued and Frozen](screenshots/09_commercial_invoice_issued.png)
- **Purpose:** Pre-issue validation checks mandatory fields before consuming any sequence number, then claims `CI-GG-043-2026` from the Master Registry and locks the document.
- **Voiceover Script:** *"When you are ready to issue the Commercial Invoice, click Issue. Built-in validation checks every required field first so no registry number is ever wasted. Upon approval, the system claims the next sequential number—CI-GG-043-2026—and freezes the document for audit integrity."*

### Scene 11 — Master Registry Sequential Audit Log & Void Safety
![Master Registry Audit Log](screenshots/10_master_registry.png)
- **Purpose:** Single source of truth for document numbering (`PREFIX-GG-SEQ-YEAR`), seeded at `PI-GG-042-2026` with permanent `VOID` retirement and offline `.xlsx` synchronization.
- **Voiceover Script:** *"Every issued document is permanently logged in the Master Registry, continuing cleanly from sequence 042 to 043. If an invoice ever needs cancellation, marking it Void retires the number permanently with a recorded audit reason."*

### Scene 12 — Excel Studio: 3 Style Themes, Branding, JSON Backup & Offline Builds
![Excel Studio and Offline Downloads](screenshots/11_excel_studio_themes_offline.png)
- **Purpose:** Switch between **v1 Navy**, **v2 Emerald**, and **v3 Charcoal & Amber**; download live-formula `GG_Job_Template.xlsx`, current job `.xlsx`, or `Master_Registry.xlsx`; edit company logo/bank info/End-Users with zero code; transfer data via **Backup/Restore JSON**; and download **Windows `.html`** or **iPad `.zip`**.
- **Voiceover Script:** *"Finally, Excel Studio puts full control in your hands. Switch between three executive Excel color themes, replace your company logo or End-User list with zero coding, back up everything to a single JSON file, or download the standalone offline Windows and iPad versions."*

### Scene 13 — Dedicated iPad / iOS Offline Touch Interface
![iPad Touch Card Mode](screenshots/13_ipad_touch_mode.png)
- **Purpose:** Large touch-friendly item cards, quick document switcher, and bottom navigation dock running 100% offline from the iPad/iPhone Home Screen via Service Worker cache.
- **Voiceover Script:** *"On the road or in the warehouse, iPad and iOS mode transforms wide tables into large, touch-friendly item cards with a bottom navigation bar—running full-screen and completely offline straight from your Home Screen."*

---

## 6. How to Create Your Video Tutorial in Google NotebookLM

### Best Way to Share This App with Google NotebookLM
Because AI Studio preview links are private (behind Google login) and interactive web apps require clicks to reveal every tab, **the highest-quality way to give Google NotebookLM full access to every screen and feature** is to download the **1-file Visual Guide PDF (with all 13 real screenshots embedded)** right from **Excel Studio & Import → Video Tutorial & NotebookLM Source Kit**:

1. Open **Excel Studio & Import** inside the app and click:
   - **`Visual Guide + 13 Screenshots (.pdf)`** — *Best single upload for NotebookLM (contains all text, formulas, storyboard, voiceover script, AND all 13 full-color screenshots).*
   - **`Demo Video Walkthrough (.webm)`** — *Real start-to-finish recorded screen demo you can use directly or drop into Google Drive / any video editor.*
   - **`Text Guide (.md)`** — *Optional plain-text Markdown copy.*
2. Go to **[notebooklm.google.com](https://notebooklm.google.com)** → **Create New Notebook**.
3. Click **Add Sources → Upload File** and upload `Grand_Group_Procurement_NotebookLM_Tutorial_Guide.pdf`.
4. *(Optional)* You can also add your public Netlify URL (`https://dynamic-heliotrope-28ad25.netlify.app`) as a secondary website source, though the PDF alone gives NotebookLM 100% of every screen, button, formula, and visual!
5. In NotebookLM's right-hand **Studio** panel, click **Video Overview** (or **Audio Overview**) → click the **Customize (pencil)** icon, and paste this exact prompt:

```text
Create a clear, engaging step-by-step product tutorial video explaining the Grand Group FZC Procurement Document System from start to finish.
Cover:
1. Why the system exists: single-entry workflow across 6 linked trade documents, commercial confidentiality firewall (keeping supplier costs hidden from customers), sequential Master Registry numbering, and 100% offline operation on both Windows PC and iPad/iOS.
2. Walk chronologically through a real procurement job using all 13 storyboard scenes:
   - Proforma Invoice single-entry setup, Custom VAT %, editable remarks/signatures, and Flexible Extra Columns (adding & moving Part No. / CAS No. across 5 positions)
   - Offline Supplier Quote Importer (.xlsx, .pdf, and pasted text)
   - Confidential Price Calc engine with Custom Markup % (e.g., 12.5%) and MAX(CBM, KG) Freight + Doc Fee + Markup calculation
   - Auto-linked Quotation, Cost-Only Supplier PO, Packing List, Delivery Note, and Commercial Invoice
   - Pre-issue validation, claiming sequence #043 in the Master Registry, and exporting Values-Only .xlsx
   - Switching between the 3 Excel themes (Navy, Emerald, Charcoal), JSON Backup/Restore, and offline Windows/iPad deployment.
Keep the tone practical, executive, and easy for procurement staff to follow on day one.
```
