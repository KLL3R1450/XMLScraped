# Codebase Documentation & Context

This document provides a technical overview of the XMLScraped repository, outlining its architecture, modules, mapping logic, and build/compilation processes.

---

## 1. Project Overview
**XMLScraped** is a Python desktop application (using Tkinter) designed to extract financial information from Mexican CFDI (Comprobante Fiscal Digital por Internet) XML files stored inside ZIP archives. It queries the SAT SOAP web service to check the cancellation status of each invoice and saves the compiled results into an Excel (`.xlsx`) file.

---

## 2. Directory Structure
```
XMLScraped/
├── main.py            # GUI setup and workflow runner (Tkinter)
├── parser.py          # XML parser & schema extractor (Ingresos, Gastos, Nómina, Deducciones, Pagos)
├── sat_client.py      # SOAP client for SAT cancelation queries
├── exporter.py        # Excel sheet writer (openpyxl formatting)
├── XMLScraped.py      # [Legacy] Monolithic backup script
└── requirements.txt   # Project dependencies
```

---

## 3. Module Breakdown

### A. Entry Point: [`main.py`](file:///C:/Users/Usuario/Documents/XMLScraped/main.py)
* Manages the Tkinter interface.
* Asks user for CFDI options ("Ingresos" or "Gastos") and additional options ("Retención IVA", "Retención ISR", "Ambas", "Ninguna", "Deducciones", "Nómina").
* Runs processing in a separate background thread (`threading.Thread`) so the UI does not freeze.
* **Persistent Window**: Instead of quitting the process when finished, it deiconifies the root window (`root.deiconify()`), allowing the user to queue multiple ZIP files sequentially.

### B. Extractor / Schema Parser: [`parser.py`](file:///C:/Users/Usuario/Documents/XMLScraped/parser.py)
This module reads XML nodes and attributes, converting them into structured lists or dictionaries.

#### Key Features & CFDI Mappings:
1. **CFDI "Egreso" (Type E) Inversion**:
   * If a document's `TipoDeComprobante` is `"E"` (representing returns or expenditures), all calculated numeric fields (Subtotal, Total, IVA, Retenciones) are multiplied by `-1.0` to reflect negative transactions in accounting.
2. **CFDI "Pagos" (Type P) Support**:
   * Uses namespace `http://www.sat.gob.mx/Pagos20` (prefix `pago20`).
   * Extracts fields from the `<pago20:Totales>` and `<pago20:Pago>` nodes:
     * **Total**: Mapped to `MontoTotalPagos`.
     * **Subtotal**: Mapped to `TotalTrasladosBaseIVA16`.
     * **IVA**: Mapped to `TotalTrasladosImpuestoIVA16`.
     * **Forma Pago**: Mapped to `FormaDePagoP`.
     * **Moneda**: Mapped to `MonedaP` (defaults to MXN/Moneda if not specified).
     * **Retención IVA**: Mapped to `TotalRetencionesIVA`.
     * **Retención ISR**: Mapped to `TotalRetencionesISR`.
     * **Conceptos**: Instead of showing standard descriptions, it extracts all `IdDocumento` attributes from child `<pago20:DoctoRelacionado>` nodes, concatenated with commas.

### C. SAT Client: [`sat_client.py`](file:///C:/Users/Usuario/Documents/XMLScraped/sat_client.py)
* Uses `zeep` client to query the SAT validation endpoint:
  `https://consultaqr.facturaelectronica.sat.gob.mx/ConsultaCFDIService.svc?wsdl`
* Returns the status (e.g. "Vigente", "Cancelado") of the invoice based on its UUID, issuer RFC, receiver RFC, and formatted Total value.

### D. Exporter: [`exporter.py`](file:///C:/Users/Usuario/Documents/XMLScraped/exporter.py)
* Appends headers and rows to an `openpyxl.Workbook`.
* Automatically parses values to actual floats if they are numerical, ensuring Excel recognizes totals/subtotals as numbers rather than strings.

---

## 4. How to Compile / Run
To run the python GUI:
```pwsh
python main.py
```

To compile a single-file executable package using the virtual environment:
```pwsh
.\venv\Scripts\activate
python -m PyInstaller --noconfirm --onefile --windowed --name XMLScraped main.py
```
Output executable is generated in the `dist/` directory as `XMLScraped.exe`.
