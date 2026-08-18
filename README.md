# XMLScraped

XMLScraped is a Python desktop application (using Tkinter) designed to extract and scrape financial invoice data from Mexican CFDI (Comprobante Fiscal Digital por Internet) XML files stored inside ZIP archives. It validates invoice status directly with the SAT SOAP service and outputs a fully structured Excel (`.xlsx`) sheet.

---

## Features Implemented
We have introduced several key enhancements to the initial version:
- **Modular Architecture**: Refactored the single monolithic file `XMLScraped.py` into separate cleanly structured modules (`main.py`, `parser.py`, `sat_client.py`, `exporter.py`) for easier troubleshooting and future scaling.
- **Persistent GUI Loop**: The application window stays open after processing a ZIP file, letting you run multiple tasks without having to relaunch it each time.
- **Numeric Excel Cells**: Amounts (Subtotal, Total, IVA, Retenciones) are correctly formatted as float values in the Excel sheet instead of raw text strings, allowing Excel formulas (like SUM) to work instantly.
- **Negative Inversion for Expenditures (Egresos)**: CFDI files of type `E` (Egreso) automatically invert their financial outputs (Subtotal, Total, IVA, Retenciones) to negative values (`value * -1.0`) to balance accounting figures.
- **"P" Type (Payments/Pagos) parsing**: Fully supports CFDI payment files (v2.0 payment complements). It maps specific attributes like `MontoTotalPagos`, `TotalTrasladosBaseIVA16`, `TotalTrasladosImpuestoIVA16`, and aggregates associated document UUIDs (`IdDocumento`) from child elements into the concepts column.

---

## Comparison: First vs. Latest Version

| Feature | First (Monolithic) Version | Latest (Modular) Version |
| :--- | :--- | :--- |
| **Structure** | Single monolithic script (`XMLScraped.py`) | Modular codebase (`main.py`, `parser.py`, etc.) |
| **Application Lifecycle** | Automatically closes/terminates when processing finishes | Remains open (`root.deiconify()`) for multiple runs |
| **Excel Cell Formats** | Financial data written as string text | Financial values formatted as true float numbers |
| **CFDI "Egreso" (Type E)** | Totals represented as positive numbers | Automatically inverted to negative numbers |
| **CFDI "Pagos" (Type P)** | Standard node parser (frequently missing fields) | Targeted parsing of `Pago20` and associated IDs |

---

## Setup and Usage Steps

### Prerequisites
- Python 3.10+
- Installed virtual environment dependencies (`openpyxl`, `zeep`, `requests`, `lxml`)

### Run Application
1. Open terminal and run:
   ```pwsh
   python main.py
   ```
2. Choose **Ingresos** or **Gastos**.
3. Select any additional parsing filters (e.g., *Ambas*, *Nómina*, *Deducciones*).
4. Click **Procesar Archivos** and select your ZIP file containing CFDI XMLs.
5. Save the generated `.xlsx` file.

### Build Executable
To package the app into a single windowed executable using PyInstaller:
```pwsh
.\venv\Scripts\activate
python -m PyInstaller --noconfirm --onefile --windowed --name XMLScraped main.py
```
Find the output executable `XMLScraped.exe` in the `dist/` directory.
