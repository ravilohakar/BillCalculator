# New Viransh Wine Shop - Bill Calculator

A desktop billing application for managing sales, generating bills, and tracking payments.

## Features

- Search and browse items from Excel data
- Add items to bill with box and bottle quantities
- Editable discount per item
- Auto-calculated totals and balance
- Save bills to Excel with unique bill numbers
- Fetch and view previously saved bills
- Print bill receipts
- Input validation for quantities and discounts

## Requirements

- Python 3.8+
- Windows OS

## Installation

### Option 1: Quick Install (Python required)

1. Install Python 3.8+ from https://python.org (check "Add Python to PATH")
2. Copy the `BillCalculator` folder to your desktop
3. Double-click `install.bat`

### Option 2: Standalone EXE (no Python needed)

1. On a machine with Python, run `build_exe.bat`
2. Copy the `dist` folder to the target PC
3. Double-click `BillCalculator.exe`

## Manual Run

```
pip install -r requirements.txt
py bill_calculator.py
```

## Files

| File | Description |
|------|-------------|
| `bill_calculator.py` | Main application |
| `items_data.xlsx` | Item master data (name, size, price, discount) |
| `Bills_generated.xlsx` | Saved bills (auto-created on first save) |
| `create_sample_data.py` | Script to generate sample item data |
| `install.bat` | Installation script |
| `build_exe.bat` | Builds standalone .exe |
| `requirements.txt` | Python dependencies |

## Usage

1. Launch the app
2. Search/browse items and enter quantities
3. Click "Add Items to Bill" to add to the bill
4. Enter "Amount Received" to see balance
5. Click "Save Bill" to store the bill
6. Click "Fetch Bill" to retrieve a saved bill by number
7. Click "Print Bill" to print a receipt
8. Click "Clear All" to reset everything
