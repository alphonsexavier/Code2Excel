# PowerApps YAML to Excel Exporter

This Python script processes **PowerApps YAML files** and converts their screen, control, and property information into a structured **Excel file**.  
It extracts **screens, controls, properties, and pseudocode instructions** from YAML files and organizes them into sheets (one per YAML file).

---

## Features

- Reads all `.yaml` / `.yml` files in a specified folder.
- Handles PowerApps YAML files containing:
  - Screens
  - Controls (and nested children)
  - Properties
- Generates a **pseudocode description** for each property.
- Exports results into an **Excel file**, with each YAML file represented as a separate sheet.
- Supports nested control hierarchies using recursive parsing.

---

## Output Structure

The generated Excel file contains columns:

| SCREEN NAME   | CONTROL NAME         | PROPERTY   | PSEUDOCODE                      |
|---------------|----------------------|------------|---------------------------------|
| HomeScreen    | (Screen Property)    | Fill       | Set the Fill as RGBA(0,0,0,1)   |
| HomeScreen    | Label1               | Text       | Set the Text as "Hello World"   |
| HomeScreen    | Gallery1/Label2      | Color      | Set the Color as Red            |

- **SCREEN NAME** → The PowerApps screen where the control belongs.
- **CONTROL NAME** → The control (with parent hierarchy if nested).
- **PROPERTY** → The property name (e.g., Text, Fill, Color).
- **PSEUDOCODE** → A natural-language representation of the property setting.

---

## Installation

Make sure you have **Python 3.8+** installed.  
Then, install the required dependencies:

```bash
pip install pandas pyyaml openpyxl
