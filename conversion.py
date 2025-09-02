import pandas as pd
import yaml
import os

class NoTagLoader(yaml.SafeLoader):
    pass

def unknown_constructor(loader, tag_suffix, node):
    return loader.construct_scalar(node)

NoTagLoader.add_multi_constructor('', unknown_constructor)

def extract_controls(screen, ctrl_name, ctrl_data, rows, parent=""):
    if "Properties" in ctrl_data:
        for prop, val in ctrl_data["Properties"].items():
            full_name = f"{parent}/{ctrl_name}" if parent else ctrl_name
            rows.append([screen, full_name, prop, f"Set the {prop} as {val}"])

    if "Children" in ctrl_data:
        for child in ctrl_data["Children"]:
            for sub_ctrl_name, sub_ctrl_data in child.items():
                full_name = f"{parent}/{ctrl_name}" if parent else ctrl_name
                extract_controls(screen, sub_ctrl_name, sub_ctrl_data, rows, full_name)


def powerapps_to_dataframe(input_file):
    with open(input_file, "r") as f:
        data = yaml.load(f, Loader=NoTagLoader)

    rows = []

    for screen, screen_data in data.get("Screens", {}).items():
        # Screen-level properties
        if "Properties" in screen_data:
            for prop, val in screen_data["Properties"].items():
                rows.append([screen, "(Screen Property)", prop, f"Set the {prop} as {val}"])

        # Recursively extract children
        if "Children" in screen_data:
            for child in screen_data["Children"]:
                for ctrl_name, ctrl_data in child.items():
                    extract_controls(screen, ctrl_name, ctrl_data, rows)

    df = pd.DataFrame(rows, columns=["SCREEN NAME", "CONTROL NAME", "PROPERTY", "PSEUDOCODE"])
    return df


def process_folder_to_excel(folder_path, output_file):
    yaml_files = [f for f in os.listdir(folder_path) if f.endswith((".yaml", ".yml"))]
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        for yaml_file in yaml_files:
            file_path = os.path.join(folder_path, yaml_file)
            df = powerapps_to_dataframe(file_path)
            # Use file name without extension as sheet name (max 31 chars for Excel)
            sheet_name = os.path.splitext(yaml_file)[0][:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"Processed {yaml_file} -> Sheet: {sheet_name}")

    print(f"All files exported to {output_file}")


if __name__ == "__main__":
    folder_path = "Files"  # Folder containing YAML files
    output_file = "PowerApps_Export.xlsx"
    process_folder_to_excel(folder_path, output_file)
