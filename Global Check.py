import pandas as pd
import xlwings as xw
import os
from pathlib import Path
from tkinter import Tk, filedialog
from tkinter import messagebox
import traceback

try:
    # === Prompt user to select Input file ===
    print("📁 Launching file picker for Input.xlsx...")

    Tk().withdraw()
    input_file = filedialog.askopenfilename(
        title="Select Input.xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )

    if not input_file:
        raise FileNotFoundError("❌ No input file selected.")
    else:
        print(f"✅ Input file selected: {input_file}")

    global_file = "Global.xlsm"
    target_sheet = "Data"

    if not Path(global_file).is_file():
        raise FileNotFoundError(f"File not found: {global_file}")
    else:
        print(f"🔄 Found global file: {global_file}")

    # === Load all DataFrames ===
    print("📥 Loading: Story Definitions")
    df_story = pd.read_excel(input_file, sheet_name="Story Definitions", usecols="A:E", skiprows=3, header=None)
    df_story.columns = ["Name", "Height", "Master Story", "Similar To", "Splice Story"]
    df_story["Height"] = pd.to_numeric(df_story["Height"], errors="coerce")
    df_story = df_story[::-1].reset_index(drop=True)
    df_story["Elevation"] = df_story["Height"].cumsum()
    df_story = df_story[::-1].reset_index(drop=True)

    print("📥 Loading: Modal Participating Mass Ratios")
    df_modal = pd.read_excel(input_file, sheet_name="Modal Participating Mass Ratios", skiprows=3, header=None)
    df_modal.columns = ["Case", "Mode", "Period", "UX", "UY", "UZ", "SumUX", "SumUY", "SumUZ", "RX", "RY", "RZ", "SumRX", "SumRY", "SumRZ"]

    print("📥 Loading: Story Drifts")
    df_drift = pd.read_excel(input_file, sheet_name="Story Drifts", skiprows=3, header=None)
    df_drift.columns = ["Story", "Output Case", "Case Type", "Step Type", "Direction", "Drift", "Label", "X", "Y", "Z"]

    print("📥 Loading: Diaphragm Max Over Avg Drifts")
    df_diaphragm = pd.read_excel(input_file, sheet_name="Diaphragm Max Over Avg Drifts", skiprows=3, header=None)
    df_diaphragm.columns = ["Story", "Output Case", "Case Type", "Step Type", "Item", "Max Drift", "Avg Drift", "Ratio", "Label", "Max Loc X", "Max Loc Y", "Max Loc Z"]

    print("📥 Loading: Story Forces")
    df_force = pd.read_excel(input_file, sheet_name="Story Forces", skiprows=3, header=None)
    df_force.columns = ["Story", "Output Case", "Case Type", "Step Type", "Location", "P", "VX", "VY", "T", "MX", "MY"]
    df_force = df_force.reset_index(drop=True)

    print("📥 Loading: Joint Displacements")
    df_joint = pd.read_excel(input_file, sheet_name="Joint Displacements", skiprows=3, header=None)
    df_joint.columns = ["Story", "Label", "Unique Name", "Output Case", "Case Type", "Step Type", "UX", "UY", "UZ", "RX", "RY", "RZ"]

    print("📥 Loading: Diaphragm CM Displacements")
    df_cm = pd.read_excel(input_file, sheet_name="Diaphragm CM Displacements", skiprows=3, header=None)
    df_cm.columns = ["Story", "Diaphragm", "Output Case", "Case Type", "Step Type", "UX", "UY", "RZ", "Point", "X", "Y", "Z"]

    print("📥 Loading: Story Stiffness")    
    df_stiffness = pd.read_excel(input_file, sheet_name="Story Stiffness", skiprows=3, header=None)
    df_stiffness.columns = ["Story", "Output Case", "Case Type", "Step Type", "Shear X", "Drift X", "Stiff X", "Shear Y", "Drift Y", "Stiff Y"]

    # === Excel helper functions ===
    def col_letter(idx):
        return xw.utils.col_name(idx)

    def write_block(ws, cell, title, headers, units, df, name):
        print("✍️  Writing to Excel...")
        start_col = ws.range(cell).column
        start_row = ws.range(cell).row
        data_start = start_row + 3

        ws.range(cell).value = title
        ws.range((start_row + 1, start_col)).value = headers
        ws.range((start_row + 2, start_col)).value = units
        ws.range((start_row + 1, start_col), (start_row + 2, start_col + len(headers) - 1)).color = (198, 239, 255)
        ws.range((data_start, start_col)).value = df.values.tolist()
        end_row = data_start + len(df) - 1
        end_col = start_col + len(headers) - 1
        ws.range((start_row + 1, start_col), (end_row, end_col)).api.Borders.Weight = 2

        ref = f"{target_sheet}!${col_letter(start_col)}${start_row + 1}:${col_letter(end_col)}${end_row}"
        if name in [n.name for n in wb.names]:
            wb.names[name].delete()
        wb.names.add(name, refers_to=f"={ref}")
        return data_start

    # === Open Excel and clear contents ===
    app = xw.App(visible=False)
    wb = xw.Book(global_file)
    ws = wb.sheets[target_sheet]

    ws.range("A:F").clear()
    ws.range("H:V").clear()
    ws.range("X:AG").clear()
    ws.range("AI:AT").clear()
    ws.range("AV:BF").clear()
    ws.range("BH:BS").clear()
    ws.range("BU:CF").clear()
    ws.range("CH:CR").clear()

    # === Write data blocks ===
    write_block(ws, "A1", "Story Definitions", df_story.columns.tolist(), ["", "in", "", "", "", "in"], df_story, "StoryDefinitions")
    write_block(ws, "H1", "Modal Participating Mass Ratios", df_modal.columns.tolist(), ["", "", "sec"], df_modal, "ModalMassRatios")
    write_block(ws, "X1", "Story Drift", df_drift.columns.tolist(), ["", "", "", "", "", "", "", "in", "in", "in"], df_drift, "StoryDrifts")
    write_block(ws, "AI1", "Diaphragm Max Over Avg Drifts", df_diaphragm.columns.tolist(), ["", "", "", "", "", "", "", "", "", "in", "in", "in"], df_diaphragm, "DiaphragmMaxOverAvgDrifts")
    force_data_start = write_block(ws, "AV1", "Story Forces", df_force.columns.tolist(), ["", "", "", "", "", "kip", "kip", "kip", "kip-in", "kip-in", "kip-in"], df_force, "StoryForces")
    write_block(ws, "BH1", "Joint Displacements", df_joint.columns.tolist(), ["", "", "", "", "", "", "in", "in", "in", "rad", "rad", "rad"], df_joint, "JointDisplacements")
    write_block(ws, "BU1", "Diaphragm Center Of Mass Displacements", df_cm.columns.tolist(), ["", "", "", "", "", "in", "in", "rad", "", "in", "in", "in"], df_cm, "DiaphragmCMDisplacements")
    write_block(ws, "CH1", "Story Stiffness", df_stiffness.columns.tolist(), ["", "", "", "", "kip", "in", "kip/in", "kip", "in", "kip/in"], df_stiffness, "StoryStiffness")

    # === Final formatting ===
    ws.range("A:CR").autofit()
    ws.range("A:CR").api.HorizontalAlignment = -4108
    ws.range("A:CR").api.VerticalAlignment = -4108
    for cell in ["A1", "H1", "X1", "AI1", "AV1", "BH1", "BU1", "CH1"]:
        ws.range(cell).api.Font.Bold = True

    wb.save()
    print("✅ Excel saved successfully.")
    wb.close()

except Exception as e:
    error = traceback.format_exc()
    messagebox.showerror("❌ Unexpected Error", error)

finally:
    try:
        app.quit()
    except:
        pass

print("🎉 All tasks completed. Closing Excel.")


