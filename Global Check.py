import pandas as pd
import xlwings as xw
import os
from pathlib import Path
from tkinter import Tk, filedialog, messagebox
import traceback
import sys

try:
    # === Select Input File ===
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
        xls = pd.ExcelFile(input_file)

    if getattr(sys, 'frozen', False):
        # Running in a PyInstaller bundle
        script_dir = Path(sys.executable).parent
    else:
        # Running in normal Python environment
        script_dir = Path(__file__).parent

    global_file = script_dir / "Global.xlsm"
    target_sheet = "Data"

    if not Path(global_file).is_file():
        raise FileNotFoundError(f"❌ File not found: {global_file}")
    else:
        print(f"🔄 Found global file: {global_file}")

    # === Story Definitions ===
    print("📥 Loading: Story Definitions")
    if "Story Definitions" in xls.sheet_names:
        raw_df = pd.read_excel(input_file, sheet_name="Story Definitions", header=None, skiprows=1, usecols="A:E")
        headers = raw_df.iloc[0].tolist()
        units = raw_df.iloc[1].tolist()
        df_story = raw_df.iloc[2:].reset_index(drop=True)
        df_story.columns = headers
        df_story = df_story.loc[:, ~df_story.columns.duplicated()]
        headers = list(df_story.columns)
        print("🧪 Early header check:", headers)  # 👈 加在這裡
        
        units = [u if pd.notna(u) and not str(u).startswith("Unnamed") else "" for u in units]

        # === Load Base Elevation ===
        print("📥 Reading base elevation...")
        try:
            base_sheet_name = next((s for s in xls.sheet_names if s.strip() == "Tower and Base Story Definition"), None)

            if base_sheet_name:
                base_df = pd.read_excel(input_file, sheet_name=base_sheet_name, header=None)
                print("📋 Preview (row 4):", base_df.iloc[3].tolist())  # Debug：印出第 4 列

                # ✅ 找出 "BSElev" 欄位的索引位置（通常在第 5~7 行）
                header_row = base_df.iloc[1].tolist()
                if "BSElev" in header_row:
                    elev_col_idx = header_row.index("BSElev")
                    base_elevation = pd.to_numeric(base_df.iloc[3, elev_col_idx], errors="coerce")
                    if pd.isna(base_elevation):
                        print("⚠️ Base Elevation is NaN. Using 0.0 fallback.")
                        base_elevation = 0.0
                    else:
                        print(f"✅ Base Elevation loaded from row 4: {base_elevation}")
                else:
                    print("❌ 'BSElev' column not found in header row.")
                    base_elevation = 0.0
            else:
                print("⚠️ Sheet 'Tower and Base Story Definition' not found.")
                base_elevation = 0.0
        except Exception as e:
            print(f"⚠️ Failed to read base elevation: {e}")
            base_elevation = 0.0

        # === Compute Elevation
        if "Height" in df_story.columns:
            df_story["Height"] = pd.to_numeric(df_story["Height"], errors="coerce")
            if df_story["Height"].dropna().empty:
                print("⚠️ All Height values are NaN.")
                df_story["Elevation"] = None
                headers.append("Elevation")
                units.append("")
            else:
                df_story = df_story[::-1].reset_index(drop=True)
                df_story["Elevation"] = df_story["Height"].cumsum() + base_elevation
                df_story = df_story[::-1].reset_index(drop=True)

                # ✅ 先 append Elevation header，再用 header 的 index 對應單位
                headers.append("Elevation")
                if "Height" in headers:
                    height_index = headers.index("Height")
                    height_unit = units[height_index] if height_index < len(units) else ""
                else:
                    height_unit = ""

                units.append(height_unit)

                print("✅ Elevation calculated from base elevation.")
                
        else:
            print("❌ 'Height' column not found.")
            df_story["Elevation"] = None
            headers.append("Elevation")
            units.append("")
    else:
        print("⚠️ 'Story Definitions' sheet not found.")
        df_story = None

    # === Generic function to load MultiIndex Excel with units ===
    def load_multiindex_sheet(name):
        if name in xls.sheet_names:
            df = pd.read_excel(input_file, sheet_name=name, skiprows=1, header=[0, 1])
            df.columns = pd.MultiIndex.from_tuples([
                (a if not str(a).startswith("Unnamed") else "", b if not str(b).startswith("Unnamed") else "")
                for a, b in df.columns
            ])
            return df
        else:
            print(f"⚠️ Sheet '{name}' not found. Skipping.")
            return None

    df_modal = load_multiindex_sheet("Modal Participating Mass Ratios")
    df_drift = load_multiindex_sheet("Story Drifts")
    df_diaphragm = load_multiindex_sheet("Diaphragm Max Over Avg Drifts")
    df_force = load_multiindex_sheet("Story Forces")
    df_joint = load_multiindex_sheet("Joint Displacements")
    df_cm = load_multiindex_sheet("Diaphragm CM Displacements")
    df_stiffness = load_multiindex_sheet("Story Stiffness")
    df_jointdrift = load_multiindex_sheet("Joint Drifts")
    df_base = load_multiindex_sheet("Base Reactions")

    # === Excel Helper ===
    def col_letter(idx):
        return xw.utils.col_name(idx)

    def col_name_to_number(col_str):
        col_str = col_str.upper()
        exp = 0
        col_num = 0
        for char in reversed(col_str):
            col_num += (ord(char) - ord('A') + 1) * (26 ** exp)
            exp += 1
        return col_num
    
    def write_block(ws, cell, title, df, name, units=None):
        print(f"✍️  Writing '{title}' to Excel...")
        start_col = ws.range(cell).column
        start_row = ws.range(cell).row
        row1, row2, data_start = start_row + 1, start_row + 2, start_row + 3

        ws.range(cell).value = title
        ws.range(cell).api.Font.Bold = True

        # Write header
        ws.range((row1, start_col)).value = df.columns.tolist() if not isinstance(df.columns, pd.MultiIndex) else df.columns.get_level_values(0).tolist()
        # Write unit row
        if isinstance(df.columns, pd.MultiIndex):
            ws.range((row2, start_col)).value = df.columns.get_level_values(1).tolist()
        elif units:
            ws.range((row2, start_col)).value = units
        else:
            ws.range((row2, start_col)).value = [""] * len(df.columns)

        # Write data
        ws.range((data_start, start_col)).value = df.values.tolist()
        
        # Define the range: from the header row (row1) to the last data row
        end_row = data_start + len(df) - 1
        end_col = start_col + len(df.columns) - 1
        full_range = ws.range((row1, start_col), (end_row, end_col))
        # Light blue background fill: applied only to the header and unit rows
        ws.range((row1, start_col), (row2, end_col)).color = (198, 239, 255)
        # Apply borders to the entire block (header + unit + data)
        for i in range(7, 13):  # xlEdgeLeft to xlInsideHorizontal
            full_range.api.Borders(i).LineStyle = 1  # xlContinuous
            full_range.api.Borders(i).Weight = 2     # xlThin
        # Center-align all cells within the block
        full_range.api.HorizontalAlignment = -4108  # xlCenter
        full_range.api.VerticalAlignment = -4108    # xlCenter
        
        # Define named range from header row (row1) to data end
        ref = f"{target_sheet}!${col_letter(start_col)}${row1}:${col_letter(end_col)}${end_row}"
        if name in [n.name for n in wb.names]:
            wb.names[name].delete()
        wb.names.add(name, refers_to=f"={ref}")
        
    def create_placeholder_from_range(rng):
        from_col = rng.split(":")[0]
        to_col = rng.split(":")[1]
        col_count = col_name_to_number(to_col) - col_name_to_number(from_col) + 1
        return pd.DataFrame(columns=[""] * col_count)
    
    # === Define all data blocks: (title, anchor cell, Excel range, variable name) ===
    name_mapping = {
    "df_modal": "ModalMassRatios",
    "df_drift": "StoryDrifts",
    "df_diaphragm": "DiaphragmMaxOverAvgDrifts",
    "df_force": "StoryForces",
    "df_joint": "JointDisplacements",
    "df_cm": "DiaphragmCMDisplacements",
    "df_stiffness": "StoryStiffness",
    "df_jointdrift": "JointDrifts",
    "df_base": "BaseReactions"
    }

    table_blocks = [
        ("Story Definitions",      "A1",  "A:F",    "df_story"),
        ("Modal Participating Mass Ratios", "H1",  "H:V",    "df_modal"),
        ("Story Drift",            "X1",  "X:AH",   "df_drift"),
        ("Diaphragm Max Over Avg Drifts", "AJ1", "AJ:AU",  "df_diaphragm"),
        ("Story Forces",           "AW1", "AW:BG",  "df_force"),
        ("Joint Displacements",    "BI1", "BI:BT",  "df_joint"),
        ("Diaphragm Center Of Mass Displacements", "BV1", "BV:CG", "df_cm"),
        ("Story Stiffness",        "CI1", "CI:CS",  "df_stiffness"),
        ("Joint Drifts",           "CU1", "CU:DD",  "df_jointdrift"),
        ("Base Reactions",         "DF1", "DF:DQ", "df_base"),
    ]

    # === Start Excel Write ===
    app = xw.App(visible=False)
    wb = xw.Book(str(global_file))
    ws = wb.sheets[target_sheet]

    # Clear all predefined blocks
    for _, _, rng, _ in table_blocks:
        ws.range(rng).clear()

    # Write Story Definitions (special case with units)
    if df_story is not None:
        write_block(ws, "A1", "Story Definitions", df_story, "StoryDefinitions", units)
    else:
        print("📌 Writing placeholder for Story Definitions (sheet missing).")
        story_rng = next(r for t, _, r, v in table_blocks if v == "df_story")
        placeholder = create_placeholder_from_range(story_rng)
        write_block(ws, "A1", "Story Definitions", placeholder, "StoryDefinitions")

    # Write all remaining blocks (auto loop)
    for title, cell, rng, var_name in table_blocks:
        if var_name == "df_story":
            continue  # Already handled above
        df = globals().get(var_name)
        print(f"📝 Writing block: {title}")
        if df is not None:
            write_block(ws, cell, title, df, name_mapping.get(var_name, var_name))
        else:
            print(f"📌 Writing placeholder for {title} (sheet missing).")
            from_column = rng.split(":")[0]
            to_column = rng.split(":")[1]
            col_count = col_name_to_number(to_column) - col_name_to_number(from_column) + 1
            placeholder = create_placeholder_from_range(rng)
            write_block(ws, cell, title, placeholder, name_mapping.get(var_name, var_name))

    ws.range("A:CS").autofit()
    ws.range("A:CS").api.HorizontalAlignment = -4108
    ws.range("A:CS").api.VerticalAlignment = -4108

    for _, cell, _, _ in table_blocks:
        ws.range(cell).api.Font.Bold = True

    # ✅ Freeze top 3 rows
    ws.api.Activate()
    ws.api.Application.ActiveWindow.SplitRow = 3
    ws.api.Application.ActiveWindow.FreezePanes = True
    
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
