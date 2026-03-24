from pathlib import Path
import pandas as pd
import json

base_dir = Path(__file__).resolve().parent

excel_file = base_dir / "Test -Overseas Leave Application.xlsx"
json_file = base_dir / "data.json"

def convert_value(value):
    if pd.isna(value):
        return ""
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")
    return value

def main():
    print("Current working folder:", base_dir)
    print("Looking for Excel file at:", excel_file)
    print("Will save JSON to:", json_file)

    if not excel_file.exists():
        print("ERROR: Excel file not found.")
        return

    try:
        df = pd.read_excel(excel_file)
        records = []

        for _, row in df.iterrows():
            record = {}
            for col in df.columns:
                record[str(col)] = convert_value(row[col])
            records.append(record)

        with open(json_file, "w", encoding="utf-8") as f:
            json.dump(records, f, ensure_ascii=False, indent=2)

        print(f"JSON exported successfully: {json_file}")
        print(f"Total rows exported: {len(records)}")

    except Exception as e:
        print("ERROR:", e)

if __name__ == "__main__":
    main()
