from pathlib import Path
import subprocess
import time

desktop = Path.home() / "Desktop"
excel_file = desktop / "Test -Overseas Leave Application.xlsx"
export_script = desktop / "export_json.py"

CHECK_EVERY_SECONDS = 2

def run_export():
    try:
        result = subprocess.run(
            ["py", str(export_script)],
            capture_output=True,
            text=True
        )
        print(result.stdout.strip())
        if result.stderr.strip():
            print("ERROR:", result.stderr.strip())
    except Exception as e:
        print("Failed to run export_json.py:", e)

def main():
    if not excel_file.exists():
        print("Excel file not found:", excel_file)
        return

    if not export_script.exists():
        print("export_json.py not found:", export_script)
        return

    print("Watching:", excel_file)
    print("Will run:", export_script)
    print("Press Ctrl + C to stop.\n")

    last_mtime = excel_file.stat().st_mtime

    # Optional: run once at startup
    run_export()

    while True:
        try:
            current_mtime = excel_file.stat().st_mtime
            if current_mtime != last_mtime:
                last_mtime = current_mtime
                print("\nChange detected. Regenerating data.json...")
                time.sleep(1)  # small delay so Excel finishes saving
                run_export()

            time.sleep(CHECK_EVERY_SECONDS)

        except KeyboardInterrupt:
            print("\nStopped watching.")
            break
        except PermissionError:
            print("Excel file is busy. Will retry...")
            time.sleep(CHECK_EVERY_SECONDS)
        except Exception as e:
            print("Watcher error:", e)
            time.sleep(CHECK_EVERY_SECONDS)

if __name__ == "__main__":
    main()