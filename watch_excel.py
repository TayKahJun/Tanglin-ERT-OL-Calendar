from pathlib import Path
import subprocess
import time

base_dir = Path(__file__).resolve().parent

excel_file = base_dir / "Test -Overseas Leave Application.xlsx"
export_script = base_dir / "export_json.py"

CHECK_EVERY_SECONDS = 2

def run_command(command, description):
    try:
        print(f"\n{description}...")
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            cwd=base_dir
        )

        if result.stdout.strip():
            print(result.stdout.strip())

        if result.stderr.strip():
            print(result.stderr.strip())

        return result.returncode == 0

    except Exception as e:
        print(f"Failed during {description}: {e}")
        return False

def run_export():
    return run_command(["py", str(export_script)], "Running export_json.py")

def git_push():
    # Stage updated files
    if not run_command(["git", "add", "data.json"], "Staging data.json"):
        return

    # Check whether there is actually anything to commit
    result = subprocess.run(
        ["git", "status", "--porcelain"],
        capture_output=True,
        text=True,
        cwd=base_dir
    )

    changed_files = result.stdout.strip()
    if not changed_files:
        print("No Git changes detected. Nothing to commit.")
        return

    # Commit
    commit_message = f"Auto update data.json at {time.strftime('%Y-%m-%d %H:%M:%S')}"
    if not run_command(["git", "commit", "-m", commit_message], "Committing changes"):
        return

    # Push
    if not run_command(["git", "push", "origin", "main"], "Pushing to GitHub"):
        return

    print("GitHub push completed successfully.")

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

    # Run once at start
    if run_export():
        git_push()

    while True:
        try:
            current_mtime = excel_file.stat().st_mtime

            if current_mtime != last_mtime:
                last_mtime = current_mtime
                print("\nChange detected. Regenerating data.json...")
                time.sleep(1)  # allow Excel to finish saving

                if run_export():
                    git_push()

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