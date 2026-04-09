import subprocess
import sys

repo_path = r"c:\Users\mariannafulcher\Downloads\procard_app"

commands = [
    (["git", "-C", repo_path, "add", "app.py"], "Add app.py"),
    (["git", "-C", repo_path, "commit", "-m", "Make tkinter optional for Streamlit Cloud deployment"], "Commit"),
    (["git", "-C", repo_path, "pull"], "Pull"),
    (["git", "-C", repo_path, "push", "origin", "main"], "Push"),
]

for cmd, desc in commands:
    print(f"\n{desc}...")
    result = subprocess.run(cmd, capture_output=True, text=True)
    print(result.stdout)
    if result.stderr:
        print("STDERR:", result.stderr)
    if result.returncode != 0:
        print(f"ERROR: Command failed with return code {result.returncode}")
        sys.exit(1)

print("\n✓ All done!")
