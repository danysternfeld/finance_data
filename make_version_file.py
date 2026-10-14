import subprocess

from github import Github

def create_version_module():
    # read git comit count
    result = subprocess.run(["git", "rev-list", "--count", "HEAD"], capture_output=True, text=True)
    githash = subprocess.run(["git", "rev-parse", "HEAD"], capture_output=True, text=True)
    if result.returncode == 0  and githash.returncode == 0:
        with open("version.py", "w", encoding="utf-8") as file:
            file.write(f"def get_version():\n\treturn {result.stdout}\n")
            file.write(f"def get_hash():\n\treturn '{githash.stdout.strip()}'\n")
            file.close()

if __name__ == "__main__":
    create_version_module()
    