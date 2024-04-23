import subprocess
import json
import sys


PACKAGE_FILE = "../RapidConfiguration/installApps.json"


def check_winget():
    """Check if the Windows Package Manager is installed"""

    try:
        subprocess.run(["winget", "--version"], check=True)
        return True
    except FileNotFoundError:
        print("Winget is not installed")
        return False


def check_package(package):
    """Check if a package exists in the Windows Package Manager"""

    try:
        subprocess.run(["winget", "show", "--accept-source-agreements", "--id", package], check=True)
    except subprocess.CalledProcessError:
        print(f"Package '{package}' is missing from Winget")


def unnest_json(json_obj):
    """Un-nest a JSON object"""

    unnested = []
    for key, value in json_obj.items():
        if isinstance(value, list):
            unnested.extend(value)
        elif isinstance(value, dict):
            unnested.extend(unnest_json(value))
    return unnested


def main():
    """Main function"""

    # Check if Winget is installed
    if not check_winget():
        sys.exit(1)

    # Load the JSON data from the file
    with open(PACKAGE_FILE, "r") as file:
        packages_json = json.load(file)

        # Un-nest all elements
        packages = unnest_json(packages_json)
        print(packages)

        # Check each package
        failed = False
        for package in packages:
            if not check_package(package):
                failed = True

        if failed:
            sys.exit(1)
        else:
            sys.exit(0)


if __name__ == "__main__":
    main()
