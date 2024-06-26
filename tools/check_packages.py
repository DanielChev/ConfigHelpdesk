import subprocess
import json
import sys


PACKAGE_FILE = "../RapidConfiguration/installApps.json"


def check_winget():
    """Check if the Windows Package Manager is installed"""

    try:
        subprocess.run(["winget", "--version"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True
    except FileNotFoundError:
        print("Winget is not installed")
        return False


def check_package(package):
    """Check if a package exists in the Windows Package Manager"""

    try:
        subprocess.run(["winget", "show", "--accept-source-agreements", "--disable-interactivity", "--id", package], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        print(f"Package '{package}' is available")
        return True
    except subprocess.CalledProcessError:
        print(f"Package '{package}' is not available on Winget")
        return False


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

        print("Checking availability of packages...")

        # Check each package
        failed = False
        for package in packages:
            if not check_package(package):
                failed = True

        if failed:
            print("Some packages are not available on Winget")
            sys.exit(1)

        print("All packages are available on Winget")


if __name__ == "__main__":
    main()
