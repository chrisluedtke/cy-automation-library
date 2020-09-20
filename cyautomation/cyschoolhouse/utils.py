import subprocess
import os
import sys

def map_sharepoint_drive():
    try:
        subprocess.run(
            r'net use z: /del /Y', 
            shell=True, 
            check=True, 
            capture_output=True
        )
    except subprocess.CalledProcessError:
        pass

    try:
        subprocess.run(
            f'net use z: {os.environ["SHAREPOINT_URL"]}', 
            shell=True, 
            check=True, 
            capture_output=True
        )
    except subprocess.CalledProcessError as e:
        print(e.stderr)
        sys.exit(1)
