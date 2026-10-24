import os
import glob
import subprocess

user_data = os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\MQAA_Automation")
locks = glob.glob(os.path.join(user_data, "*lock*"))
print("Found locks:", locks)
for l in locks:
    try:
        os.remove(l)
        print("Removed:", l)
    except Exception as e:
        print("Could not remove:", l, e)

chrome_exe = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
cmd = [
    chrome_exe,
    "--remote-debugging-port=9222",
    f"--user-data-dir={user_data}",
    "https://sfc.chingluh.com/vhmes/QA/QCHFPA.aspx?tree_id=Q504"
]
print("Launching Chrome...")
subprocess.Popen(cmd)
