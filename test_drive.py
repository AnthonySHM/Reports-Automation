"""Scan ALL client folders for NDR files to understand naming conventions."""
import os
os.environ["SHM_GOOGLE_CREDENTIALS"] = (
    r"C:\Users\User\Desktop\Reports Automation\Python Engine"
    r"\Credentials\focal-set-486609-s9-f052f8c1a756.json"
)

from core.drive_agent import DriveAgent

agent = DriveAgent()
client_folders = agent._list_folders(agent.root_folder_id)

for cf in client_folders:
    if cf["name"] in ("Test", "Archive of completed reports"):
        continue
    print(f"\n=== {cf['name']} ===")
    subfolders = agent._list_folders(cf["id"])
    ndr_folders = [f for f in subfolders if f["name"].upper().startswith("NDR")]
    if not ndr_folders:
        print("  (no NDR folders)")
        continue
    for ndr in ndr_folders:
        print(f"  NDR folder: {ndr['name']}")
        files = agent._list_image_files(ndr["id"])
        if files:
            for f in files:
                print(f"    - {f['name']}")
        else:
            # Check all files
            all_files = agent.service.files().list(
                q=f"'{ndr['id']}' in parents and trashed=false",
                fields="files(id, name, mimeType)",
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
            ).execute().get("files", [])
            if all_files:
                for f in all_files:
                    print(f"    - {f['name']} ({f['mimeType']})")
            else:
                print("    (empty)")
