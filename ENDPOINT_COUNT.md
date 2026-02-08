# Endpoint Deployment and Agent Counts (SIEM, XDR, DFIR)

## Overview

The report generation system now automatically populates **endpoint deployment count** and **agent counts** on **Slide 4 (Service Coverage & Deployment)** from the inventory report CSV file found in the VN folder.

## Features

### Automatic System Counts

The system extracts four metrics from the inventory report and displays them on Slide 4:

**Original placeholders:**
```
• Endpoint Deployment ([Count] systems detected):
    a. EPIC SIEM: [Count] agents
    b. SentinelOne XDR: [Count] agents
    c. DFIR Agent: [Count] agents
• Network Detection & Response (NDR): Active
• MS365 Cloud Coverage:
    a. Active user accounts: [Count]
• Vulnerability Management:
    a. Host-Based (OS / 3rd Party SW): [Count] systems reporting
```

**After population (with agents):**
```
• Endpoint Deployment (42 systems detected):
    a. EPIC SIEM: 38 agents
    b. SentinelOne XDR: 35 agents
    c. DFIR Agent: 30 agents
• Network Detection & Response (NDR): Active
• MS365 Cloud Coverage:
    a. Active user accounts: [Count]
• Vulnerability Management:
    a. Host-Based (OS / 3rd Party SW): 38 systems reporting
```

**After population (all agents Client Managed):**
```
• Endpoint Deployment (42 systems detected):
    a. EPIC SIEM: Client Managed
    b. SentinelOne XDR: Client Managed
    c. DFIR Agent: Client Managed
• Network Detection & Response (NDR): Active
• MS365 Cloud Coverage:
    a. Active user accounts: [Count]
• Vulnerability Management:
    a. Host-Based (OS / 3rd Party SW): [Count] systems reporting
```

**Note**: 
- The Host-Based vulnerability management count equals the EPIC SIEM agent count, as these represent the same systems
- When any agent count is 0, "Client Managed" is displayed instead of "0 agents" for that agent type
- This applies to SIEM, XDR, and DFIR agents independently

### Data Source

All counts are derived from the same **inventory report CSV** file:

- **File pattern**: `inventoryreport_(client name)-(date).csv` or `InventoryReport_(client name)-(date).csv`
- **Location**: Inside the most recent `VN-*` folder for the client

### Count Methods

1. **Total Endpoints**: Count of all data rows (excluding header)
2. **SIEM Agents**: Count of rows where the `SIEM` column value is `"Installed"` (case-insensitive)
   - Also used for **Host-Based (OS / 3rd Party SW)** count
   - If count = 0, displays **"Client Managed"** instead of "0 agents"
3. **XDR Agents**: Count of rows where the `XDR` column value is `"Installed"` (case-insensitive)
   - If count = 0, displays **"Client Managed"** instead of "0 agents"
4. **DFIR Agents**: Count of rows where the `IR_Agent` column value is `"Installed"` (case-insensitive)
   - If count = 0, displays **"Client Managed"** instead of "0 agents"

## Implementation

### Files Modified

#### `core/drive_agent.py`

**New Method: `fetch_inventory_csv()`**
- Searches for the inventory report CSV in the latest VN folder
- Uses case-insensitive matching for "inventoryreport"
- Downloads the file to the local cache
- Returns the local path or `None` if not found

**New Function: `count_inventory_systems()`**
- Opens the inventory CSV file
- Counts all data rows (excluding the header)
- Returns the total count or `None` if reading fails

**New Function: `count_xdr_installations()`**
- Opens the inventory CSV file
- Finds the `XDR` column (case-insensitive)
- Counts all rows where the value is `"Installed"` (case-insensitive)
- Returns the XDR count or `None` if column not found or reading fails

**New Function: `count_dfir_installations()`**
- Opens the inventory CSV file
- Finds the `IR_Agent` column (case-insensitive, flexible with separators: `IR_Agent`, `IR Agent`, `ir-agent`)
- Counts all rows where the value is `"Installed"` (case-insensitive)
- Returns the DFIR count or `None` if column not found or reading fails

- **Updated Function: `replace_endpoint_count_in_slide()`**
- Now accepts `endpoint_count`, optional `siem_count`, optional `xdr_count`, and optional `dfir_count` parameters
- Targets the `service_coverage_body` shape on Slide 4
- Replaces five placeholders:
  - `[Count] systems detected` → endpoint count
  - `EPIC SIEM: [Count] agents` → SIEM count or **"Client Managed"** if count is 0
  - `SentinelOne XDR: [Count] agents` → XDR count or **"Client Managed"** if count is 0
  - `DFIR Agent: [Count] agents` → DFIR count or **"Client Managed"** if count is 0
  - `Host-Based (OS / 3rd Party SW): [Count] systems reporting` → SIEM count (if provided and > 0)
- **Important**: Host-Based count equals SIEM agent count (they represent the same systems)
- **Special Logic**: When any agent count is 0, replaces `[Count] agents` with `Client Managed`
- Returns the number of replacements made (0-5)
- Saves the PowerPoint file after replacements

#### `api.py`

**Integration Logic:**
1. After AI insight generation, fetch the inventory CSV
2. Count the total systems in the inventory
3. Count the systems with SIEM installed
4. Count the systems with XDR installed
5. Count the systems with DFIR agent installed
6. Replace all placeholders on Slide 4 (including Host-Based with SIEM count)
7. Apply "Client Managed" logic for XDR if count is 0
8. Log success or errors

**Error Handling:**
- Graceful fallback if inventory file is not found
- If SIEM column not found, only endpoint, XDR, and DFIR counts are populated (Host-Based remains as placeholder)
- If XDR column not found, only endpoint, SIEM, DFIR, and Host-Based counts are populated
- If IR_Agent column not found, only endpoint, SIEM, XDR, and Host-Based counts are populated
- Logs warnings without blocking report generation
- Placeholders remain if counts cannot be determined

### Code Flow

```python
# In api.py, after AI insight generation
try:
    inventory_csv = _drive_agent.fetch_inventory_csv(client_name, NDR_CACHE_DIR)
    if inventory_csv:
        endpoint_count = count_inventory_systems(inventory_csv)
        siem_count = count_siem_installations(inventory_csv)
        dfir_count = count_dfir_installations(inventory_csv)
        
        if endpoint_count is not None:
            replacements = replace_endpoint_count_in_slide(
                result_path, endpoint_count, siem_count, dfir_count
            )
            if replacements > 0:
                msg_parts = [f"{endpoint_count} systems"]
                if siem_count is not None:
                    msg_parts.append(f"{siem_count} SIEM agents")
                if dfir_count is not None:
                    msg_parts.append(f"{dfir_count} DFIR agents")
                logger.info("Service coverage counts inserted on slide 4: %s", ", ".join(msg_parts))
except Exception as exc:
    logger.warning("Failed to populate service coverage counts: %s", exc)
```

## Example Inventory CSV

```csv
Hostname,IP Address,OS,SIEM,XDR,IR_Agent,Agent Status,Last Seen
SERVER01,192.168.1.10,Windows Server 2019,Installed,Installed,Installed,Active,2026-02-08
DESKTOP01,192.168.1.20,Windows 10,Installed,Installed,Not Installed,Active,2026-02-08
DESKTOP02,192.168.1.21,Windows 11,Not Installed,Not Installed,Installed,Active,2026-02-07
SERVER02,192.168.1.11,Windows Server 2022,installed,installed,installed,Inactive,2026-02-05
LAPTOP01,192.168.1.30,Windows 11,Installed,not_installed,Installed,Active,2026-02-08
WORKSTATION01,192.168.1.40,Windows 10,,,,Active,2026-02-08
```

**Results**:
- Total endpoints: `6 systems detected` (6 data rows)
- SIEM agents: `4 agents` (4 rows with "Installed" in SIEM column, case-insensitive)
- XDR agents: `3 agents` (3 rows with "Installed" in XDR column, case-insensitive)
- DFIR agents: `4 agents` (4 rows with "Installed" in IR_Agent column, case-insensitive)
- Host-Based systems: `4 systems reporting` (same as SIEM agent count)

**Example with all Client Managed:**
If all agent columns have "Not Installed" or "not_installed" values:
```
a. EPIC SIEM: Client Managed
b. SentinelOne XDR: Client Managed
c. DFIR Agent: Client Managed
```

**Example with mixed:**
If SIEM has 3 installed, XDR has 0, and DFIR has 1:
```
a. EPIC SIEM: 3 agents
b. SentinelOne XDR: Client Managed
c. DFIR Agent: 1 agents
```

## File Name Variations

The system handles various file naming conventions:

- `inventoryreport_clientname_2026-02-08.csv`
- `InventoryReport_ClientName_2026-02-08.csv`
- `inventory_report_clientname_2026-02-08.csv`
- `Inventory Report_ClientName_2026-02-08.csv`

All variations are detected using case-insensitive matching with spaces/underscores normalized.

## Testing

### Unit Test - Endpoint Count

```bash
python -c "
from pathlib import Path
import csv
from core.drive_agent import count_inventory_systems

test_csv = Path('test_inventory.csv')
with open(test_csv, 'w', newline='') as f:
    writer = csv.writer(f)
    writer.writerow(['Host', 'IP', 'SIEM'])
    writer.writerow(['SRV01', '192.168.1.1', 'Installed'])
    writer.writerow(['SRV02', '192.168.1.2', 'Not Installed'])

count = count_inventory_systems(test_csv)
print(f'Total: {count}')  # Should print 2
test_csv.unlink()
"
```

### Unit Test - SIEM Count

```bash
python -c "
from pathlib import Path
import csv
from core.drive_agent import count_siem_installations

test_csv = Path('test_inventory.csv')
with open(test_csv, 'w', newline='') as f:
    writer = csv.writer(f)
    writer.writerow(['Host', 'IP', 'SIEM'])
    writer.writerow(['SRV01', '192.168.1.1', 'Installed'])
    writer.writerow(['SRV02', '192.168.1.2', 'installed'])  # Case-insensitive
    writer.writerow(['SRV03', '192.168.1.3', 'Not Installed'])

count = count_siem_installations(test_csv)
print(f'SIEM: {count}')  # Should print 2
test_csv.unlink()
"
```

### Unit Test - XDR Count

```bash
python -c "
from pathlib import Path
import csv
from core.drive_agent import count_xdr_installations

test_csv = Path('test_inventory.csv')
with open(test_csv, 'w', newline='') as f:
    writer = csv.writer(f)
    writer.writerow(['Host', 'IP', 'XDR'])
    writer.writerow(['SRV01', '192.168.1.1', 'Installed'])
    writer.writerow(['SRV02', '192.168.1.2', 'installed'])  # Case-insensitive
    writer.writerow(['SRV03', '192.168.1.3', 'Not Installed'])

count = count_xdr_installations(test_csv)
print(f'XDR: {count}')  # Should print 2
test_csv.unlink()
"
```

### Unit Test - DFIR Count

```bash
python -c "
from pathlib import Path
import csv
from core.drive_agent import count_dfir_installations

test_csv = Path('test_inventory.csv')
with open(test_csv, 'w', newline='') as f:
    writer = csv.writer(f)
    writer.writerow(['Host', 'IP', 'IR_Agent'])
    writer.writerow(['SRV01', '192.168.1.1', 'Installed'])
    writer.writerow(['SRV02', '192.168.1.2', 'installed'])  # Case-insensitive
    writer.writerow(['SRV03', '192.168.1.3', 'Not Installed'])

count = count_dfir_installations(test_csv)
print(f'DFIR: {count}')  # Should print 2
test_csv.unlink()
"
```

### Integration Test

Generate a report with an inventory file present and verify:
1. Slide 4 shows the correct endpoint count
2. Slide 4 shows the correct SIEM agent count
3. Slide 4 shows the correct XDR agent count (or "Client Managed" if 0)
4. Slide 4 shows the correct DFIR agent count
5. Slide 4 shows the correct Host-Based count (should equal SIEM count)
6. No `[Count]` placeholders remain
7. Logs show successful replacements

## Troubleshooting

### Counts Not Populated

**Symptoms**: `[Count]` placeholders still visible on Slide 4

**Possible causes**:

1. **Inventory file not found**
   - Check file naming convention
   - Verify file exists in the latest VN-* folder
   
2. **CSV reading error**
   - Check file encoding (should be UTF-8 or UTF-8-BOM)
   - Verify CSV has valid structure

3. **Agent columns not found**
   - Endpoint count will still populate
   - Only missing agent counts will remain as `[Count]`
   - Check CSV has columns named "SIEM" and/or "IR_Agent" (case-insensitive)
   - IR_Agent column accepts variations: `IR_Agent`, `IR Agent`, `ir-agent`

4. **Shape name mismatch**
   - Ensure slide 4 has a shape named `service_coverage_body`

**Check logs**:
```
INFO: Downloaded inventory CSV 'inventoryreport_client_2026-02-08.csv' -> ...
INFO: Counted 42 systems in inventory from ...
INFO: Counted 38 SIEM installations from ...
INFO: Counted 35 DFIR agent installations from ...
INFO: Replaced endpoint count [Count] -> 42 in shape 'service_coverage_body'
INFO: Replaced SIEM count [Count] -> 38 in shape 'service_coverage_body'
INFO: Replaced DFIR count [Count] -> 35 in shape 'service_coverage_body'
INFO: Replaced Host-Based count [Count] -> 38 (SIEM count) in shape 'service_coverage_body'
INFO: Service coverage counts inserted on slide 4: 42 systems, 38 SIEM agents, 35 DFIR agents
```

### Agent Column Not Found

If a specific agent column is missing from the inventory CSV:
```
INFO: Column 'SIEM' not found in ... (columns: ['Hostname', 'IP', 'OS', 'IR_Agent'])
```
or
```
INFO: Column 'IR_Agent' not found in ... (columns: ['Hostname', 'IP', 'OS', 'SIEM'])
```

The endpoint count and any found agent counts will still be populated, but missing agent count placeholders will remain.

**Important**: If SIEM column is not found, the Host-Based count will also remain as a placeholder since it depends on the SIEM count.

### File Not Found

If the inventory file is missing:
```
WARNING: No inventory report CSV found in VN folder 'VN-ClientName-2026-02-08'
WARNING: Failed to populate service coverage counts: ...
```

The report will still generate successfully with both placeholders intact.

## Notes

- All counts work for single-sensor and multi-sensor report configurations
- Agent matching is case-insensitive: "Installed", "installed", "INSTALLED" all count
- Empty agent values or other values (e.g., "Not Installed", "Pending") are excluded
- Endpoint count reflects total systems regardless of any agent status
- Agent counts are optional - if a column is missing, only available counts are populated
- The IR_Agent column accepts flexible naming: `IR_Agent`, `IR Agent`, `ir-agent`, `iragent`
- **Host-Based count equals SIEM agent count** - these represent the same systems with vulnerability scanning capabilities
- **XDR "Client Managed" logic**: When XDR count is 0, "Client Managed" is displayed to indicate XDR is managed by the client rather than having no coverage
- **SIEM "Client Managed" logic**: When SIEM count is 0, "Client Managed" is displayed to indicate SIEM is managed by the client
- **DFIR "Client Managed" logic**: When DFIR count is 0, "Client Managed" is displayed to indicate DFIR is managed by the client
- Each agent type independently shows "Client Managed" when its count is 0
- The feature runs after KB count and AI insight population
- Independent of other VN-folder extractions (patches, vulnerabilities)
