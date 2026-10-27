# SMART-WIP: RPA Bot for transferring STT bulk to Master WIP Excel

An enterprise RPA Bot written in **Java (JDK 11+)** following the **3-Layer Architecture** pattern and powered by **Apache POI**.


---

## 1. Key Features

### Intelligent Dynamic Column Mapping
- Automatically inspects header row keywords (`JOB ERP`, `Fabric`, `Style`, `Color`, `Remarks`, etc.).
- Dynamically adapts when factory planners shift or insert columns between weekly WIP templates (e.g., Fabric ETA was at column `AG` in Week 39, but shifted to column `AF` in Week 42/43).

### 100% Non-Destructive Source Protection
- **D365 STT File**: Opened in strict read-only mode via `PackageAccess.READ`.
- **Master WIP Source File**: Read strictly through a standard `InputStream`. All modifications are written exclusively to a brand-new file: `Updated_WIP.xlsx`. The source file is **never overwritten, locked, or modified**.

### Formula Integrity Guardrails
- **Write Whitelist**: The bot is strictly restricted to writing into 4 columns: `Fabric ETA`, `Remarks`, `Material_Short_Qty`, and `Bot_Note`.
- **Formula Overwrite Prevention**: Evaluates `CellType.FORMULA`. If a target cell contains an Excel formula, write execution is aborted and throws a `BotException("FORMULA_PROTECTED")`.
- **Integrity Snapshot**: `protectedSnapshot()` captures the exact values and styles of sensitive protected columns (`AH, AI, AJ, BM, CG`) before and after execution to guarantee zero data corruption.

### Big Data Handling via Apache POI
- Configured with `IOUtils.setByteArrayMaxOverride(1_000_000_000)` (1 GB) to seamlessly read large D365 STT files (> 25 MB with over **75,000 rows**) without triggering POI's `RecordFormatException` (zip bomb safety limit).

---

## 2. System Architecture

```
├── pom.xml                                  # Maven configuration (Main: com.smartwip.gui.App)
├── README.md                                # Project documentation
├── src/main/java/com/smartwip/
│   ├── dal/                                 # DATA ACCESS LAYER
│   │   ├── WipGrid.java                     # Abstract interface for WIP spreadsheet operations
│   │   ├── PoiWipGrid.java                  # Apache POI Excel implementation, Guardrails, Cell styles
│   │   └── FileManager.java                 # Auto file-discovery, STT reader, UTF-8 BOM CSV writers
│   │
│   ├── bll/                                 # BUSINESS LOGIC LAYER
│   │   ├── SmartWipService.java             # Key normalization, Remark 355 regex parser, ETA calculator
│   │   └── model/                           # Domain Models
│   │       ├── BotException.java            # Domain exceptions & guardrail violations
│   │       ├── SttLine.java                 # Single line model from D365 STT
│   │       ├── Group.java                   # Aggregated STT lines mapped to 1 WIP row
│   │       ├── Remark.java                  # Parsed model for Remark 355
│   │       ├── LogEntry.java                # Audit & warning record
│   │       ├── ChangeRow.java               # Detailed record of every cell mutation
│   │       └── Result.java                  # Aggregated execution metrics
│   │
│   └── gui/                                 # PRESENTATION / CLI LAYER
│       └── App.java                         # CLI argument parsing, workflow orchestration
```

---

## 3. Business Rules (R1 - R8)

| Rule ID | Rule Name | Specification Details |
| :---: | :--- | :--- |
| **R1** | **Remark 355 Parsing** | Extracted via RegEx: `^[FU\|BL]\s+[AF\|SF]\s+ETD\s+(\d{1,2})/(\d{1,2})(?:/(\d{4}))?$`. If invalid syntax $\rightarrow$ logs `SYNTAX_ERROR` and falls back to ERP `Final ETA Date`. |
| **R2** | **Transit Lead-Time** | $\text{ETA} = \text{ETD} + \text{Transit Days}$. Sea Freight (**SF**) = $+25$ days; Air Freight (**AF**) = $+5$ days. |
| **R3** | **Composite Match Key** | Maps 1:N relations between STT and WIP via composite key: `JOB_ERP + STYLE + COLOR` (normalized via `normStyle` and `normColor`). Only updates WIP rows where status is `"OPEN"`. |
| **R4** | **Maximum ETA Aggregation** | $\text{ETA}_{\text{final}} = \max(\text{ETA of all POs belonging to the order})$. Garment sewing can only start after the latest fabric arrives. If D365 has no ETA $\rightarrow$ retain existing WIP date and report `NO_ETA`. |
| **R5** | **Preserving Manual Remarks** | The `Remarks` column appends parsed Remark 355 strings (e.g. ` \| FU SF ETD...`). Never overwrites merchandiser manual comments, and avoids duplicate appended strings. |
| **R6** | **Material Shortage** | $\text{Shortage} = \sum (\text{Require Qty} - \text{Reserve Qty})$ across lines with unit `YDS` (fabric). Stored in the dedicated column `Material_Short_Qty (YDS)`, leaving `Quantity short ship` intact. |
| **R7** | **Visual Cell Highlighting** | Cells created or updated by the Bot are highlighted with a light yellow background (`IndexedColors.LIGHT_YELLOW`). |
| **R8** | **Audit Trail & Reporting** | Discrepancies, conflicts, and syntax errors are logged into `Exception_Log.csv` and `Bot_Changes.csv`. Non-WIP sample orders (`TS/PA`) are logged with level `INFO`. |

---

## 4. Environment Requirements & Setup

- **Operating System**: Windows 10/11, macOS, or Linux.
- **Java Development Kit (JDK)**: Version **11** or newer.
- **Apache Maven**: Version **3.8+** (or IntelliJ IDEA bundled Maven).

Verify your environment in terminal / command prompt:
```bash
java -version
mvn -version
```

---

## 5. How to Run

### Method 1: Compile and Run via Maven CLI

1. **Compile source code:**
   ```powershell
   mvn clean compile
   ```

2. **Run in headless / auto mode (no prompt, no auto-open of Excel):**
   ```powershell
   mvn exec:java "-Dexec.args=--auto --no-open"
   ```

3. **Specify custom input file paths:**
   ```powershell
   mvn exec:java "-Dexec.args=--stt=\"D:/Data/STT_OCT.xlsx\" --wip=\"D:/Data/WIP_WK43.xlsx\" --auto"
   ```

4. **Simulate a Purchasing Remark 355 update for a specific PO:**
   ```powershell
   mvn exec:java "-Dexec.args=--remark=POPU0022923:FU_SF_ETD_20/10/2026 --auto"
   ```

---

### Method 2: Package into a Standalone Executable (Fat JAR)

1. **Build the Fat JAR package:**
   ```powershell
   mvn clean package
   ```
   *(The standalone artifact `smart-wip-demo.jar` will be generated in the `target/` directory)*.

2. **Run the JAR (Allocating 4GB RAM is recommended for large enterprise workbooks):**
   ```powershell
   java -Xmx4g -jar target/smart-wip-demo.jar --auto
   ```

---

## 6. Input & Output Data Structure

### Input Files:
- **`STT BULK *.xlsx`**: D365 ERP export containing Sales Order, Style, Color, PO, Final ETA Date, Require/Reserve Qty, and Remark 355.
- **`WK *-WIP_*.xlsx`**: Master WIP Excel spreadsheet tracking production schedules across factories.

### Output Files Generated:
1. **`Updated_WIP.xlsx`**: Resulting WIP spreadsheet with updated ETA dates (highlighted in light yellow), appended remarks, and 2 additional trailing columns:
   - `Material_Short_Qty (YDS)`: Total fabric shortage.
   - `Bot_Note`: Detailed update reasons, lead-time variance (+/- days), and bot run timestamp.
2. **`Bot_Changes.csv`** *(UTF-8 with BOM format)*: Detailed audit trail for every modified WIP row:
   - Columns: `WIP_Excel_Row, JOB_ERP, STYLE, COLOR, ETA_old, ETA_new, Diff_days, Action, ETA_source, PO_sample, Short_YDS, Note`.
3. **`Exception_Log.csv`** *(UTF-8 with BOM format)*: Comprehensive list of remark syntax errors, out-of-scope orders, missing ERP ETAs, or status discrepancies.

---

## 7. Troubleshooting Common Issues

### 1. `FileSystemException: The process cannot access the file because it is being used by another process`
- **Cause**: `Updated_WIP.xlsx` is currently open in Microsoft Excel or another viewer.
- **Resolution**: Close the file in Excel before triggering the bot run.

### 2. `OutOfMemoryError: Java heap space`
- **Cause**: Large D365 STT or WIP workbooks (> 20 MB with over 75,000 rows) exceeding JVM default memory limits.
- **Resolution**: Increase heap size for Maven or the Java process:
  - In Windows PowerShell: `$env:MAVEN_OPTS="-Xmx4g"`
  - In Windows CMD: `set MAVEN_OPTS=-Xmx4g`
  - In Linux / macOS: `export MAVEN_OPTS="-Xmx4g"`
