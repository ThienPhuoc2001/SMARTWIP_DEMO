package com.smartwip.dal;

import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.smartwip.bll.model.BotException;

/**
 * WipGrid implementation that operates directly on an Excel workbook via Apache POI.
 * Includes a Guardrail mechanism that prevents overwriting data/formula cells,
 * and auto-detects column positions from header names (Dynamic Header Mapping).
 */
public class PoiWipGrid implements WipGrid {

    public static final String SHEET_NAME = "WIP";
    public static final String NEW_COL_SHORT = "Material_Short_Qty (YDS)";
    public static final String NEW_COL_NOTE = "Bot_Note";

    // Default column positions (fallback when not found by keyword)
    public static final int WCOL_STYLE_DEFAULT = col("I");
    public static final int WCOL_COLOR_DEFAULT = col("J");
    public static final int WCOL_ETA_DEFAULT = col("AG");
    public static final int WCOL_ACCS_DEFAULT = col("AH");
    public static final int WCOL_LATEST_DEFAULT = col("AI");
    public static final int WCOL_BUFFER_DEFAULT = col("AJ");
    public static final int WCOL_MATSTATUS_DEFAULT = col("BN");
    public static final int WCOL_JOB_DEFAULT = col("BT");
    public static final int WCOL_STATUS_DEFAULT = col("BY");
    public static final int WCOL_REMARKS_DEFAULT = col("BZ");
    public static final int WCOL_SHORTSHIP_DEFAULT = col("CI");

    private final Sheet sheet;
    private final GridStyles styles;
    private final DataFormatter fmt = new DataFormatter();
    private final int[] rows;
    private final int headerRow;

    private final int colStyle;
    private final int colColor;
    private final int colEta;
    private final int colAccs;
    private final int colLatest;
    private final int colBuffer;
    private final int colMatStatus;
    private final int colJob;
    private final int colStatus;
    private final int colRemarks;
    private final int colShortShip;
    private final int colShort;
    private final int colNote;

    private final Set<Integer> writable = new HashSet<>();

    private PoiWipGrid(Sheet sheet, GridStyles styles, int[] rows, int headerRow,
                       int colStyle, int colColor, int colEta, int colAccs, int colLatest, int colBuffer,
                       int colMatStatus, int colJob, int colStatus, int colRemarks, int colShortShip,
                       int colShort, int colNote) {
        this.sheet = sheet;
        this.styles = styles;
        this.rows = rows;
        this.headerRow = headerRow;

        this.colStyle = colStyle;
        this.colColor = colColor;
        this.colEta = colEta;
        this.colAccs = colAccs;
        this.colLatest = colLatest;
        this.colBuffer = colBuffer;
        this.colMatStatus = colMatStatus;
        this.colJob = colJob;
        this.colStatus = colStatus;
        this.colRemarks = colRemarks;
        this.colShortShip = colShortShip;
        this.colShort = colShort;
        this.colNote = colNote;

        // Build the whitelist of columns the bot is allowed to update
        writable.add(colEta);
        writable.add(colRemarks);
        writable.add(colShort);
        writable.add(colNote);
    }

    /**
     * Open and initialize a PoiWipGrid from an XSSFWorkbook.
     */
    public static PoiWipGrid open(XSSFWorkbook wb) {
        Sheet sh = wb.getSheet(SHEET_NAME);
        if (sh == null) {
            throw new IllegalStateException("Không tìm thấy sheet '" + SHEET_NAME + "' trong file WIP.");
        }

        DataFormatter df = new DataFormatter();
        int hr = -1;

        // Find the header row (scan the first 40 rows for a cell containing 'job erp')
        for (int r = 0; r <= 40; r++) {
            Row row = sh.getRow(r);
            if (row == null) {
                continue;
            }
            for (int c = 0; c < row.getLastCellNum(); c++) {
                String t = norm(cellText(df, row.getCell(c)));
                if (t.contains("job erp") || "job erp".equals(t)) {
                    hr = r;
                    break;
                }
            }
            if (hr >= 0) {
                break;
            }
        }

        if (hr < 0) {
            throw new IllegalStateException("Không tìm thấy dòng tiêu đề (chứa 'JOB ERP') trong 40 dòng đầu của sheet WIP.");
        }

        Row h = sh.getRow(hr);

        // Dynamically map columns by header keywords on the header row (guards against column shifts between weekly WIP files)
        int colJob = findCol(h, df, "job erp", new String[]{"job erp"}, null, WCOL_JOB_DEFAULT);
        int colStyle = findCol(h, df, "style no", new String[]{"style"}, null, WCOL_STYLE_DEFAULT);
        int colColor = findCol(h, df, "color", new String[]{"color"}, new String[]{"desc"}, WCOL_COLOR_DEFAULT);
        int colEta = findCol(h, df, null, new String[]{"fabric"}, null, WCOL_ETA_DEFAULT);
        int colAccs = findCol(h, df, null, new String[]{"accs"}, null, WCOL_ACCS_DEFAULT);
        int colLatest = findCol(h, df, null, new String[]{"latest in house"}, null, WCOL_LATEST_DEFAULT);
        int colBuffer = findCol(h, df, null, new String[]{"buffer"}, null, WCOL_BUFFER_DEFAULT);
        int colMatStatus = findCol(h, df, "material status", new String[]{"material status"}, null, WCOL_MATSTATUS_DEFAULT);
        int colStatus = findCol(h, df, "order status", new String[]{"order status"}, null, WCOL_STATUS_DEFAULT);
        int colRemarks = findCol(h, df, "remarks", new String[]{"remark"}, new String[]{"short", "delay"}, WCOL_REMARKS_DEFAULT);
        int colShortShip = findCol(h, df, "quantity short ship", new String[]{"short ship"}, new String[]{"remark"}, WCOL_SHORTSHIP_DEFAULT);

        System.out.println("  [TỰ ĐỘNG NHẬN DIỆN CỘT] Header tại dòng " + (hr + 1)
                + " | JOB: " + CellReference.convertNumToColString(colJob)
                + ", Style: " + CellReference.convertNumToColString(colStyle)
                + ", Color: " + CellReference.convertNumToColString(colColor)
                + ", Fabric_ETA: " + CellReference.convertNumToColString(colEta)
                + ", Accs: " + CellReference.convertNumToColString(colAccs)
                + ", Remarks: " + CellReference.convertNumToColString(colRemarks));

        GridStyles styles = new GridStyles(wb);

        // Find or create the two new columns at the end of the sheet
        int colShort = -1;
        int colNote = -1;
        int firstFree = -1;
        int limit = Math.max(h.getLastCellNum(), colShortShip + 3) + 5;
        DataFormatter f = new DataFormatter();
        for (int c = colShortShip + 1; c < limit; c++) {
            String t = norm(cellText(f, h.getCell(c)));
            if (t.equals(norm(NEW_COL_SHORT))) {
                colShort = c;
            } else if (t.equals(norm(NEW_COL_NOTE))) {
                colNote = c;
            } else if (t.isEmpty() && firstFree < 0 && c >= colShortShip + 2) {
                firstFree = c;
            }
        }
        if (colShort < 0) {
            colShort = firstFree;
            firstFree = firstFree + 1;
            Cell hc = h.createCell(colShort);
            hc.setCellValue(NEW_COL_SHORT);
            hc.setCellStyle(styles.header);
        }
        if (colNote < 0) {
            colNote = firstFree;
            Cell hc = h.createCell(colNote);
            hc.setCellValue(NEW_COL_NOTE);
            hc.setCellStyle(styles.header);
        }
        sh.setColumnWidth(colShort, 256 * 20);
        sh.setColumnWidth(colNote, 256 * 70);

        // Collect valid data rows (those with a JOB_ERP value)
        List<Integer> idx = new ArrayList<>();
        for (int r = hr + 1; r <= sh.getLastRowNum(); r++) {
            Row row = sh.getRow(r);
            if (row == null) {
                continue;
            }
            Cell jc = row.getCell(colJob);
            if (jc != null && !cellText(f, jc).isEmpty()) {
                idx.add(r);
            }
        }
        int[] arr = new int[idx.size()];
        for (int i = 0; i < arr.length; i++) {
            arr[i] = idx.get(i);
        }

        return new PoiWipGrid(sh, styles, arr, hr,
                colStyle, colColor, colEta, colAccs, colLatest, colBuffer,
                colMatStatus, colJob, colStatus, colRemarks, colShortShip,
                colShort, colNote);
    }

    private static int findCol(Row h, DataFormatter df, String exactMatch, String[] keywords, String[] excludeKeywords, int defaultCol) {
        // Priority 1: exact whole-cell match
        if (exactMatch != null && !exactMatch.isEmpty()) {
            for (int c = 0; c < h.getLastCellNum(); c++) {
                Cell cell = h.getCell(c);
                if (cell == null) continue;
                String text = norm(cellText(df, cell));
                if (text.equals(norm(exactMatch))) {
                    return c;
                }
            }
        }
        // Priority 2: contains all keywords and none of the excludeKeywords
        if (keywords != null && keywords.length > 0) {
            for (int c = 0; c < h.getLastCellNum(); c++) {
                Cell cell = h.getCell(c);
                if (cell == null) continue;
                String text = norm(cellText(df, cell));
                boolean match = true;
                for (String kw : keywords) {
                    if (!text.contains(kw)) {
                        match = false;
                        break;
                    }
                }
                if (match && excludeKeywords != null) {
                    for (String ex : excludeKeywords) {
                        if (text.contains(ex)) {
                            match = false;
                            break;
                        }
                    }
                }
                if (match) {
                    return c;
                }
            }
        }
        return defaultCol;
    }

    private Row row(int i) {
        return sheet.getRow(rows[i]);
    }

    @Override
    public int size() {
        return rows.length;
    }

    @Override
    public int excelRow(int i) {
        return rows[i] + 1;
    }

    @Override
    public String job(int i) {
        return text(fmt, row(i), colJob);
    }

    @Override
    public String style(int i) {
        return text(fmt, row(i), colStyle);
    }

    @Override
    public String color(int i) {
        return text(fmt, row(i), colColor);
    }

    @Override
    public String status(int i) {
        return text(fmt, row(i), colStatus);
    }

    @Override
    public LocalDate eta(int i) {
        return readDate(row(i).getCell(colEta));
    }

    @Override
    public String remarks(int i) {
        return text(fmt, row(i), colRemarks);
    }

    @Override
    public String materialStatus(int i) {
        return cellText(fmt, row(i).getCell(colMatStatus));
    }

    @Override
    public void writeEta(int i, LocalDate d) throws BotException {
        Cell c = guard(row(i), colEta);
        c.setCellValue(toDate(d));
        c.setCellStyle(styles.botDate);
    }

    @Override
    public void writeRemarks(int i, String text) throws BotException {
        Cell c = guard(row(i), colRemarks);
        c.setCellValue(text);
        c.setCellStyle(styles.botText);
    }

    @Override
    public void writeShort(int i, double yds) throws BotException {
        Cell c = guard(row(i), colShort);
        c.setCellValue(yds);
        c.setCellStyle(styles.num);
    }

    @Override
    public void writeNote(int i, String text) throws BotException {
        Cell c = guard(row(i), colNote);
        c.setCellValue(text);
        c.setCellStyle(styles.note);
    }

    /**
     * Guardrail mechanism: only allow writing to whitelisted columns
     * and never overwrite cells that contain a formula (FORMULA).
     */
    private Cell guard(Row row, int col) throws BotException {
        if (!writable.contains(col)) {
            throw new BotException("GUARD_VIOLATION", "Cột " + CellReference.convertNumToColString(col)
                    + " không nằm trong whitelist ghi của Bot");
        }
        Cell c = row.getCell(col);
        if (c == null) {
            c = row.createCell(col);
        }
        if (c.getCellType() == CellType.FORMULA) {
            throw new BotException("FORMULA_PROTECTED", "Ô " + new CellReference(c).formatAsString()
                    + " chứa công thức tính toán, Bot không được ghi đè!");
        }
        return c;
    }

    /**
     * Snapshot the protected columns (Accs, Latest, Buffer, MatStatus, ShortShip)
     * to verify integrity before and after the bot runs.
     */
    public Map<String, String> protectedSnapshot() {
        Map<String, String> snap = new HashMap<>();
        int[] cols = {colAccs, colLatest, colBuffer, colMatStatus, colShortShip};
        for (int r = headerRow + 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                continue;
            }
            for (int c : cols) {
                Cell cell = row.getCell(c);
                String v = "";
                if (cell != null) {
                    v = cell.getCellType() == CellType.FORMULA ? "=" + cell.getCellFormula() : fmt.formatCellValue(cell);
                }
                snap.put(r + ":" + c, v);
            }
        }
        return snap;
    }

    public int getHeaderRow() {
        return headerRow;
    }

    public int getColShort() {
        return colShort;
    }

    public int getColNote() {
        return colNote;
    }

    public int getColEta() {
        return colEta;
    }

    public int getColRemarks() {
        return colRemarks;
    }

    // ====================== CELL READ / FORMAT UTILITIES ======================

    public static String cellText(DataFormatter fmt, Cell c) {
        if (c == null) {
            return "";
        }
        if (c.getCellType() == CellType.FORMULA) {
            CellType t = c.getCachedFormulaResultType();
            if (t == CellType.STRING) {
                return c.getStringCellValue().trim();
            }
            if (t == CellType.NUMERIC) {
                double v = c.getNumericCellValue();
                return v == Math.rint(v) ? String.valueOf((long) v) : String.valueOf(v);
            }
            return "";
        }
        return fmt.formatCellValue(c).trim();
    }

    private static String text(DataFormatter fmt, Row row, int col) {
        return cellText(fmt, row.getCell(col));
    }

    public static double number(Cell c) {
        if (c == null) {
            return 0;
        }
        CellType t = c.getCellType();
        if (t == CellType.FORMULA) {
            t = c.getCachedFormulaResultType();
        }
        if (t == CellType.NUMERIC) {
            return c.getNumericCellValue();
        }
        if (t == CellType.STRING) {
            try {
                return Double.parseDouble(c.getStringCellValue().trim().replace(",", ""));
            } catch (NumberFormatException e) {
                return 0;
            }
        }
        return 0;
    }

    public static LocalDate readDate(Cell c) {
        if (c == null) {
            return null;
        }
        CellType t = c.getCellType();
        if (t == CellType.FORMULA) {
            t = c.getCachedFormulaResultType();
        }
        if (t == CellType.NUMERIC) {
            double v = c.getNumericCellValue();
            if (DateUtil.isCellDateFormatted(c) || (v > 20000 && v < 80000)) {
                Date d = c.getDateCellValue();
                return d == null ? null : toLocalDate(d);
            }
            return null;
        }
        if (t == CellType.STRING) {
            String s = c.getStringCellValue().trim();
            if (s.isEmpty()) {
                return null;
            }
            String[] patterns = {"dd/MM/yyyy", "d/M/yyyy", "yyyy-MM-dd"};
            for (String p : patterns) {
                try {
                    return LocalDate.parse(s, DateTimeFormatter.ofPattern(p));
                } catch (DateTimeParseException ignored) {
                    // Try the next date format
                }
            }
        }
        return null;
    }

    public static int col(String letters) {
        return CellReference.convertColStringToIndex(letters);
    }

    public static Date toDate(LocalDate d) {
        return Date.from(d.atStartOfDay(ZoneId.systemDefault()).toInstant());
    }

    public static LocalDate toLocalDate(Date d) {
        return d.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
    }

    public static String norm(String s) {
        return s == null ? "" : s.replace('\n', ' ').replace('\r', ' ').trim().replaceAll("\\s+", " ").toLowerCase(Locale.ROOT);
    }

    // ====================== CELL STYLE MANAGEMENT ======================

    private static class GridStyles {
        final CellStyle header;
        final CellStyle botDate;
        final CellStyle botText;
        final CellStyle num;
        final CellStyle note;

        GridStyles(XSSFWorkbook wb) {
            DataFormat df = wb.createDataFormat();
            short fmtDate = df.getFormat("dd/mm/yyyy");
            short fmtNum = df.getFormat("#,##0.0");
            Font bold = wb.createFont();
            bold.setBold(true);

            header = make(wb, bold, IndexedColors.LIGHT_YELLOW, (short) -1, true, HorizontalAlignment.CENTER);
            botDate = make(wb, null, IndexedColors.LIGHT_YELLOW, fmtDate, false, HorizontalAlignment.CENTER);
            botText = make(wb, null, IndexedColors.LIGHT_YELLOW, (short) -1, false, null);
            num = make(wb, null, null, fmtNum, false, HorizontalAlignment.RIGHT);
            note = make(wb, null, null, (short) -1, false, null);
        }

        private static CellStyle make(XSSFWorkbook wb, Font font, IndexedColors fill, short fmt, boolean wrap,
                                      HorizontalAlignment align) {
            CellStyle s = wb.createCellStyle();
            if (font != null) {
                s.setFont(font);
            }
            if (fill != null) {
                s.setFillForegroundColor(fill.getIndex());
                s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }
            if (fmt >= 0) {
                s.setDataFormat(fmt);
            }
            if (align != null) {
                s.setAlignment(align);
            }
            s.setWrapText(wrap);
            s.setVerticalAlignment(VerticalAlignment.CENTER);
            s.setBorderTop(BorderStyle.THIN);
            s.setBorderBottom(BorderStyle.THIN);
            s.setBorderLeft(BorderStyle.THIN);
            s.setBorderRight(BorderStyle.THIN);
            return s;
        }
    }
}
