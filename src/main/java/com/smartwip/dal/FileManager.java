package com.smartwip.dal;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.smartwip.bll.model.ChangeRow;
import com.smartwip.bll.model.LogEntry;
import com.smartwip.bll.model.Result;
import com.smartwip.bll.model.SttLine;

/**
 * Handles file I/O: auto-detects files, loads the D365 STT file, and writes CSV reports.
 */
public class FileManager {

    public static final DateTimeFormatter DMY = DateTimeFormatter.ofPattern("dd/MM/yyyy");
    public static final DateTimeFormatter LOGTS = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");

    private FileManager() {
    }

    /**
     * Auto-detect the STT or WIP file in the current working directory when no path is given.
     */
    public static Path findFile(String explicit, boolean isStt) throws IOException {
        if (explicit != null && !explicit.trim().isEmpty()) {
            Path p = Paths.get(explicit.trim());
            if (!Files.exists(p)) {
                throw new IOException("Không tìm thấy file: " + p.toAbsolutePath());
            }
            return p;
        }

        File[] files = new File(".").listFiles();
        File best = null;
        List<String> seen = new ArrayList<>();
        if (files != null) {
            for (File f : files) {
                String n = f.getName().toLowerCase(Locale.ROOT);
                if (!f.isFile() || !n.endsWith(".xlsx") || n.startsWith("~$")) {
                    continue;
                }
                seen.add(f.getName());
                boolean match = isStt
                        ? n.startsWith("stt")
                        : (n.contains("wip") && !n.startsWith("updated_") && !n.startsWith("sample_") && !n.startsWith("wip_assign"));
                if (match && (best == null || f.lastModified() > best.lastModified())) {
                    best = f;
                }
            }
        }
        if (best == null) {
            throw new IOException("Không tự động tìm thấy file " + (isStt ? "STT (tên bắt đầu bằng 'STT')" : "WIP (tên chứa 'WIP')")
                    + " trong thư mục hiện tại. Các file .xlsx đang có: " + seen
                    + ". Vui lòng copy file vào thư mục hoặc truyền tham số: --" + (isStt ? "stt" : "wip") + "=<đường_dẫn>");
        }
        return best.toPath();
    }

    /**
     * Read and parse the D365 ERP STT BULK file.
     */
    public static List<SttLine> readStt(Path file) throws IOException {
        IOUtils.setByteArrayMaxOverride(1_000_000_000);
        List<SttLine> list = new ArrayList<>();
        try (OPCPackage pkg = OPCPackage.open(file.toFile(), PackageAccess.READ);
             XSSFWorkbook wb = new XSSFWorkbook(pkg)) {
            Sheet sh = wb.getSheetAt(0);
            Row h = sh.getRow(0);
            if (h == null) {
                throw new IllegalStateException("File STT không có dòng tiêu đề ở dòng 1.");
            }

            Map<String, Integer> hm = new LinkedHashMap<>();
            for (int c = 0; c < h.getLastCellNum(); c++) {
                Cell hc = h.getCell(c);
                if (hc != null) {
                    hm.put(norm(hc.getStringCellValue()), c);
                }
            }

            int cJob = need(hm, "Sales order");
            int cStyle = need(hm, "Item number");
            int cColor = need(hm, "Garment Color");
            int cType = need(hm, "Sales order type");
            int cPo = need(hm, "Invent BatchId / PO");
            int cPoSt = need(hm, "Purchase order status");
            int cUnit = need(hm, "Unit Of RM");
            int cEta = need(hm, "Final ETA Date");
            int cReq = need(hm, "Require Qty");
            int cRes = need(hm, "Reserve Qty");
            int c355 = opt(hm, "Remark Status(355)");

            DataFormatter fmt = new DataFormatter();
            for (int r = 1; r <= sh.getLastRowNum(); r++) {
                Row row = sh.getRow(r);
                if (row == null) {
                    continue;
                }
                String job = text(fmt, row, cJob);
                if (job.isEmpty()) {
                    continue;
                }

                SttLine l = new SttLine();
                l.setJob(job);
                l.setStyle(text(fmt, row, cStyle));
                l.setColor(text(fmt, row, cColor));
                l.setSoType(text(fmt, row, cType));
                l.setPo(text(fmt, row, cPo));
                l.setPoStatus(text(fmt, row, cPoSt));
                l.setUnit(text(fmt, row, cUnit));
                l.setFinalEta(PoiWipGrid.readDate(row.getCell(cEta)));
                l.setReq(PoiWipGrid.number(row.getCell(cReq)));
                l.setRes(PoiWipGrid.number(row.getCell(cRes)));
                l.setRemark355(c355 >= 0 ? text(fmt, row, c355) : "");
                list.add(l);
            }
        } catch (InvalidFormatException e) {
            throw new IOException("Định dạng file STT không hợp lệ: " + file.toAbsolutePath(), e);
        }
        return list;
    }

    /**
     * Write the Exception_Log.csv exception log as UTF-8 with BOM.
     */
    public static void writeExceptionLog(Result res, Path path) throws IOException {
        StringBuilder sb = new StringBuilder("\uFEFF");
        sb.append("Timestamp,Severity,Code,JOB_ERP,STYLE,COLOR,SO_type,WIP_Excel_Row,PO_sample,Message\r\n");
        String ts = LocalDateTime.now().format(LOGTS);
        for (LogEntry e : res.getLog()) {
            sb.append(csv(ts)).append(',')
                    .append(csv(e.getSeverity())).append(',')
                    .append(csv(e.getCode())).append(',')
                    .append(csv(e.getGroup().getJob())).append(',')
                    .append(csv(e.getGroup().getStyle())).append(',')
                    .append(csv(e.getGroup().getColor())).append(',')
                    .append(csv(e.getGroup().getSoType())).append(',')
                    .append(e.getExcelRow() > 0 ? String.valueOf(e.getExcelRow()) : "").append(',')
                    .append(csv(e.getGroup().poSample())).append(',')
                    .append(csv(e.getMessage())).append("\r\n");
        }
        Files.write(path, sb.toString().getBytes(StandardCharsets.UTF_8));
    }

    /**
     * Write the Bot_Changes.csv change history as UTF-8 with BOM.
     */
    public static void writeChanges(Result res, Path path) throws IOException {
        StringBuilder sb = new StringBuilder("\uFEFF");
        sb.append("WIP_Excel_Row,JOB_ERP,STYLE,COLOR,ETA_old,ETA_new,Diff_days,Action,ETA_source,PO_sample,Short_YDS,Note\r\n");
        for (ChangeRow c : res.getChanges()) {
            sb.append(c.getExcelRow()).append(',')
                    .append(csv(c.getGroup().getJob())).append(',')
                    .append(csv(c.getGroup().getStyle())).append(',')
                    .append(csv(c.getGroup().getColor())).append(',')
                    .append(fmtDate(c.getAgOld())).append(',')
                    .append(fmtDate(c.getAgNew())).append(',')
                    .append(c.getDiff()).append(',')
                    .append(csv(c.getAction())).append(',')
                    .append(csv(c.getSource())).append(',')
                    .append(csv(c.getGroup().poSample())).append(',')
                    .append(String.format(Locale.ROOT, "%.1f", c.getShortYds())).append(',')
                    .append(csv(c.getNote())).append("\r\n");
        }
        Files.write(path, sb.toString().getBytes(StandardCharsets.UTF_8));
    }

    // ====================== INTERNAL UTILITY METHODS ======================

    private static int need(Map<String, Integer> hm, String name) {
        Integer i = hm.get(norm(name));
        if (i == null) {
            throw new IllegalStateException("File STT thiếu cột bắt buộc: '" + name + "'. Các cột tìm thấy: " + hm.keySet());
        }
        return i;
    }

    private static int opt(Map<String, Integer> hm, String name) {
        Integer i = hm.get(norm(name));
        return i == null ? -1 : i;
    }

    private static String text(DataFormatter fmt, Row row, int col) {
        return PoiWipGrid.cellText(fmt, row.getCell(col));
    }

    private static String norm(String s) {
        return PoiWipGrid.norm(s);
    }

    public static String fmtDate(LocalDate d) {
        return d == null ? "" : d.format(DMY);
    }

    public static String csv(String v) {
        if (v == null) {
            return "";
        }
        String s = v.replace("\"", "\"\"");
        boolean quote = s.contains(",") || s.contains("\"") || s.contains("\n") || s.contains("\r");
        return quote ? "\"" + s + "\"" : s;
    }
}
