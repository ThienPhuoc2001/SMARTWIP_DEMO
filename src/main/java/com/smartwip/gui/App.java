package com.smartwip.gui;

import java.awt.Desktop;
import java.awt.GraphicsEnvironment;
import java.io.File;
import java.io.FileDescriptor;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.FileSystemException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.smartwip.bll.SmartWipService;
import com.smartwip.bll.model.ChangeRow;
import com.smartwip.bll.model.LogEntry;
import com.smartwip.bll.model.Result;
import com.smartwip.bll.model.SttLine;
import com.smartwip.dal.FileManager;
import com.smartwip.dal.PoiWipGrid;

/**
 * Command-line presentation and orchestration layer (CLI Presentation Layer) for Smart-WIP v2.
 */
public final class App {

    public static final Path OUT_WIP = Paths.get("Updated_WIP.xlsx");
    public static final Path OUT_EXC = Paths.get("Exception_Log.csv");
    public static final Path OUT_CHG = Paths.get("Bot_Changes.csv");

    private static final DateTimeFormatter STAMP = DateTimeFormatter.ofPattern("dd/MM HH:mm");

    private App() {
    }

    public static void main(String[] args) {
        System.setProperty("java.awt.headless", "false");
        System.setProperty("log4j2.loggerContextFactory", "org.apache.logging.log4j.simple.SimpleLoggerContextFactory");
        System.setProperty("log4j2.StatusLogger.level", "OFF");
        System.setProperty("org.apache.logging.log4j.simplelog.StatusLogger.level", "OFF");
        System.setOut(new PrintStream(new FileOutputStream(FileDescriptor.out), true, StandardCharsets.UTF_8));
        IOUtils.setByteArrayMaxOverride(1_000_000_000);

        String sttArg = null;
        String wipArg = null;
        boolean auto = false;
        boolean noOpen = false;
        Map<String, String> remarkInputs = new LinkedHashMap<>();

        for (String a : args) {
            if ("--auto".equals(a)) {
                auto = true;
            } else if ("--no-open".equals(a)) {
                noOpen = true;
            } else if (a.startsWith("--stt=")) {
                sttArg = a.substring(6);
            } else if (a.startsWith("--wip=")) {
                wipArg = a.substring(6);
            } else if (a.startsWith("--remark=")) {
                String v = a.substring(9);
                int c = v.indexOf(':');
                if (c > 0) {
                    remarkInputs.put(v.substring(0, c).trim().toUpperCase(Locale.ROOT),
                            v.substring(c + 1).replace('_', ' ').trim());
                }
            }
        }

        int exit = 0;
        try {
            run(sttArg, wipArg, auto, noOpen, remarkInputs);
        } catch (FileSystemException e) {
            say("");
            say("[LỖI] Không ghi/đọc được file: '" + e.getFile() + "'. Nếu file đang mở trong Excel, hãy ĐÓNG lại rồi chạy lại.");
            exit = 1;
        } catch (OutOfMemoryError e) {
            say("");
            say("[LỖI] Tràn bộ nhớ (OutOfMemoryError) do file Excel lớn. Vui lòng tăng bộ nhớ JVM:");
            say("   Windows  : set MAVEN_OPTS=-Xmx4g");
            say("   Mac/Linux: export MAVEN_OPTS=-Xmx4g");
            say("   Chạy JAR : java -Xmx4g -jar target/smart-wip-demo.jar");
            exit = 1;
        } catch (Exception e) {
            say("");
            say("[LỖI KHÔNG MONG MUỐN] " + e.getMessage());
            e.printStackTrace();
            exit = 1;
        }

        sleepMs(800);
        System.exit(exit);
    }

    private static void run(String sttArg, String wipArg, boolean auto, boolean noOpen,
                            Map<String, String> remarkInputs) throws Exception {
        LocalDate today = LocalDate.now();
        String stamp = LocalDateTime.now().format(STAMP);

        say("================================================================");
        say("  SMART-WIP v2  |  D365 (STT BULK)  ->  Bot RPA  ->  Excel WIP");
        say("================================================================");

        Path sttFile = FileManager.findFile(sttArg, true);
        Path wipFile = FileManager.findFile(wipArg, false);

        title("BƯỚC 1: ĐỌC DỮ LIỆU D365 VÀ FILE WIP GỐC");
        say("  D365 STT : " + sttFile.toAbsolutePath());
        say("  WIP File : " + wipFile.toAbsolutePath());
        say("  Đang đọc dữ liệu D365 STT ... (vui lòng chờ)");
        List<SttLine> lines = FileManager.readStt(sttFile);

        Set<String> sos = new HashSet<>();
        Set<String> pos = new HashSet<>();
        int with355 = 0;
        for (SttLine l : lines) {
            sos.add(l.getJob());
            if (!l.getPo().isEmpty()) {
                pos.add(l.getPo());
            }
            if (!l.getRemark355().isEmpty()) {
                with355++;
            }
        }
        say("  -> " + lines.size() + " dòng STT | " + sos.size() + " Sales Order | " + pos.size()
                + " PO | Dòng đã có Remark 355: " + with355);

        // Simulate Purchasing entering Remark 355 when the --remark=PO:TEXT argument is given
        for (Map.Entry<String, String> e : remarkInputs.entrySet()) {
            int n = 0;
            for (SttLine l : lines) {
                if (l.getPo().trim().toUpperCase(Locale.ROOT).equals(e.getKey())) {
                    l.setRemark355(e.getValue());
                    n++;
                }
            }
            say("  [MÔ PHỎNG D365] PO " + e.getKey() + " > Remark 355 := '" + e.getValue()
                    + "' -> Lưu thành công cho " + n + " dòng STT.");
        }

        say("  Đang nạp file WIP gốc (chế độ chỉ đọc) ...");
        try (InputStream in = Files.newInputStream(wipFile);
             XSSFWorkbook wb = new XSSFWorkbook(in)) {
            PoiWipGrid grid = PoiWipGrid.open(wb);
            say("  -> Sheet 'WIP': " + grid.size() + " dòng dữ liệu (Dòng tiêu đề ở dòng " + (grid.getHeaderRow() + 1)
                    + "). Cột mới: " + CellReference.convertNumToColString(grid.getColShort()) + " và "
                    + CellReference.convertNumToColString(grid.getColNote()));

            if (with355 == 0 && remarkInputs.isEmpty()) {
                say("  Lưu ý: Remark 355 trống -> Bot sẽ tự động lấy Final ETA Date của D365 làm nguồn.");
            }

            waitForRun(auto);

            title("BƯỚC 2: BOT RPA XỬ LÝ ĐỐI CHIẾU (PROCESSING)");
            printRules();
            Map<String, String> before = grid.protectedSnapshot();

            Result res = SmartWipService.process(lines, grid, today, stamp);

            say("");
            say(before.equals(grid.protectedSnapshot())
                    ? "Guardrail Check: AH, AI, AJ, BN (công thức/nhập tay) và CI (Short Ship) KHÔNG BỊ THAY ĐỔI -> AN TOÀN [PASSED]"
                    : "Guardrail Check: [CẢNH BÁO NGUY HIỂM] Có ô được bảo vệ bị thay đổi dữ liệu!");

            title("BƯỚC 3: XUẤT FILE KẾT QUẢ VÀ BÁO CÁO");
            wb.setForceFormulaRecalculation(true);
            say("  Đang ghi file kết quả: " + OUT_WIP.toAbsolutePath() + " ...");
            try (OutputStream out = Files.newOutputStream(OUT_WIP)) {
                wb.write(out);
            }

            FileManager.writeExceptionLog(res, OUT_EXC);
            FileManager.writeChanges(res, OUT_CHG);

            printSummary(res);
        }

        if (noOpen) {
            say("Bỏ qua tự động mở file (--no-open). File kết quả: " + OUT_WIP.toAbsolutePath());
        } else {
            say("");
            say("Đang tự động mở Updated_WIP.xlsx ...");
            openFile(OUT_WIP.toFile());
        }
    }

    private static void printRules() {
        say("Các Business Rule áp dụng:");
        say("  R1: Remark 355 = [FU|BL] [AF|SF] ETD dd/MM[/yyyy]; sai cú pháp -> SYNTAX_ERROR, fallback dùng Final ETA");
        say("  R2: ETA = ETD + Transit (Sea Freight +" + SmartWipService.SF_TRANSIT_DAYS + " ngày; Air Freight +"
                + SmartWipService.AF_TRANSIT_DAYS + " ngày)");
        say("  R3: Khóa ghép JOB_ERP + STYLE + COLOR; chỉ cập nhật các dòng WIP trạng thái 'Open'");
        say("  R4: AG = MAX(ETA của mọi dòng PO thuộc SO); D365 không có ETA -> giữ nguyên AG, báo NO_ETA");
        say("  R5: Cột BZ (Remarks) chỉ nối thêm Remark 355 (không ghi đè comment tay, không nối lặp)");
        say("  R6: Material_Short_Qty = SUM(Require - Reserve) các dòng YDS ghi cột mới; không sửa cột CI");
        say("  R7: Chỉ ghi AG, BZ và 2 cột mới; tuyệt đối không ghi đè ô công thức (CellType.FORMULA)");
        say("  R8: Lỗi/cảnh báo xuất ra Exception_Log.csv; SO loại TS/PA ghi nhận INFO");
    }

    private static void printSummary(Result res) {
        say("");
        say("Số liệu đối chiếu (dòng WIP):");
        say("  Khớp khóa và cập nhật  : " + res.stat("rows.matched") + " dòng (" + res.stat("groups.matched") + " Sales Order)");
        say("  AG đã đúng (NO_CHANGE) : " + res.stat("action.NO_CHANGE"));
        say("  D365 trễ hơn (LATER)   : " + res.stat("action.LATER"));
        say("  D365 sớm hơn (EARLIER) : " + res.stat("action.EARLIER"));
        say("  WIP trống (FILL)       : " + res.stat("action.FILL"));
        say("  D365 chưa có ETA       : " + res.stat("action.NO_ETA"));
        say("  WIP 'Rcvd' nhưng D365 còn PO mở: " + res.stat("warn.rcvd_but_d365_open"));
        say("  Remarks (BZ) đã nối    : " + res.stat("remarks.appended"));
        say("  Tổng thiếu YDS (theo SO): " + String.format(Locale.ROOT, "%,.1f", res.getTotalShortYds()));
        say("  SO ngoài phạm vi (TS/PA): " + res.stat("log.OUT_OF_SCOPE"));

        int errors = 0;
        int warns = 0;
        for (LogEntry e : res.getLog()) {
            if (SmartWipService.SEV_ERROR.equals(e.getSeverity())) {
                errors++;
            } else if (SmartWipService.SEV_WARN.equals(e.getSeverity())) {
                warns++;
            }
        }

        say("");
        say("Đã xuất : " + OUT_WIP.toAbsolutePath());
        say("Đã ghi  : " + OUT_EXC.toAbsolutePath() + " (" + res.getLog().size() + " dòng: " + errors + " lỗi, " + warns + " cảnh báo)");
        say("Đã ghi  : " + OUT_CHG.toAbsolutePath() + " (" + res.getChanges().size() + " dòng có cập nhật/cảnh báo)");

        List<ChangeRow> later = new ArrayList<>();
        for (ChangeRow c : res.getChanges()) {
            if ("LATER".equals(c.getAction())) {
                later.add(c);
            }
        }
        later.sort(Comparator.comparingLong(ChangeRow::getDiff).reversed());
        if (!later.isEmpty()) {
            say("");
            say("Top dòng WIP hàng về TRỄ hơn so với D365 nhiều nhất:");
            for (int i = 0; i < Math.min(5, later.size()); i++) {
                ChangeRow c = later.get(i);
                say("  Dòng " + c.getExcelRow() + " | " + c.getGroup().getJob() + " | Fabric ETA "
                        + FileManager.fmtDate(c.getAgOld()) + " -> " + FileManager.fmtDate(c.getAgNew())
                        + " (+" + c.getDiff() + " ngày) | PO: " + c.getGroup().poSample());
            }
        }

        if (errors + warns > 0) {
            say("");
            say("[GIẢ LẬP EMAIL CẢNH BÁO] To: Merchandiser / Purchaser");
            say("Subject: SMART-WIP - " + errors + " lỗi, " + warns + " cảnh báo cần xử lý");
            for (Map.Entry<String, Integer> e : res.getStats().entrySet()) {
                if (e.getKey().startsWith("log.") && !"log.OUT_OF_SCOPE".equals(e.getKey())
                        && !"log.DUP_KEY".equals(e.getKey()) && !"log.ROW_NOT_OPEN".equals(e.getKey())) {
                    say("   - " + e.getKey().substring(4) + ": " + e.getValue());
                }
            }
        }
    }

    private static void openFile(File f) {
        File abs = f.getAbsoluteFile();
        try {
            if (GraphicsEnvironment.isHeadless() || !Desktop.isDesktopSupported()
                    || !Desktop.getDesktop().isSupported(Desktop.Action.OPEN)) {
                say("Môi trường không hỗ trợ mở cửa sổ tự động. Vui lòng mở thủ công: " + abs);
                return;
            }
            Desktop.getDesktop().open(abs);
        } catch (Exception e) {
            say("Không thể tự động mở file (" + e.getMessage() + "). Vui lòng mở thủ công: " + abs);
        }
    }

    private static void waitForRun(boolean auto) {
        if (auto) {
            return;
        }
        System.out.print("\n>>> Nhấn ENTER để bấm RUN (chạy Bot RPA)... ");
        System.out.flush();
        try {
            System.in.read();
        } catch (IOException ignored) {
        }
        say("");
    }

    private static void sleepMs(int ms) {
        if (ms <= 0) {
            return;
        }
        try {
            Thread.sleep(ms);
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
        }
    }

    private static void say(String s) {
        System.out.println(s);
    }

    private static void title(String s) {
        say("");
        say("================================================================");
        say(" " + s);
        say("================================================================");
    }
}
