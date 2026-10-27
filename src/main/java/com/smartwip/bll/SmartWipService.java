package com.smartwip.bll;

import java.time.DateTimeException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoUnit;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import com.smartwip.bll.model.BotException;
import com.smartwip.bll.model.ChangeRow;
import com.smartwip.bll.model.Group;
import com.smartwip.bll.model.Remark;
import com.smartwip.bll.model.Result;
import com.smartwip.bll.model.SttLine;
import com.smartwip.dal.WipGrid;

/**
 * Core business logic layer (Business Logic Layer) of the Smart-WIP system.
 * Handles all data reconciliation rules, string processing and ETA calculation.
 */
public class SmartWipService {

    // ====================== BUSINESS RULE CONFIGURATION ======================
    /** Sea Freight = +25 transit days */
    public static final int SF_TRANSIT_DAYS = 25;
    /** Air Freight = +5 transit days */
    public static final int AF_TRANSIT_DAYS = 5;
    /** true: Remark 355 (ETD + transit) overrides the D365 Final ETA Date; false: only used when Final ETA is empty */
    public static final boolean REMARK_355_OVERRIDES_FINAL_ETA = true;
    /** true: when the D365 ETA is earlier than the WIP ETA, still overwrite it and attach a warning */
    public static final boolean WRITE_WHEN_EARLIER = true;
    /** Sales Order types out of WIP scope: treated as INFO, not errors */
    public static final Set<String> OUT_OF_SCOPE_TYPES = Collections.unmodifiableSet(new HashSet<>(Arrays.asList("TS", "PA")));

    public static final String SEV_ERROR = "ERROR";
    public static final String SEV_WARN = "WARN";
    public static final String SEV_INFO = "INFO";

    private static final Pattern REMARK_PATTERN = Pattern.compile(
            "^(FU|BL)\\s+(AF|SF)\\s+ETD\\s+(\\d{1,2})/(\\d{1,2})(?:/(\\d{4}))?$", Pattern.CASE_INSENSITIVE);
    public static final DateTimeFormatter DMY = DateTimeFormatter.ofPattern("dd/MM/yyyy");

    // ====================== DATA NORMALIZATION FOR RECONCILIATION ======================

    /**
     * Normalize the Style code (strip a trailing ".0" when Excel auto-formats it as a number).
     */
    public static String normStyle(String raw) {
        String s = raw == null ? "" : raw.trim().toUpperCase(Locale.ROOT);
        if (s.matches("\\d+\\.0+")) {
            s = s.substring(0, s.indexOf('.'));
        }
        return s;
    }

    /**
     * Normalize the Color code:
     * D365 typically uses the form '638666_01' while WIP stores '1' -> both reduce to '1'.
     */
    public static String normColor(String raw) {
        String s = raw == null ? "" : raw.trim();
        int u = s.lastIndexOf('_');
        if (u >= 0) {
            s = s.substring(u + 1);
        }
        s = s.trim().toUpperCase(Locale.ROOT);
        if (s.matches("\\d+(\\.0+)?")) {
            if (s.contains(".")) {
                s = s.substring(0, s.indexOf('.'));
            }
            s = s.replaceFirst("^0+(?!$)", "");
        }
        return s;
    }

    /**
     * Build the composite key: JOB_ERP + STYLE + COLOR to uniquely identify a WIP row.
     */
    public static String key(String job, String style, String color) {
        return (job == null ? "" : job.trim().toUpperCase(Locale.ROOT)) + "|" + normStyle(style) + "|" + normColor(color);
    }

    // ====================== REMARK 355 PARSING ======================

    /**
     * Parse a Purchaser's Remark 355 in the standard format: [FU|BL] [AF|SF] ETD dd/MM[/yyyy].
     */
    public static Remark parseRemark(String raw, LocalDate today) {
        String t = raw == null ? "" : raw.trim().replaceAll("\\s+", " ");
        if (t.isEmpty()) {
            return null;
        }
        Remark r = new Remark();
        Matcher m = REMARK_PATTERN.matcher(t);
        if (!m.matches()) {
            r.setError("Sai cú pháp '" + t + "'. Chuẩn: [FU|BL] [AF|SF] ETD dd/MM[/yyyy]");
            return r;
        }

        int d = Integer.parseInt(m.group(3));
        int mo = Integer.parseInt(m.group(4));
        LocalDate etd;
        try {
            if (m.group(5) != null) {
                etd = LocalDate.of(Integer.parseInt(m.group(5)), mo, d);
            } else {
                etd = LocalDate.of(today.getYear(), mo, d);
                // If no year is given and the date is more than 90 days in the past, infer the next year
                if (etd.isBefore(today.minusDays(90))) {
                    etd = etd.plusYears(1);
                }
            }
        } catch (DateTimeException e) {
            r.setError("Ngày ETD không hợp lệ: '" + t + "'");
            return r;
        }

        String tag = m.group(1).toUpperCase(Locale.ROOT);
        String mode = m.group(2).toUpperCase(Locale.ROOT);
        r.setEtd(etd);
        r.setEta(etd.plusDays("SF".equals(mode) ? SF_TRANSIT_DAYS : AF_TRANSIT_DAYS));
        r.setNormalized(tag + " " + mode + " ETD " + etd.format(DMY));
        return r;
    }

    // ====================== MAIN PROCESSING FLOW ======================

    /**
     * Run the processing and reconciliation flow between the D365 STT list and the WIP sheet.
     */
    public static Result process(List<SttLine> lines, WipGrid wip, LocalDate today, String stamp) {
        Result res = new Result();

        // 1. Index WIP rows by the JOB_ERP + STYLE + COLOR key
        Map<String, List<Integer>> idx = new HashMap<>();
        Set<String> jobs = new HashSet<>();
        for (int i = 0; i < wip.size(); i++) {
            String j = wip.job(i).trim().toUpperCase(Locale.ROOT);
            if (j.isEmpty()) {
                continue;
            }
            jobs.add(j);
            idx.computeIfAbsent(key(j, wip.style(i), wip.color(i)), k -> new ArrayList<>()).add(i);
        }

        // 2. Group STT lines: many PO/BOM/Size rows collapse into one group matching one WIP row
        Map<String, Group> groups = new LinkedHashMap<>();
        for (SttLine l : lines) {
            Group g = groups.computeIfAbsent(key(l.getJob(), l.getStyle(), l.getColor()), k -> new Group(l));
            addLine(g, l, today, res);
        }

        // 3. Perform reconciliation and apply updates
        for (Map.Entry<String, Group> e : groups.entrySet()) {
            applyGroup(e.getValue(), idx.get(e.getKey()), jobs, wip, res, stamp);
        }

        return res;
    }


    private static void addLine(Group g, SttLine l, LocalDate today, Result res) {
        g.setLines(g.getLines() + 1);
        if (l.getPo() != null && !l.getPo().trim().isEmpty()) {
            g.getPos().add(l.getPo().trim());
        }
        if ("OPEN ORDER".equalsIgnoreCase(l.getPoStatus() == null ? "" : l.getPoStatus().trim())) {
            g.setOpenPo(true);
        }

        double shortLine = Math.max(0, l.getReq() - l.getRes());
        if ("YDS".equalsIgnoreCase(l.getUnit() == null ? "" : l.getUnit().trim())) {
            g.setShortYds(g.getShortYds() + shortLine);
        } else {
            g.setShortOther(g.getShortOther() + shortLine);
        }

        LocalDate lineEta = l.getFinalEta();
        Remark r = parseRemark(l.getRemark355(), today);
        if (r != null) {
            if (r.getError() != null) {
                // Syntax error: log the error; this line falls back to the D365 Final ETA
                res.log(SEV_ERROR, "SYNTAX_ERROR", g, 0, "PO " + l.getPo() + ": " + r.getError());
            } else {
                g.getRemarks().add(r.getNormalized());
                g.setUsed355(true);
                lineEta = (REMARK_355_OVERRIDES_FINAL_ETA || l.getFinalEta() == null) ? r.getEta() : l.getFinalEta();
            }
        }

        if (lineEta == null) {
            g.setNoEta(g.getNoEta() + 1);
        } else if (g.getEta() == null || lineEta.isAfter(g.getEta())) {
            g.setEta(lineEta);
        }
    }

    private static void applyGroup(Group g, List<Integer> hits, Set<String> jobs, WipGrid wip, Result res, String stamp) {
        res.setGroups(res.getGroups() + 1);
        res.setTotalShortYds(res.getTotalShortYds() + g.getShortYds());

        if (hits == null) {
            String jobU = g.getJob().trim().toUpperCase(Locale.ROOT);
            if (jobs.contains(jobU)) {
                res.log(SEV_ERROR, "KEY_NOT_FOUND", g, 0,
                        "JOB có trong WIP nhưng STYLE/COLOR không khớp [" + g.getJob() + " | " + g.getStyle() + " | " + g.getColor() + "]");
            } else if (OUT_OF_SCOPE_TYPES.contains(g.getSoType().trim().toUpperCase(Locale.ROOT))) {
                res.log(SEV_INFO, "OUT_OF_SCOPE", g, 0,
                        "SO loại " + g.getSoType() + " không có trong WIP (ngoài phạm vi theo dõi), bỏ qua");
            } else {
                res.log(SEV_ERROR, "KEY_NOT_FOUND", g, 0,
                        "Không tìm thấy JOB trong WIP [" + g.getJob() + " | " + g.getStyle() + " | " + g.getColor() + "]");
            }
            return;
        }

        res.inc("groups.matched");
        if (g.getEta() == null) {
            res.log(SEV_WARN, "NO_ETA", g, 0, "D365 chưa có ETA cho " + g.getLines() + " dòng STT, giữ nguyên AG");
        } else if (g.getNoEta() > 0) {
            res.log(SEV_WARN, "PARTIAL_ETA", g, 0,
                    g.getNoEta() + "/" + g.getLines() + " dòng D365 chưa có ETA, ETA tính theo các dòng có sẵn");
        }
        if (hits.size() > 1) {
            res.log(SEV_INFO, "DUP_KEY", g, 0, "Khóa trùng khớp với " + hits.size() + " dòng trên WIP, cập nhật toàn bộ");
        }

        for (int i : hits) {
            int xr = wip.excelRow(i);
            if (!"OPEN".equalsIgnoreCase(wip.status(i).trim())) {
                res.inc("rows.skipped_not_open");
                res.log(SEV_INFO, "ROW_NOT_OPEN", g, xr, "Dòng WIP ở trạng thái '" + wip.status(i) + "', không cập nhật");
                continue;
            }
            res.inc("rows.matched");

            LocalDate cur = wip.eta(i);
            String action;
            long diff = 0;
            if (g.getEta() == null) {
                action = "NO_ETA";
            } else if (cur == null) {
                action = "FILL";
            } else {
                diff = ChronoUnit.DAYS.between(cur, g.getEta());
                action = diff == 0 ? "NO_CHANGE" : (diff > 0 ? "LATER" : "EARLIER");
            }
            res.inc("action." + action);

            List<String> notes = new ArrayList<>();
            boolean etaWritten = false;
            boolean remarkWritten = false;

            try {
                if ("FILL".equals(action) || "LATER".equals(action)
                        || ("EARLIER".equals(action) && WRITE_WHEN_EARLIER)) {
                    wip.writeEta(i, g.getEta());
                    etaWritten = true;
                }
                if ("FILL".equals(action)) {
                    notes.add("WIP chưa có ETA, đã điền mới");
                } else if ("LATER".equals(action)) {
                    notes.add("ETA mới (" + (g.isUsed355() ? "Remark 355" : "D365") + ") trễ hơn WIP " + diff + " ngày");
                } else if ("EARLIER".equals(action)) {
                    notes.add("ETA mới (" + (g.isUsed355() ? "Remark 355" : "D365") + ") sớm hơn WIP " + (-diff) + " ngày"
                            + (WRITE_WHEN_EARLIER ? "" : " (chưa ghi đè)"));
                } else if ("NO_ETA".equals(action)) {
                    notes.add("D365 chưa có thông tin ETA");
                }

                if (g.getEta() != null && g.getNoEta() > 0) {
                    notes.add(g.getNoEta() + "/" + g.getLines() + " dòng D365 chưa có ETA");
                }
                if (g.isUsed355()) {
                    notes.add("ETA theo Remark 355: " + String.join("; ", g.getRemarks()));
                }
                if (hits.size() > 1) {
                    notes.add("khóa trùng " + hits.size() + " dòng WIP");
                }
                if ("RCVD MATERIAL".equalsIgnoreCase(wip.materialStatus(i).trim()) && g.isOpenPo()) {
                    notes.add("WIP đang 'Rcvd Material' nhưng D365 vẫn còn PO mở");
                    res.inc("warn.rcvd_but_d365_open");
                    res.log(SEV_WARN, "RCVD_BUT_D365_OPEN", g, xr, "WIP='Rcvd Material' nhưng D365 còn PO 'Open order'");
                }

                // Rule R5: Remarks are append-only; never overwrite manual comments and never append duplicates
                if (!g.getRemarks().isEmpty()) {
                    String add = String.join("; ", g.getRemarks());
                    String existing = wip.remarks(i).trim();
                    String nt = null;
                    if (existing.isEmpty()) {
                        nt = add;
                    } else if (!existing.toUpperCase(Locale.ROOT).contains(add.toUpperCase(Locale.ROOT))) {
                        nt = existing + " | " + add;
                    }
                    if (nt != null) {
                        wip.writeRemarks(i, nt);
                        remarkWritten = true;
                        res.inc("remarks.appended");
                    }
                }

                wip.writeShort(i, g.getShortYds());
                if (!notes.isEmpty()) {
                    wip.writeNote(i, "[" + stamp + "] " + String.join(" | ", notes));
                }
            } catch (BotException ex) {
                res.log(SEV_ERROR, ex.getCode(), g, xr, ex.getMessage());
                notes.add("LỖI " + ex.getCode());
            }

            if (etaWritten) {
                res.inc("eta.written");
            }

            if (etaWritten || remarkWritten || !notes.isEmpty()) {
                ChangeRow c = new ChangeRow();
                c.setExcelRow(xr);
                c.setGroup(g);
                c.setAgOld(cur);
                c.setAgNew(etaWritten ? g.getEta() : null);
                c.setDiff(diff);
                c.setAction(action);
                c.setSource(g.isUsed355() ? "REMARK_355" : "D365_FINAL_ETA");
                c.setShortYds(g.getShortYds());
                c.setNote(String.join(" | ", notes));
                res.getChanges().add(c);
            }
        }
    }
}
