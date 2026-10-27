package com.smartwip.bll.model;

/**
 * A log entry for an exception or warning raised during reconciliation.
 */
public class LogEntry {
    private final String severity;
    private final String code;
    private final Group group;
    private final int excelRow;
    private final String message;

    public LogEntry(String severity, String code, Group group, int excelRow, String message) {
        this.severity = severity;
        this.code = code;
        this.group = group;
        this.excelRow = excelRow;
        this.message = message;
    }

    public String getSeverity() {
        return severity;
    }

    public String getCode() {
        return code;
    }

    public Group getGroup() {
        return group;
    }

    public int getExcelRow() {
        return excelRow;
    }

    public String getMessage() {
        return message;
    }
}
