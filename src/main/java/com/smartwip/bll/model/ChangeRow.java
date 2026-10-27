package com.smartwip.bll.model;

import java.time.LocalDate;

/**
 * Record of the changes made to a single row of the WIP Excel file.
 */
public class ChangeRow {
    private int excelRow;
    private Group group;
    private LocalDate agOld;
    private LocalDate agNew;
    private long diff;
    private String action;
    private String source;
    private double shortYds;
    private String note;

    public ChangeRow() {
    }

    public int getExcelRow() {
        return excelRow;
    }

    public void setExcelRow(int excelRow) {
        this.excelRow = excelRow;
    }

    public Group getGroup() {
        return group;
    }

    public void setGroup(Group group) {
        this.group = group;
    }

    public LocalDate getAgOld() {
        return agOld;
    }

    public void setAgOld(LocalDate agOld) {
        this.agOld = agOld;
    }

    public LocalDate getAgNew() {
        return agNew;
    }

    public void setAgNew(LocalDate agNew) {
        this.agNew = agNew;
    }

    public long diff() {return diff;}

    public long getDiff() {
        return diff;
    }

    public void setDiff(long diff) {
        this.diff = diff;
    }

    public String getAction() {
        return action;
    }

    public void setAction(String action) {
        this.action = action;
    }

    public String getSource() {
        return source;
    }

    public void setSource(String source) {
        this.source = source;
    }

    public double getShortYds() {
        return shortYds;
    }

    public void setShortYds(double shortYds) {
        this.shortYds = shortYds;
    }

    public String getNote() {
        return note;
    }

    public void setNote(String note) {
        this.note = note;
    }
}
