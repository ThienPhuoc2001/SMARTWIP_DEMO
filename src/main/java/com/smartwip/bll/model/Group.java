package com.smartwip.bll.model;

import java.time.LocalDate;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;

/**
 * Groups STT lines by the composite key JOB_ERP + STYLE + COLOR (equivalent to one WIP row).
 */
public class Group {
    private final String job;
    private final String style;
    private final String color;
    private final String soType;
    private final Set<String> pos = new LinkedHashSet<>();
    private final Set<String> remarks = new LinkedHashSet<>();
    private LocalDate eta;
    private int lines;
    private int noEta;
    private double shortYds;
    private double shortOther;
    private boolean openPo;
    private boolean used355;

    public Group(SttLine line) {
        this.job = line.getJob();
        this.style = line.getStyle();
        this.color = line.getColor();
        this.soType = line.getSoType();
    }

    public Group(String job, String style, String color, String soType) {
        this.job = job;
        this.style = style;
        this.color = color;
        this.soType = soType;
    }

    public String poSample() {
        List<String> list = new ArrayList<>(pos);
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < Math.min(3, list.size()); i++) {
            if (i > 0) {
                sb.append(", ");
            }
            sb.append(list.get(i));
        }
        if (list.size() > 3) {
            sb.append(" (+").append(list.size() - 3).append(")");
        }
        return sb.toString();
    }

    public String getJob() {
        return job;
    }

    public String getStyle() {
        return style;
    }

    public String getColor() {
        return color;
    }

    public String getSoType() {
        return soType;
    }

    public Set<String> getPos() {
        return pos;
    }

    public Set<String> getRemarks() {
        return remarks;
    }

    public LocalDate getEta() {
        return eta;
    }

    public void setEta(LocalDate eta) {
        this.eta = eta;
    }

    public int getLines() {
        return lines;
    }

    public void setLines(int lines) {
        this.lines = lines;
    }

    public int getNoEta() {
        return noEta;
    }

    public void setNoEta(int noEta) {
        this.noEta = noEta;
    }

    public double getShortYds() {
        return shortYds;
    }

    public void setShortYds(double shortYds) {
        this.shortYds = shortYds;
    }

    public double getShortOther() {
        return shortOther;
    }

    public void setShortOther(double shortOther) {
        this.shortOther = shortOther;
    }

    public boolean isOpenPo() {
        return openPo;
    }

    public void setOpenPo(boolean openPo) {
        this.openPo = openPo;
    }

    public boolean isUsed355() {
        return used355;
    }

    public void setUsed355(boolean used355) {
        this.used355 = used355;
    }
}
