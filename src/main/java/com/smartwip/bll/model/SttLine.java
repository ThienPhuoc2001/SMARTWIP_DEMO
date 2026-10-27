package com.smartwip.bll.model;

import java.time.LocalDate;

/**
 * One row extracted from the D365 STT BULK file:
 * corresponds to one BOM x Size x Reservation x PO line.
 */
public class SttLine {
    private String job = "";
    private String style = "";
    private String color = "";
    private String soType = "";
    private String po = "";
    private String poStatus = "";
    private String unit = "";
    private String remark355 = "";
    private LocalDate finalEta;
    private double req;
    private double res;



    public SttLine() {
    }

    public String getJob() {
        return job;
    }

    public void setJob(String job) {
        this.job = job == null ? "" : job;
    }

    public String getStyle() {
        return style;
    }

    public void setStyle(String style) {
        this.style = style == null ? "" : style;
    }

    public String getColor() {
        return color;
    }

    public void setColor(String color) {
        this.color = color == null ? "" : color;
    }

    public String getSoType() {
        return soType;
    }

    public void setSoType(String soType) {
        this.soType = soType == null ? "" : soType;
    }

    public String getPo() {
        return po;
    }

    public void setPo(String po) {
        this.po = po == null ? "" : po;
    }

    public String getPoStatus() {
        return poStatus;
    }

    public void setPoStatus(String poStatus) {
        this.poStatus = poStatus == null ? "" : poStatus;
    }

    public String getUnit() {
        return unit;
    }

    public void setUnit(String unit) {
        this.unit = unit == null ? "" : unit;
    }

    public String getRemark355() {
        return remark355;
    }

    public void setRemark355(String remark355) {
        this.remark355 = remark355 == null ? "" : remark355;
    }

    public LocalDate getFinalEta() {
        return finalEta;
    }

    public void setFinalEta(LocalDate finalEta) {
        this.finalEta = finalEta;
    }

    public double getReq() {
        return req;
    }

    public void setReq(double req) {
        this.req = req;
    }

    public double getRes() {
        return res;
    }

    public void setRes(double res) {
        this.res = res;
    }
}
