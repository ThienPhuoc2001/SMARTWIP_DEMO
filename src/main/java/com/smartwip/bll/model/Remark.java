package com.smartwip.bll.model;

import java.time.LocalDate;

/**
 * Result of parsing a Purchasing department Remark 355.
 */
public class Remark {
    private String normalized;
    private LocalDate etd;
    private LocalDate eta;
    private String error;

    public Remark() {
    }

    public String getNormalized() {
        return normalized;
    }

    public void setNormalized(String normalized) {
        this.normalized = normalized;
    }

    public LocalDate getEtd() {
        return etd;
    }

    public void setEtd(LocalDate etd) {
        this.etd = etd;
    }

    public LocalDate getEta() {
        return eta;
    }

    public void setEta(LocalDate eta) {
        this.eta = eta;
    }

    public String getError() {
        return error;
    }

    public void setError(String error) {
        this.error = error;
    }
}
