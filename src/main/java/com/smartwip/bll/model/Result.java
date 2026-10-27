package com.smartwip.bll.model;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * Aggregates results, statistics, the list of changes and the exception log after the bot runs.
 */
public class Result {
    private final Map<String, Integer> stats = new LinkedHashMap<>();
    private final List<LogEntry> log = new ArrayList<>();
    private final List<ChangeRow> changes = new ArrayList<>();
    private int groups;
    private double totalShortYds;

    public Result() {
    }

    public void inc(String key) {
        stats.merge(key, 1, Integer::sum);
    }

    public int stat(String key) {
        Integer v = stats.get(key);
        return v == null ? 0 : v;
    }

    public void log(String severity, String code, Group group, int excelRow, String message) {
        log.add(new LogEntry(severity, code, group, excelRow, message));
        inc("log." + code);
    }

    public Map<String, Integer> getStats() {
        return stats;
    }

    public List<LogEntry> getLog() {
        return log;
    }

    public List<ChangeRow> getChanges() {
        return changes;
    }

    public int getGroups() {
        return groups;
    }

    public void setGroups(int groups) {
        this.groups = groups;
    }

    public double getTotalShortYds() {
        return totalShortYds;
    }

    public void setTotalShortYds(double totalShortYds) {
        this.totalShortYds = totalShortYds;
    }
}
