package com.smartwip.bll.model;

/**
 * Specialized exception for safety-rule violations or RPA bot operations.
 */
public class BotException extends Exception {
    private static final long serialVersionUID = 1L;
    private final String code;

    public BotException(String code, String message) {
        super(message);
        this.code = code;
    }

    public String getCode() {
        return code;
    }
}
