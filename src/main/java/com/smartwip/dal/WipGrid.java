package com.smartwip.dal;

import java.time.LocalDate;
import com.smartwip.bll.model.BotException;

/**
 * Abstracts read/write access to the WIP (Work In Progress) sheet.
 */
public interface WipGrid {
    /**
     * @return the number of valid data rows in the WIP sheet.
     */
    int size();

    /**
     * @param i internal row index (0-based)
     * @return the actual Excel row number (1-based)
     */
    int excelRow(int i);

    /**
     * @param i internal row index
     * @return the JOB_ERP code
     */
    String job(int i);

    /**
     * @param i internal row index
     * @return the Style code
     */
    String style(int i);

    /**
     * @param i internal row index
     * @return the Garment Color code
     */
    String color(int i);

    /**
     * @param i internal row index
     * @return the Order Status
     */
    String status(int i);

    /**
     * @param i internal row index
     * @return the current fabric ETA date in WIP
     */
    LocalDate eta(int i);

    /**
     * @param i internal row index
     * @return the current Remarks text in WIP
     */
    String remarks(int i);

    /**
     * @param i internal row index
     * @return the Material Status
     */
    String materialStatus(int i);

    /**
     * Write the new fabric ETA date (column AG).
     */
    void writeEta(int i, LocalDate d) throws BotException;

    /**
     * Write updated Remarks content (column BZ).
     */
    void writeRemarks(int i, String text) throws BotException;

    /**
     * Write the YDS fabric shortage quantity to the new Material_Short_Qty column.
     */
    void writeShort(int i, double yds) throws BotException;

    /**
     * Write the bot's note/explanation to the new Bot_Note column.
     */
    void writeNote(int i, String text) throws BotException;
}
