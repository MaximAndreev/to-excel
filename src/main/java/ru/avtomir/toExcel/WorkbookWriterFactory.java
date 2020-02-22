package ru.avtomir.toExcel;

public class WorkbookWriterFactory {

    public final static String DEFAULT_SHEET_NAME = "Sheet";

    private WorkbookWriterFactory() {
    }

    public static WorkbookWriter getNoStyles() {
        WorkbookWriter ww = new NoStyleWorkbookWriter();
        ww.setSheetName(DEFAULT_SHEET_NAME);
        return ww;
    }
}
