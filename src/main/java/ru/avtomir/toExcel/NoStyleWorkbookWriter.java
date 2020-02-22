package ru.avtomir.toExcel;

import org.apache.poi.hssf.usermodel.*;

import java.io.IOException;
import java.io.OutputStream;
import java.util.*;
import java.util.stream.Collectors;

public class NoStyleWorkbookWriter implements WorkbookWriter {

    private String sheetName = "";
    private List<Map<String, String>> tableBody = Collections.emptyList();
    private List<String> tableHeaders = Collections.emptyList();

    @Override
    public void setSheetName(String name) {
        this.sheetName = Objects.requireNonNull(name, "Sheet name can't be null");
    }

    @Override
    public void setTableBody(List<Map<String, String>> body) {
        this.tableBody = Objects.requireNonNull(body, "Sheet body can't be null");
    }

    @Override
    public void setHeadersOrder(List<String> headers) {
        this.tableHeaders = Objects.requireNonNull(headers, "Sheet headers can't be null");
    }

    @Override
    public void write(OutputStream os) throws IOException {
        HSSFWorkbook workbook = HSSFWorkbookFactory.createWorkbook();
        HSSFSheet sheet = workbook.createSheet(sheetName);
        ifHeadersEmptyTryToGuessFromTableBody();
        writeHeaders(sheet);
        writeContent(sheet);
        workbook.write(os);
    }

    private void ifHeadersEmptyTryToGuessFromTableBody() {
        if (tableHeaders.isEmpty() && !tableBody.isEmpty()) {
            tableHeaders = new ArrayList<>(tableBody.get(0).keySet());
        }
    }

    private void writeHeaders(HSSFSheet sheet) {
        if (!tableHeaders.isEmpty()) {
            HSSFRow headerRow = sheet.createRow(0);
            writeCellValues(tableHeaders, headerRow);
        }
    }

    private void writeContent(HSSFSheet sheet) {
        for (int i = 0; i < tableBody.size(); i++) {
            List<String> cellValues = parseCellValues(i);
            if (!cellValues.isEmpty()) {
                HSSFRow row = sheet.createRow(i + 1);
                writeCellValues(cellValues, row);
            }
        }
    }

    private List<String> parseCellValues(int i) {
        Map<String, String> rowAsMap = tableBody.get(i);
        return tableHeaders.stream()
                .map(header -> rowAsMap.getOrDefault(header, ""))
                .collect(Collectors.toList());
    }

    private void writeCellValues(List<String> values, HSSFRow row) {
        for (int i = 0; i < values.size(); i++) {
            HSSFCell cell = row.createCell(i);
            cell.setCellValue(values.get(i));
        }
    }
}
