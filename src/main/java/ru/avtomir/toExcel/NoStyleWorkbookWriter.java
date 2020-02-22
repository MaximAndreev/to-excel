package ru.avtomir.toExcel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbookFactory;

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
        Workbook workbook = XSSFWorkbookFactory.createWorkbook();
        Sheet sheet = workbook.createSheet(sheetName);
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

    private void writeHeaders(Sheet sheet) {
        if (!tableHeaders.isEmpty()) {
            Row headerRow = sheet.createRow(0);
            writeCellValues(tableHeaders, headerRow);
        }
    }

    private void writeCellValues(List<String> values, Row row) {
        for (int i = 0; i < values.size(); i++) {
            Cell cell = row.createCell(i, CellType.STRING);
            cell.setCellValue(values.get(i));
        }
    }

    private void writeContent(Sheet sheet) {
        for (int i = 0; i < tableBody.size(); i++) {
            List<String> cellValues = getOrderedCellValuesForRow(i);
            if (!cellValues.isEmpty()) {
                Row row = sheet.createRow(i + 1);  // on the first row headers is written, so `+ 1`
                writeCellValues(cellValues, row);
            }
        }
    }

    private List<String> getOrderedCellValuesForRow(int i) {
        Map<String, String> rowAsMap = tableBody.get(i);
        return tableHeaders.stream()
                .map(header -> rowAsMap.getOrDefault(header, ""))
                .collect(Collectors.toList());
    }
}
