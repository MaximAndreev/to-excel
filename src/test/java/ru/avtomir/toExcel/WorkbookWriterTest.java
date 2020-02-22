package ru.avtomir.toExcel;

import org.apache.poi.hssf.usermodel.HSSFWorkbookFactory;
import org.apache.poi.ss.usermodel.*;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

import static org.junit.jupiter.api.Assertions.assertEquals;

class WorkbookWriterTest {

    private WorkbookWriter ww;

    @BeforeEach
    public void setUp() {
        ww = WorkbookWriterFactory.getNoStyles();
    }

    @Test
    public void sheetNameIsNull_throwNPE() {
        Assertions.assertThrows(NullPointerException.class, () -> ww.setSheetName(null));
    }

    @Test
    public void sheetBodyIsNull_throwNPE() {
        Assertions.assertThrows(NullPointerException.class, () -> ww.setTableBody(null));
    }

    @Test
    public void sheetHeadersOrderIsNull_throwNPE() {
        Assertions.assertThrows(NullPointerException.class, () -> ww.setHeadersOrder(null));
    }

    @Nested
    public class SheetIsCreated {

        private String sheetName;
        private List<Map<String, String>> tableBody;
        private List<String> headersOrder;

        @BeforeEach
        public void setUp() {
            sheetName = "MySheet";
            tableBody = List.of(
                    Map.of("ColumnName 1", "Row0Cell0", "ColumnName 2", "Row0Cell1"),
                    Map.of("ColumnName 1", "Row1Cell1", "ColumnName 2", "Row1Cell1"),
                    Map.of("ColumnName 1", "Row2Cell1", "ColumnName 2", "Row2Cell1"));
            headersOrder = List.of("ColumnName 2", "ColumnName 1");
        }

        @Test
        public void sheetNameIsDefault() throws IOException {
            // given
            ByteArrayOutputStream os = new ByteArrayOutputStream();

            // when
            ww.write(os);

            // then
            assertSheetNameDefault(os);
        }

        private void assertSheetNameDefault(ByteArrayOutputStream os) throws IOException {
            Workbook workbook = getWorkbook(os);
            Sheet sheet = workbook.getSheetAt(0);
            assertEquals(WorkbookWriterFactory.DEFAULT_SHEET_NAME, sheet.getSheetName());
        }

        private Workbook getWorkbook(ByteArrayOutputStream os) throws IOException {
            InputStream is = new ByteArrayInputStream(os.toByteArray());
            return HSSFWorkbookFactory.create(is);
        }

        @Test
        public void sheetNameIsNonDefault() throws IOException {
            // given
            ByteArrayOutputStream os = new ByteArrayOutputStream();

            // when
            ww.setSheetName(sheetName);
            ww.write(os);

            // then
            assertSheetName(os);
        }

        private void assertSheetName(ByteArrayOutputStream os) throws IOException {
            Workbook workbook = getWorkbook(os);
            Sheet sheet = workbook.getSheetAt(0);
            assertEquals(sheetName, sheet.getSheetName());
        }

        @Test
        public void tableBodyIsEmpty() throws IOException {
            // given
            ByteArrayOutputStream os = new ByteArrayOutputStream();
            List<Map<String, String>> body = Collections.emptyList();

            // when
            ww.setSheetName(sheetName);
            ww.setTableBody(body);
            ww.write(os);

            // then
            assertNoContent(os);
        }

        private void assertNoContent(ByteArrayOutputStream os) throws IOException {
            Workbook workbook = getWorkbook(os);
            Sheet sheet = workbook.getSheet(sheetName);
            int firstRowNum = sheet.getFirstRowNum();
            assertEquals(-1, firstRowNum);
        }

        @Test
        public void tableBodyIsNonEmptyAndHeadersOrderNotSet() throws IOException {
            // given
            ByteArrayOutputStream os = new ByteArrayOutputStream();

            // when
            ww.setSheetName(sheetName);
            ww.setTableBody(tableBody);
            ww.write(os);

            // then
            assertContentIsEqualToTableBody(os);
        }

        private void assertContentIsEqualToTableBody(ByteArrayOutputStream os) throws IOException {
            Workbook workbook = getWorkbook(os);
            Sheet sheet = workbook.getSheet(sheetName);
            List<LinkedHashMap<String, String>> tableContent = readSheet(sheet);
            assertEquals(tableBody, tableContent);
        }

        private List<LinkedHashMap<String, String>> readSheet(Sheet sheet) {
            List<LinkedHashMap<String, String>> sheetContent = new ArrayList<>();
            List<String> headers = readHeaders(sheet);
            int rowN = 1;
            Row row;
            while ((row = sheet.getRow(rowN)) != null) {
                sheetContent.add(readRow(headers, row));
                rowN++;
            }
            return sheetContent;
        }

        private List<String> readHeaders(Sheet sheet) {
            List<String> headers = new ArrayList<>();
            Row row = sheet.getRow(0);
            int i = 0;
            Cell cell;
            while (isNotEmpty(cell = row.getCell(i))) {
                headers.add(getValueAsString(cell));
                i++;
            }
            return headers;
        }

        private String getValueAsString(Cell cell) {
            if (cell == null) {
                return "";
            }
            CellType cellType = cell.getCellType();
            switch (cellType) {
                case BOOLEAN:
                    return String.valueOf(cell.getBooleanCellValue());
                case NUMERIC:
                    return String.valueOf(cell.getNumericCellValue());
                case STRING:
                    return cell.getStringCellValue();
                case FORMULA:
                    return cell.getCellFormula();
                default:
                    return "";
            }
        }

        private LinkedHashMap<String, String> readRow(List<String> headers, Row row) {
            LinkedHashMap<String, String> rowAsMap = new LinkedHashMap<>();
            for (int cellN = 0; cellN < headers.size(); cellN++) {
                Cell cell = row.getCell(cellN);
                rowAsMap.put(headers.get(cellN), getValueAsString(cell));
            }
            return rowAsMap;
        }

        private boolean isNotEmpty(Cell cell) {
            return cell != null && cell.getCellType() != CellType.BLANK;
        }

        @Test
        public void tableBodyIsNonEmptyAndHeadersOrderSet() throws IOException {
            // given
            ByteArrayOutputStream os = new ByteArrayOutputStream();

            // when
            ww.setSheetName(sheetName);
            ww.setHeadersOrder(headersOrder);
            ww.setTableBody(tableBody);
            ww.write(os);

            // then
            assertContentHaveCorrectHeaderOrder(os, headersOrder);
        }

        private void assertContentHaveCorrectHeaderOrder(ByteArrayOutputStream os, List<String> headersOrder) throws IOException {
            Workbook workbook = getWorkbook(os);
            Sheet sheet = workbook.getSheet(sheetName);
            List<LinkedHashMap<String, String>> tableContent = readSheet(sheet);
            tableContent.forEach(row -> {
                List<String> actualHeaderOrder = new ArrayList<>(row.keySet());
                assertEquals(headersOrder, actualHeaderOrder);
            });
        }

        @Test
        public void tableBodyIsNonEmptyAndHeadersOrderSetButMissOneColumn() throws IOException {
            // given
            ByteArrayOutputStream os = new ByteArrayOutputStream();
            List<String> headersOrderWithoutOneColumn = headersOrder.subList(0, 1);

            // when
            ww.setSheetName(sheetName);
            ww.setHeadersOrder(headersOrderWithoutOneColumn);
            ww.setTableBody(tableBody);
            ww.write(os);

            // then
            assertContentHaveCorrectHeaderOrder(os, headersOrderWithoutOneColumn);
        }

        @Test
        public void tableBodyIsNonEmptyAndHeadersOrderSetButHaveOneExcessiveColumn() throws IOException {
            // given
            ByteArrayOutputStream os = new ByteArrayOutputStream();
            List<String> headersOrderWithExcessiveColumn = new ArrayList<>(headersOrder);
            headersOrderWithExcessiveColumn.add("SomeRandomColumnName");

            // when
            ww.setSheetName(sheetName);
            ww.setHeadersOrder(headersOrderWithExcessiveColumn);
            ww.setTableBody(tableBody);
            ww.write(os);

            // then
            assertContentHaveCorrectHeaderOrder(os, headersOrderWithExcessiveColumn);
        }
    }
}
