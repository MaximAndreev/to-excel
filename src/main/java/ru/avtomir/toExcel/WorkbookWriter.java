package ru.avtomir.toExcel;

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;

public interface WorkbookWriter {
    void setSheetName(String name);

    void setTableBody(List<Map<String, String>> body);

    void setHeadersOrder(List<String> headers);

    void write(OutputStream os) throws IOException;
}
