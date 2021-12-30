package com.example.springbootexcel.application.impl;

import com.example.springbootexcel.domain.model.User;
import java.io.IOException;
import java.util.List;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelExporter {

  private static final int ID_COLUMN = 0;

  private static final int NAME_COLUMN = 1;

  private static final int AGE_COLUMN = 2;

  private final XSSFWorkbook workbook;
  private XSSFSheet sheet;
  private final List<User> userList;

  public ExcelExporter(final List<User> userList) {
    this.userList = userList;
    workbook = new XSSFWorkbook();
  }

  private void writeHeaderLine() {
    sheet = workbook.createSheet("Technologies");

    Row row = sheet.createRow(0);

    CellStyle style = workbook.createCellStyle();

    createCell(row, ID_COLUMN, "ID", style);
    createCell(row, NAME_COLUMN, "NAME", style);
    createCell(row, AGE_COLUMN, "AGE", style);
  }

  private void createCell(
      final Row row, final int columnCount, final Object value, final CellStyle style) {
    sheet.autoSizeColumn(columnCount);
    Cell cell = row.createCell(columnCount);
    if (value instanceof Integer) {
      cell.setCellValue((Integer) value);
    } else if (value instanceof Boolean) {
      cell.setCellValue((Boolean) value);
    } else {
      cell.setCellValue((String) value);
    }
    cell.setCellStyle(style);
  }

  private void writeDataLines() {
    int rowCount = 1;

    CellStyle style = workbook.createCellStyle();

    for (User u : userList) {
      Row row = sheet.createRow(rowCount++);

      createCell(row, ID_COLUMN, u.getId().intValue(), style);
      createCell(row, NAME_COLUMN, u.getName(), style);
      createCell(row, AGE_COLUMN, u.getAge(), style);
    }
  }

  public void export(final HttpServletResponse response) throws IOException {
    writeHeaderLine();
    writeDataLines();

    ServletOutputStream outputStream = response.getOutputStream();

    workbook.write(outputStream);
    workbook.close();

    outputStream.close();
  }
}
