package fileHandling;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class ReadDataFromSheet {


    public static void writeExcel(String filePath, String fileName) throws IOException {

        File file = new File(filePath);
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet spreadsheet = workbook.createSheet(" Employee Info ");
        FileOutputStream out = new FileOutputStream(filePath + "datasheet1.xlsx");


        FileInputStream inputStream = new FileInputStream(file + "/" + fileName);
        XSSFWorkbook workbook1 = new XSSFWorkbook(inputStream);
        XSSFSheet sheet = workbook1.getSheetAt(0);
        int rowCount = sheet.getLastRowNum();

        for (int j = 0; j < rowCount; j++) {
            Row row = sheet.getRow(j);
            Row newSheetRow = spreadsheet.createRow(j);
            int cellCount = row.getLastCellNum();
            int cellId = 0;
            for (int k = 0; k < cellCount; k++) {

                if (k == 1 || k == 3 || k == 4) {
                    Cell newCell = newSheetRow.createCell(cellId++);
                    newCell.setCellValue(String.valueOf(row.getCell(k)));
                }
            }
        }

        workbook.write(out);
        out.close();
        System.out.println("xlsx written successfully");


    }

    public static void main(String... strings) throws IOException {

        writeExcel(System.getProperty("user.dir") + "src/dataFile/", "datasheet.xlsx");

    }

}

