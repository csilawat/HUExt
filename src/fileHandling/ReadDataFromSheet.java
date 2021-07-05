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


    public static void writeExcel(String filePath, String fileName,
                                  String sheetName) throws IOException {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet spreadsheet = workbook.createSheet( " Employee Info ");
        FileOutputStream out = new FileOutputStream(new File("src/dataFile/datasheet1.xlsx"));

        File file = new File("src/dataFile/");
        FileInputStream inputStream = new FileInputStream(file + "/" + "datasheet.xlsx");

        XSSFWorkbook guru99Workbook = new XSSFWorkbook(inputStream);

        XSSFSheet sheet = guru99Workbook.getSheetAt(0);
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
                    System.out.println(row.getCell(k));
                }
            }
        }

        workbook.write(out);
        out.close();
        System.out.println("Writesheet.xlsx written successfully");


    }

    public static void main(String... strings) throws IOException {

        writeExcel(System.getProperty("user.dir") + "src/dataFile/", "datasheet.xlsx", "sheet1");

//
//        //Create blank workbook
//        XSSFWorkbook workbook = new XSSFWorkbook();
//        //Create a blank sheet
//        XSSFSheet spreadsheet = workbook.createSheet( " Employee Info ");
//        //Create row object
//        XSSFRow row1 = spreadsheet.createRow(0);
//        Cell cell = row1.createCell(0);
//        cell.setCellValue("testqwwqwq");
//        FileOutputStream out = new FileOutputStream(new File("src/dataFile/datasheet1.xlsx"));
//
//
//
//
//        File file = new File("src/dataFile/");
//        FileInputStream inputStream = new FileInputStream(file + "/" + "datasheet.xlsx");
//
//        XSSFWorkbook guru99Workbook = new XSSFWorkbook(inputStream);
//
//        XSSFSheet sheet = guru99Workbook.getSheetAt(0);
//        int rowCount = sheet.getLastRowNum();
//
//        for (int j = 0; j < rowCount; j++) {
//            Row row = sheet.getRow(j);
//            Row newSheetRow = spreadsheet.createRow(j);
//            int cellCount = row.getLastCellNum();
//            int cellId = 0;
//            for (int k = 0; k < cellCount; k++) {
//
//                if (k == 1 || k == 3 || k == 4) {
//                    Cell newCell = newSheetRow.createCell(cellId++);
//                    newCell.setCellValue(String.valueOf(row.getCell(k)));
//                    System.out.println(row.getCell(k));
//                }
//            }
//        }
//
//        workbook.write(out);
//        out.close();
//        System.out.println("Writesheet.xlsx written successfully");
    }

}

