import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Iterator;

public class Demo {

    public static void main(String[] args) throws IOException {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("TEST");
        int rowNo=0;
        Row row = sheet.createRow(rowNo++);
        int cellNum = 0;
        Cell cell = row.createCell(cellNum++);
        cell.setCellValue("DECIMAL");
        cell = row.createCell(cellNum++);
        cell.setCellValue(2.5);

        cellNum=0;
        row = sheet.createRow(rowNo++);
        cell = row.createCell(cellNum++);
        cell.setCellValue("TEXT");
        cell = row.createCell(cellNum++);
        cell.setCellValue("TEXT DATA");

        FileOutputStream fileOut = new FileOutputStream(new File("C://Users//okayo//Desktop//example.xlsx"));
        wb.write(fileOut);
        fileOut.close();
        System.out.println("Excel file is created");

        readExcelFile();
    }

    private static void readExcelFile() throws IOException {
        FileInputStream fileIn = new FileInputStream(new File("C://Users//okayo//Desktop//example.xlsx"));
        Workbook wb = new XSSFWorkbook(fileIn);
        Sheet sheet =  wb.getSheetAt(0);
        Iterator<Row> it = sheet.iterator();

        while(it.hasNext()){
            Row row = it.next();
            Iterator<Cell> ci = row.iterator();
            while(ci.hasNext()){
                Cell cell = ci.next();
                if(cell.getCellType() == CellType.STRING){
                    System.out.println("STRING => " + cell.getStringCellValue());
                }else if(cell.getCellType() == CellType.NUMERIC){
                    System.out.println("NUMERIC => " + cell.getNumericCellValue());
                }else if(cell.getCellType() == CellType.FORMULA){
                    System.out.println("FORMULA => " + cell.getCellFormula());
                }else{
                    System.out.println("Other cell type...");
                }
            }
        }
    }
}
