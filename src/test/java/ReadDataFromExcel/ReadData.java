package ReadDataFromExcel;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;


public class ReadData {

    public static void main(String[] args) throws IOException {
        String file = System.getProperty("user.dir") + "./Files/TechUp.xlsx";
        FileInputStream ip = new FileInputStream(file);
        XSSFWorkbook workbook = new XSSFWorkbook(ip);
        XSSFSheet sheet = workbook.getSheetAt(0);
        int rows = sheet.getLastRowNum();

        for (int i = 0; i < rows; i++) {
            XSSFRow r = sheet.getRow(i);

            for (int j = 0; j < 5; j++) {    //reading only 5 columns as mentioned
                XSSFCell c = r.getCell(j);


                    switch (c.getCellType()) {

                        case BLANK:
                            System.out.println("No data!");
                            break;
                        case STRING:
                            System.out.println(c.getStringCellValue());
                            break;
                        case NUMERIC:
                            System.out.println(c.getNumericCellValue());
                    }

                }
            }
        }
    }
