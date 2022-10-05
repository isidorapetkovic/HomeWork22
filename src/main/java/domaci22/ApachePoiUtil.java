package domaci22;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class ApachePoiUtil {
    public static void readNameSurnameTwoRows(String path) {
        try {
            FileInputStream inputStream = new FileInputStream(
                    new File(path));

            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
            XSSFSheet sheet = workbook.getSheet("Sheet1");

            for (int j = 0; j < 2; j++) {
                XSSFRow row = sheet.getRow(j);

                for (int i = 0; i < 2; i++) {
                    XSSFCell cell = row.getCell(i);
                    String ime = cell.getStringCellValue();
                    System.out.print(ime + " ");
                }
                System.out.println();
            }
        } catch (FileNotFoundException ex) {
            System.out.println("FileNotFound.class");
        } catch (IOException e) {
            e.printStackTrace();
        } catch (NullPointerException e) {
            //e.printStackTrace();
        }
    }

    public static void readNameSurnameEightRows(String path) {
        try {
            FileInputStream inputStream = new FileInputStream(
                    new File(path));

            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
            XSSFSheet sheet = workbook.getSheet("Sheet1");

            for (int j = 2; j < 10; j++) {
                XSSFRow row = sheet.getRow(j);

                XSSFCell imeCell = row.getCell(0);
                XSSFCell prezimeCell = row.getCell(1);
                String ime = imeCell.getStringCellValue();
                String prezime = prezimeCell.getStringCellValue();

                if(ime=="" && prezime=="") {
                    break;
                }

                System.out.println(ime + " " + prezime);
            }
        } catch (FileNotFoundException ex) {
            System.out.println("FileNotFound.class");
        } catch (IOException e) {
            e.printStackTrace();
        } catch (NullPointerException e) {
            //e.printStackTrace();
        }
    }
}
