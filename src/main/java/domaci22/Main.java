package domaci22;

import com.github.javafaker.Faker;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class Main {
    public static void main(String[] args) {
        System.out.println("Two rows from existing file are:");
        ApachePoiUtil.readNameSurnameTwoRows("domaci22.xlsx");
        System.out.println();

        try {
            addNamesWithFaker("domaci22.xlsx");
            System.out.println();
        } catch (FileNotFoundException e) {
            System.out.println("File not found!");
        } catch (IOException e) {
            System.out.println("File invalid for writing!");
        }

        System.out.println("Rows 2-10 are:");
        ApachePoiUtil.readNameSurnameEightRows("domaci22.xlsx");
    }

    public static void addNamesWithFaker(String fileName) throws FileNotFoundException, IOException {
        FileInputStream inputStream = new FileInputStream(
                new File(fileName));

        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        XSSFSheet sheet = workbook.getSheet("Sheet1");

        Faker faker = new Faker();

        // 8 more rows with random name and surname
        for (int i = 2; i < 10; i++) {
            String name = faker.name().firstName();
            String surname = faker.name().lastName();

            XSSFRow row = sheet.createRow(i);
            XSSFCell cell1 = row.createCell(0);
            cell1.setCellValue(name);
            XSSFCell cell2 = row.createCell(1);
            cell2.setCellValue(surname);
        }

        FileOutputStream fileOutputStream = new FileOutputStream(new File(fileName));
        workbook.write(fileOutputStream);
        fileOutputStream.close();
    }
}
