package com.nctech.certificategenerator;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.Color;
import java.awt.Font;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class CertificateGenerator {
    public static void main(String[] args) {
        try {

            FileInputStream excelFile = new FileInputStream(new File("C:\\Abhi\\My Code\\Java\\CertificateGenerator\\students.xlsx"));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet sheet = workbook.getSheetAt(0);

            int startingRow = 2;

            BufferedImage certificateTemplate = ImageIO.read(new File("C:\\Abhi\\My Code\\Java\\CertificateGenerator\\certificate.png"));
            for (int rowIndex = startingRow; rowIndex <= sheet.getLastRowNum(); rowIndex++) {

                Row row = sheet.getRow(rowIndex);
                if (row == null) {
                    continue;
                }
                Cell cell = row.getCell(2);
                if (cell == null) {

                    continue;
                }
                String studentName = cell.getStringCellValue();
                System.out.println(studentName);

                BufferedImage certificateCopy = new BufferedImage(
                        certificateTemplate.getWidth(),
                        certificateTemplate.getHeight(),
                        BufferedImage.TYPE_INT_ARGB
                );

                Graphics2D graphics = certificateCopy.createGraphics();
                graphics.drawImage(certificateTemplate, 0, 0, null);
                graphics.setColor(Color.BLACK);
                graphics.setFont(new Font("Arial", Font.BOLD, 76));
                graphics.drawString(studentName, 100, 350);

                ImageIO.write(certificateCopy, "PNG", new FileOutputStream("C:\\Abhi\\My Code\\Java\\CertificateGenerator\\"+studentName + ".png"));

                graphics.dispose();
            }

            // Close the Excel workbook
            workbook.close();
            excelFile.close();
        } catch (IOException e) {
            System.out.println("Application Failed");
        }
    }
}

