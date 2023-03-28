package org.example;

//import java.io.File;
//import java.io.FileOutputStream;
//import java.io.IOException;
//import java.util.Iterator;
//import org.apache.poi.ss.usermodel.Cell;
//import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.ss.usermodel.Sheet;
//import org.apache.poi.ss.usermodel.Workbook;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/*--------------------------------------------------------------------------------------------------------------------*/

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Map;
import javax.xml.stream.XMLOutputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamWriter;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/*--------------------------------------------------------------------------------------------------------------------*/

/**
 * @author : D.D.Sandaruwan <dulanjayasandaruwan1998@gmail.com>
 * @Since : 23/03/2023
 **/
public class Main {
    public static void main(String[] args) throws IOException, InvalidFormatException {

        String inputFile = "C:\\Users\\Dulanjaya Sandaruwan\\Downloads\\input.xlsx";
        String outputFile = "C:\\Users\\Dulanjaya Sandaruwan\\Downloads\\test7.xml";

        try {
            Workbook workbook = new XSSFWorkbook(inputFile);
            Sheet sheet = workbook.getSheetAt(0);

            FileOutputStream outputStream = new FileOutputStream(new File(outputFile));
            outputStream.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n".getBytes());
            outputStream.write("<data>\n".getBytes());

            Iterator<Row> rowIterator = sheet.iterator();

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                outputStream.write("  <row>\n".getBytes());
                Iterator<Cell> cellIterator = row.iterator();

                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    outputStream.write(("    <column>" + cell.toString() + "</column>\n").getBytes());
                }

                outputStream.write("  </row>\n".getBytes());
            }

            outputStream.write("</data>".getBytes());
            outputStream.close();
            System.out.println("Excel file converted to XML file successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }

/*--------------------------------------------------------------------------------------------------------------------*/

//        File inputFile = new File("C:\\Users\\Dulanjaya Sandaruwan\\Downloads\\input.xlsx");
//        Workbook workbook = new XSSFWorkbook(inputFile);
//        Sheet sheet = workbook.getSheetAt(0);
//
//        Map<String, String> headerMap = new LinkedHashMap<>();
//        Row headerRow = sheet.getRow(0);
//        Iterator<Cell> headerCellIterator = headerRow.cellIterator();
//        while (headerCellIterator.hasNext()) {
//            Cell headerCell = headerCellIterator.next();
//            headerMap.put("col_" + headerCell.getColumnIndex(), headerCell.getStringCellValue());
//        }
//
//        XMLOutputFactory outputFactory = XMLOutputFactory.newInstance();
//        XMLStreamWriter writer = outputFactory.createXMLStreamWriter(new FileOutputStream("C:\\Users\\Dulanjaya Sandaruwan\\Downloads\\test.xml"), "UTF-8");
//        writer.writeStartDocument("UTF-8", "1.0");
//        writer.writeStartElement("data");
//
//        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
//            Row row = sheet.getRow(i);
//            Iterator<Cell> cellIterator = row.cellIterator();
//
//            writer.writeStartElement("row");
//
//            while (cellIterator.hasNext()) {
//                Cell cell = cellIterator.next();
//                String headerName = headerMap.get("col_" + cell.getColumnIndex());
//                writer.writeStartElement(headerName);
//                writer.writeCharacters(cell.getStringCellValue());
//                writer.writeEndElement();
//            }
//
//            writer.writeEndElement();
//        }
//
//        writer.writeEndElement();
//        writer.writeEndDocument();
//        writer.flush();
//        writer.close();
//
//        System.out.println("Conversion completed successfully!");

/*--------------------------------------------------------------------------------------------------------------------*/

//        try {
//
//            Workbook workbook = new XSSFWorkbook(new File("C:\\Users\\Dulanjaya Sandaruwan\\Downloads\\input.xlsx"));
//
//            Sheet sheet = workbook.getSheetAt(0);
//
//            FileOutputStream out = new FileOutputStream(new File("C:\\Users\\Dulanjaya Sandaruwan\\Downloads\\example.xml"));
//
//            out.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n".getBytes());
//            out.write("<data>\n".getBytes());
//
//            Iterator<Row> rowIterator = sheet.iterator();
//            while (rowIterator.hasNext()) {
//                Row row = rowIterator.next();
//
//                if (row.getRowNum() == 0) {
//                    continue;
//                }
//
//                out.write("  <record>\n".getBytes());
//
//                Iterator<Cell> cellIterator = row.iterator();
//                while (cellIterator.hasNext()) {
//                    Cell cell = cellIterator.next();
//
//                    String columnName = sheet.getRow(0).getCell(cell.getColumnIndex()).getStringCellValue();
//
//                    out.write(("    <" + columnName + ">" + cell.getStringCellValue() + "</" + columnName + ">\n").getBytes());
//                }
//
//                out.write("  </record>\n".getBytes());
//            }
//
//            out.write("</data>\n".getBytes());
//
//            out.close();
//
//            workbook.close();
//
//            System.out.println("Excel file converted to XML file successfully!");
//        } catch (IOException e) {
//            e.printStackTrace();
//        }

    }
}