package com.advancia;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingApachePoiApplication {
	
	private static final String FILE_PATH = "C:/Users/marco/_Advancia/Formazione/1-JAVA/APACHE-POI-AND-ITEXT/reading-apache-poi/file.xlsx";
	
    public static void main( String[] args ) {
        try (FileInputStream inputStream = new FileInputStream(new File(FILE_PATH))) {
        	
        	Workbook workbook = new XSSFWorkbook(inputStream);
        	
        	DataFormatter dataFormatter = new DataFormatter();
        	
        	Iterator<Sheet> sheetIterator = workbook.sheetIterator();
        	
        	while(sheetIterator.hasNext()) {
        		Sheet sheet = sheetIterator.next();
        		System.out.println("Sheet name is " + "'" + sheet.getSheetName() + "'");
        		System.out.println("---------------");
        		Iterator<Row> rowIterator = sheet.rowIterator();
        		while (rowIterator.hasNext()) {
        			Row row = rowIterator.next();
        			Iterator<Cell> cellIterator = row.cellIterator();
        			while (cellIterator.hasNext()) {
						Cell cell = cellIterator.next();
						String cellValue = dataFormatter.formatCellValue(cell);
//						if (cell.getCellType() == CellType.STRING) {
//							
//						}
						System.out.print(cellValue + " | ");
					}
        			System.out.println();
				}
        		System.out.println("---------------");
        	}
        	
        	workbook.close();
			
		} catch (Exception e) {
			e.printStackTrace();
		}
    }
}
