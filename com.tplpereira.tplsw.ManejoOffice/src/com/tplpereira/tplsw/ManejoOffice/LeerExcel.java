package com.tplpereira.tplsw.ManejoOffice;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.model.WorkbookRecordList;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class LeerExcel {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		try {
			XSSFRow row;
			File file = new File("PRUEBA.xlsx");
			FileInputStream fInputStream = new FileInputStream(file);
			XSSFWorkbook workbook = new XSSFWorkbook(fInputStream);
			if (file.isFile() && file.exists()) {
				List<Generico> genericos = new ArrayList<>();
				Generico generico;
				XSSFSheet sheet = workbook.getSheetAt(0);
				Iterator<Row> rowIterator = sheet.iterator();
				int i = 0, j;
				while (rowIterator.hasNext()) {
					row = (XSSFRow) rowIterator.next();
					Iterator<Cell> cellIterator = row.cellIterator();
					j = 0;
					generico = i > 0 ? new Generico() : null;
					while (cellIterator.hasNext()) {
						Cell cell = cellIterator.next();
						if (i > 0) {
							switch (cell.getCellType()) {
							case Cell.CELL_TYPE_NUMERIC:
								System.out.println(cell.getNumericCellValue() + "\t\t");
								generico.setId((int)cell.getNumericCellValue());
								break;
							case Cell.CELL_TYPE_STRING:
								System.out.println(cell.getStringCellValue() + "\t\t");
								System.out.println(cell.getStringCellValue());
								System.out.println("j = " + j);
								if (j == 1) {
									generico.setNombre(cell.getStringCellValue());
								} else {
									generico.setDescripcion(cell.getStringCellValue());
								}
								break;
							case Cell.CELL_TYPE_BOOLEAN:
								System.out.println(cell.getBooleanCellValue() + "\t\t");
								generico.setActivo(cell.getBooleanCellValue());
								break;
							default:
								break;
							}
						}
						j++;
					}
					if (generico != null) {
						genericos.add(generico);
					}
					System.out.println();
					i++;
				}
				fInputStream.close();
				System.out.println("Tamaño = " + genericos.size());
			} else {
				System.out.println("El archivo no pudo ser abierto. Lastima");
			}
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
		}
	}

}
