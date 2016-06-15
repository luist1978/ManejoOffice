package com.tplpereira.tplsw.ManejoOffice;

import java.io.File;
import java.io.FileOutputStream;
import java.time.LocalDate;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class EscribirExcel {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		try {
			XSSFWorkbook workbook = new XSSFWorkbook();
			XSSFSheet sheetHoja1 = workbook.createSheet("HOJA1");
			XSSFSheet sheetHoja2 = workbook.createSheet("HOJA2");
			XSSFSheet sheetHoja3 = workbook.createSheet("HOJA3");
			XSSFSheet sheetHoja4 = workbook.createSheet("HOJA4");
			XSSFSheet sheetHoja5 = workbook.createSheet("HOJA5");
			FileOutputStream fOutputStream = new FileOutputStream(
					new File("ARCHIVO FECHA " + LocalDate.now() + ".xlsx"));
			workbook.write(fOutputStream);
			fOutputStream.close();
			System.out.println("Se creo el archivo con exito!!!");
			System.out.println("Fin.");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
