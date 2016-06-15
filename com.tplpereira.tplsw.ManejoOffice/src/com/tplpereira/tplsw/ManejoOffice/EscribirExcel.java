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
			XSSFSheet sheetResumen = workbook.createSheet("RESUMEN");
			XSSFSheet sheetUnitarioRedes = workbook.createSheet("UNITARIO REDES");
			XSSFSheet sheetCasaTipo = workbook.createSheet("CASA TIPO");
			XSSFSheet sheetUnitarioAcometidas = workbook.createSheet("UNITARIO ACOMETIDAS");
			XSSFSheet sheetAcometidas = workbook.createSheet("ACOMETIDAS");
			FileOutputStream fOutputStream = new FileOutputStream(
					new File("PRESUPUESTO APROBADO " + LocalDate.now() + ".xlsx"));
			workbook.write(fOutputStream);
			fOutputStream.close();
			System.out.println("Se creo el archivo con exito!!!");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
