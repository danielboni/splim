package br.edu.ifrn.ip.splim.reading;

import java.io.File;
import java.io.IOException;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class ReadingTest {

	public static void main(String[] args) throws BiffException, IOException {
		
		Workbook workbook = Workbook.getWorkbook(new File("pubscopus.xls"));
	
		Sheet sheet = workbook.getSheet(0);
		
		int linhas = sheet.getRows();
		
		System.out.println("Número de linhas: " + linhas);
		System.out.println("Iniciando leitura de linhas: ");
		
		for (int i = 0; i < linhas; i++) {
			Cell a1 = sheet.getCell(0, i);
			Cell a2 = sheet.getCell(1, i);
			Cell a3 = sheet.getCell(2, i);
			
			String as1 = a1.getContents();
			String as2 = a2.getContents();
			String as3 = a3.getContents();
			
			System.out.println("Coluna 1: " + as1);
			System.out.println("Coluna 2: " + as2);
			System.out.println("Coluna 3: " + as3);
		}
	}

}
