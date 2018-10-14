package br.edu.ifrn.ip.splim.reading;

import java.io.File;
import java.io.IOException;
import java.util.LinkedList;
import java.util.List;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class DataCleaningRobot {

	/*
	 * Find and return lines in which at least one keyword in a list (keywords)
	 * appear in a given column.
	 */
	public List<Integer> findAndClean(String columnName, String[] keywords, String fileName) throws BiffException, IOException {
		if (keywords.length == 0) {
			return new LinkedList<Integer>();
		} else {
			List<Integer> selectedLines = new LinkedList<Integer>();
			
			Workbook workbook = Workbook.getWorkbook(new File(fileName));
			
			// Considering that we have only one tab
			Sheet sheet = workbook.getSheet(0);
			
			int numberOfLines = sheet.getRows();
			int numberOfColums = sheet.getColumns();
			
			int columnNumber = 0;
			
			// identifying the column with a given columnName
			for (int i = 0; i < numberOfColums; i++) {
				Cell cell = sheet.getCell(i, 0);
				
				if (cell.getContents().equalsIgnoreCase(columnName)) {
					columnNumber = i;
					break;
				}
			}
			
			System.out.println("Achei a coluna: " + columnNumber);
			
			for (int i = 0; i < numberOfLines; i++) {
				Cell cell = sheet.getCell(columnNumber, i);
				String cellContent = cell.getContents().toLowerCase();
				
				boolean flag = false;
				
				for (String keyword : keywords) {
					if(cellContent.contains(keyword.toLowerCase())) {
						selectedLines.add(i);
						flag = true;
						break;
					}
					
					if (flag) {
						break;
					}
				}
			}
			return selectedLines;
		}	
	}
	
	public void writeToACleanSheet(List<Integer> excludedLines, String nomeDoArquivoDeEntrada, String nomeDoArquivoDeSaida) throws IOException, RowsExceededException, WriteException, BiffException {
		WritableWorkbook cleanedWorkbook;

		Workbook wbook = Workbook.getWorkbook(new File(nomeDoArquivoDeEntrada));
		Sheet wbookSheet = wbook.getSheet(0);
		
		cleanedWorkbook = Workbook.createWorkbook(new File(nomeDoArquivoDeSaida));
		WritableSheet sheet = cleanedWorkbook.createSheet("Folha1", 0);

		int numeroDeColunas = wbookSheet.getColumns();
		int numeroDeLinhas = wbookSheet.getRows();
		
		// Uso o updatedLine para escrever no arquivo eliminando as linhas
		// que ficavam em branco, as linhas selecionadas para exclusão
		int updatedLine;
		for (int coluna = 0; coluna < numeroDeColunas; coluna++) {
			updatedLine = 0;
			for (int linha = 0; linha < numeroDeLinhas; linha++) {
				if (excludedLines.contains(linha)) {
					continue;
				}
				WritableCellFormat cf2 = new WritableCellFormat();
				Cell cell = wbookSheet.getCell(coluna, linha);
				Label label = new Label(coluna, updatedLine++, cell.getContents(), cf2);
				sheet.addCell(label);
			}
		}
		cleanedWorkbook.write();
		cleanedWorkbook.close();
	}
	
	/*
	 * Alguns abstracts, por conter vírgulas, estão desconfiguras na tabela. 
	 * Aparecendo parte por parte em células diferentes. Este método tem por
	 * funcão formatar esses abstracts em uma só coluna, conforme o esperado.
	 */
	public void formatAbstracts(String nomeDoArquivoDeEntrada, String novoArquivoDeSaida) throws BiffException, IOException, RowsExceededException, WriteException {
		// 1 - ler as colunas e identificar a coluna abstract
		// 2 - a partir dessa coluna montar um stringbuilder com o abstract
		//     completo separado por vírgula
		// 3 - reescrever a tabela atualizando o valor da coluna abstract
		
		Workbook arquivoDeEntrada = Workbook.getWorkbook(new File(nomeDoArquivoDeEntrada));
		Sheet abaNoArquivoDeEntrada = arquivoDeEntrada.getSheet(0);
		
		WritableWorkbook arquivoDeSaida;

		arquivoDeSaida = Workbook.createWorkbook(new File(novoArquivoDeSaida));
		WritableSheet abaNoArquivoDeSaida = arquivoDeSaida.createSheet("Folha1", 0);

		int numeroDeColunas = abaNoArquivoDeEntrada.getColumns();
		int numeroDeLinhas = abaNoArquivoDeEntrada.getRows();
		int colunaDoAbstract = 0;
		
		// identifying the column with a given columnName
		for (int i = 0; i < numeroDeColunas; i++) {
			Cell cell = abaNoArquivoDeEntrada.getCell(i, 0);

			if (cell.getContents().equalsIgnoreCase("abstract")) {
				colunaDoAbstract = i;
				break;
			}
		}
		
		StringBuilder abstractFormatado = new StringBuilder();
		
		for (int linha = 0; linha < numeroDeLinhas; linha++) {	
			if (linha == 0) {
				for (int coluna = 0; coluna < numeroDeColunas; coluna++) {
					Cell cell = abaNoArquivoDeEntrada.getCell(coluna, linha);
					WritableCellFormat cf2 = new WritableCellFormat();
					Label label = new Label(coluna, linha, cell.getContents(), cf2);
					abaNoArquivoDeSaida.addCell(label);
				}
			} else {
				for (int coluna = 0; coluna < numeroDeColunas; coluna++) {
					if(coluna == colunaDoAbstract) {
						// zerando string builder
						abstractFormatado.setLength(0);
						
						Cell cell = abaNoArquivoDeEntrada.getCell(coluna, linha);
						abstractFormatado.append(cell.getContents());
						
						// laço para concatenar o abstract corretamente.
						for(int i = coluna+1; i < numeroDeColunas; i++) {
							cell = abaNoArquivoDeEntrada.getCell(i, linha);
							abstractFormatado.append("," + cell.getContents());
						}
						WritableCellFormat cf2 = new WritableCellFormat();
						Label label = new Label(coluna, linha, abstractFormatado.toString(), cf2);
						abaNoArquivoDeSaida.addCell(label);
						
					} else if(coluna < colunaDoAbstract) {
						WritableCellFormat cf2 = new WritableCellFormat();
						Cell cell = abaNoArquivoDeEntrada.getCell(coluna, linha);
						Label label = new Label(coluna, linha, cell.getContents(), cf2);
						abaNoArquivoDeSaida.addCell(label);
					}
				}
			}
		}
		arquivoDeSaida.write();
		arquivoDeSaida.close();
		System.out.println("Abstracts formatados.");
	}
}
