package br.edu.ifrn.ip.splim.main;

import java.io.IOException;
import java.util.List;

import br.edu.ifrn.ip.splim.reading.DataCleaningRobot;
import jxl.read.biff.BiffException;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class Main {

	public static void main(String[] args) throws BiffException, IOException, RowsExceededException, WriteException {
		DataCleaningRobot robot = new DataCleaningRobot();
		String[] keywords = {"regularization approach"};
		List<Integer> selectedLines = robot.findAndClean("title", keywords);
		
		System.out.println("Linhas selecionadas para serem excluídas: ");
		for (Integer line: selectedLines) {
			System.out.println("-> " + (line + 1));
		}
		
		robot.writeToACleanSheet(selectedLines);
		
	}
}
