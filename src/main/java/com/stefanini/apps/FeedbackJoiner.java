/**
 * 
 */
package com.stefanini.apps;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author fpsouza
 *
 */
public class FeedbackJoiner {

	private static final int COMPLIANCE_ROW = 29;
	private static final int HONESTY_ROW = 28;
	private static final int POSTURE_ROW = 27;
	private static final int PUNCTUALITY_ROW = 26;
	private static final int MANAGER_RELATIONSHIP_ROW = 24;
	private static final int CUSTOMER_RELATIONSHIP_ROW = 23;
	private static final int PAIR_RELATIONSHIP_ROW = 22;
	private static final int CREATIVITY_ROW = 21;
	private static final int TEAM_WORK_ROW = 20;
	private static final int INITIATIVE_ROW = 19;
	private static final int INDEPENDENCE_ROW = 18;
	private static final int COMUNICATION_ROW = 17;
	private static final int GROWTH_ROW = 15;
	private static final int PRODUCTIVITY_ROW = 14;
	private static final int QUALITY_ROW = 13;
	private static final int TECH_KNOW_ROW = 12;
	private static final int COMPLIANCE_CELL = 17;
	private static final int HONESTY_CELL = 16;
	private static final int POSTURE_CELL = 16;
	private static final int PUNCTUALITY_CELL = 15;
	private static final int MANAGER_RELATIONSHIP_CELL = 14;
	private static final int CUSTOMER_RELATIONSHIP_CELL = 13;
	private static final int PAIR_RELATIONSHIP_CELL = 12;
	private static final int CREATIVITY_CELL = 11;
	private static final int TEAM_WORK_CELL = 10;
	private static final int INITIATIVE_CELL = 9;
	private static final int INDEPENDENCE_CELL = 8;
	private static final int COMUNICATION_CELL = 7;
	private static final int GROWTH_CELL = 6;
	private static final int PRODUCTIVITY_CELL = 5;
	private static final int QUALITY_CELL = 4;
	private static final int TECH_KNOW_CELL = 3;
	private static final int ROLE_CELL = 2;
	private static final int PROFESSIONAL_CELL = 1;
	private static final int MANAGER_CELL = 0;
	private static final String FEEDBACK_DIRECTORY = "C:\\Stefanini\\PJ00212 - BoB\\RH\\Feedbacks";

	public void join(String feedbackDirectory, String resultPlan) throws Exception {
		List<String> listFiles = this.listFiles(feedbackDirectory == null ? FEEDBACK_DIRECTORY : feedbackDirectory);
		Workbook workbookConsolidation = new XSSFWorkbook();
		Sheet shConsolidation = workbookConsolidation.createSheet();
		int rowCount = 0;

		mountHeader(shConsolidation, rowCount++);

		for (String fileName : listFiles) {
			if (fileName.endsWith("00_Controle.xlsx") || fileName.endsWith("feedbackConsolidado.xlsx")
					|| fileName.startsWith("~"))
				continue;
			System.out.println(fileName + " abrindo ...");
			FileInputStream arquivo = new FileInputStream(new File(fileName));
			Workbook workbook = new XSSFWorkbook(arquivo);
			Sheet sh = workbook.getSheet("Plan1");
			Row row = shConsolidation.createRow(rowCount++);

			extractFeedbackResult(sh, row);
			System.out.println(fileName + " concluído.");
			workbook.close();
			arquivo.close();
			arquivo.close();
		}

		FileOutputStream out = new FileOutputStream(
				FEEDBACK_DIRECTORY + (resultPlan == null ? "\\feedbackConsolidado.xlsx" : ("\\" + resultPlan)));
		workbookConsolidation.write(out);
		out.close();
	}

	private void extractFeedbackResult(Sheet sh, Row row) {
		Cell cell = row.createCell(MANAGER_CELL);
		cell.setCellValue(extractManager(sh));

		cell = row.createCell(PROFESSIONAL_CELL);
		cell.setCellValue(extractProfessional(sh));

		cell = row.createCell(ROLE_CELL);
		cell.setCellValue(extractRole(sh));

		cell = row.createCell(TECH_KNOW_CELL);
		cell.setCellValue(extractResult(sh, TECH_KNOW_ROW));

		cell = row.createCell(QUALITY_CELL);
		cell.setCellValue(extractResult(sh, QUALITY_ROW));

		cell = row.createCell(PRODUCTIVITY_CELL);
		cell.setCellValue(extractResult(sh, PRODUCTIVITY_ROW));

		cell = row.createCell(GROWTH_CELL);
		cell.setCellValue(extractResult(sh, GROWTH_ROW));

		cell = row.createCell(COMUNICATION_CELL);
		cell.setCellValue(extractResult(sh, COMUNICATION_ROW));

		cell = row.createCell(INDEPENDENCE_CELL);
		cell.setCellValue(extractResult(sh, INDEPENDENCE_ROW));

		cell = row.createCell(INITIATIVE_CELL);
		cell.setCellValue(extractResult(sh, INITIATIVE_ROW));

		cell = row.createCell(TEAM_WORK_CELL);
		cell.setCellValue(extractResult(sh, TEAM_WORK_ROW));

		cell = row.createCell(CREATIVITY_CELL);
		cell.setCellValue(extractResult(sh, CREATIVITY_ROW));

		cell = row.createCell(PAIR_RELATIONSHIP_CELL);
		cell.setCellValue(extractResult(sh, PAIR_RELATIONSHIP_ROW));

		cell = row.createCell(CUSTOMER_RELATIONSHIP_CELL);
		cell.setCellValue(extractResult(sh, CUSTOMER_RELATIONSHIP_ROW));

		cell = row.createCell(MANAGER_RELATIONSHIP_CELL);
		cell.setCellValue(extractResult(sh, MANAGER_RELATIONSHIP_ROW));

		cell = row.createCell(PUNCTUALITY_CELL);
		cell.setCellValue(extractResult(sh, PUNCTUALITY_ROW));

		cell = row.createCell(POSTURE_CELL);
		cell.setCellValue(extractResult(sh, POSTURE_ROW));

		cell = row.createCell(HONESTY_CELL);
		cell.setCellValue(extractResult(sh, HONESTY_ROW));

		cell = row.createCell(COMPLIANCE_CELL);
		cell.setCellValue(extractResult(sh, COMPLIANCE_ROW));
	}

	private String extractManager(Sheet sh) {
		Row row = sh.getRow(2);
		Cell cell = row.getCell(0);
		return cell.getStringCellValue().replace("GERENTE: ", "");
	}

	private String extractProfessional(Sheet sh) {
		Row row = sh.getRow(4);
		Cell cell = row.getCell(0);
		return cell.getStringCellValue().replace("NOME: ", "");
	}

	private String extractRole(Sheet sh) {
		Row row = sh.getRow(5);
		Cell cell = row.getCell(0);
		return cell.getStringCellValue().replace("CARGO / FUNÇÃO: ", "");
	}

	private String extractResult(Sheet sh, Integer rowIndex) {
		Row row = sh.getRow(rowIndex);
		return translateResult(row);
	}

	private String translateResult(Row row) {
		Cell cell = row.getCell(1);
		if (cell != null && cell.getStringCellValue().trim().toUpperCase().equals("X")) {
			return "INSATISFATÓRIO";
		}
		cell = row.getCell(2);
		if (cell != null && cell.getStringCellValue().trim().toUpperCase().equals("X")) {
			return "SATISFATÓRIO";
		}
		cell = row.getCell(3);
		if (cell != null && cell.getStringCellValue().trim().toUpperCase().equals("X")) {
			return "BOM";
		}
		cell = row.getCell(4);
		if (cell != null && cell.getStringCellValue().trim().toUpperCase().equals("X")) {
			return "EXCELENTE";
		}
		return "ERRO";
	}

	private void mountHeader(Sheet shConsolidation, int rowCount) {
		Row row = shConsolidation.createRow(rowCount);
		Cell cell = row.createCell(MANAGER_CELL);
		cell.setCellValue("GERENTE");
		cell = row.createCell(PROFESSIONAL_CELL);
		cell.setCellValue("PROFISSIONAL");
		cell = row.createCell(ROLE_CELL);
		cell.setCellValue("PERFIL");
		cell = row.createCell(TECH_KNOW_CELL);
		cell.setCellValue("DOMÍNIO TÉCNICO DA ATIVIDADE");
		cell = row.createCell(QUALITY_CELL);
		cell.setCellValue("QUALIDADE DO TRABALHO");
		cell = row.createCell(PRODUCTIVITY_CELL);
		cell.setCellValue("PRODUTIVIDADE");
		cell = row.createCell(GROWTH_CELL);
		cell.setCellValue("AUTO DESENVOLVIMENTO");
		cell = row.createCell(COMUNICATION_CELL);
		cell.setCellValue("COMUNICAÇÃO");
		cell = row.createCell(INDEPENDENCE_CELL);
		cell.setCellValue("INDEPENDÊNCIA");
		cell = row.createCell(INITIATIVE_CELL);
		cell.setCellValue("INICIATIVA");
		cell = row.createCell(TEAM_WORK_CELL);
		cell.setCellValue("TRABALHO EM GRUPO");
		cell = row.createCell(CREATIVITY_CELL);
		cell.setCellValue("CRIATIVIDADE");
		cell = row.createCell(PAIR_RELATIONSHIP_CELL);
		cell.setCellValue("RELACIONAMENTO COM PARES");
		cell = row.createCell(CUSTOMER_RELATIONSHIP_CELL);
		cell.setCellValue("RELACIONAMENTO COM CLIENTE");
		cell = row.createCell(MANAGER_RELATIONSHIP_CELL);
		cell.setCellValue("RELACIONAMENTO COM GESTOR");
		cell = row.createCell(PUNCTUALITY_CELL);
		cell.setCellValue("PONTUALIDADE");
		cell = row.createCell(POSTURE_CELL);
		cell.setCellValue("POSTURA / IMAGEM");
		cell = row.createCell(HONESTY_CELL);
		cell.setCellValue("HONESTIDADE");
		cell = row.createCell(COMPLIANCE_CELL);
		cell.setCellValue("INTEGRIDADE / COMPLIANCE");
	}

	/**
	 * List all the files under a directory
	 * 
	 * @param directoryName
	 *            to be listed
	 * @return list of absolute paths to files
	 */
	private List<String> listFiles(String directoryName) {
		List<String> list = new ArrayList<>();
		File directory = new File(directoryName);
		// get all the files from a directory
		File[] fList = directory.listFiles();
		for (File file : fList) {
			if (file.isFile()) {
				list.add(file.getAbsolutePath());
			}
		}
		return list;
	}
}
