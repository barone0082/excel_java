package br.com.inmetrics.RegistroPonto.Apoio;

import java.util.List;
import java.util.ArrayList;

import org.apache.poi.ss.format.CellDateFormatter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class ArquivoExcel {

	private static final String FILE_NAME = "C:\\Users\\Public\\RegistroPonto\\PontoEquipe.xlsx";

	FileOutputStream out = null;
	XSSFSheet sheet = null;
	CellStyle style = null;
	XSSFWorkbook workbook = null;
	
	public void gerarPasta(String diretorio) {
		File diretorioFile = new File(diretorio);
		if (!diretorioFile.exists()) {
			diretorioFile.mkdirs();
		} 
	}
	
	public void gerarArquivoExcel(File arquivo) throws IOException {
		int posicao = arquivo.toString().lastIndexOf("\\");
		String diretorio = arquivo.toString().substring(0, posicao);
		gerarPasta(diretorio);
		if (arquivo.exists()) {
			System.out.println("Arquivo existe");
			FileInputStream inputStream = new FileInputStream(arquivo);
			workbook = new XSSFWorkbook(inputStream);
		} else {
			System.out.println("Arquivo não existe");
			workbook = new XSSFWorkbook();
		}
	}
	
	public void gerarPlanilhaExcel(String nomePlanilha) {
		if (workbook.getSheet(nomePlanilha) != null) {
			System.out.println("Planilha existe");
			sheet = workbook.getSheet(nomePlanilha);
		} else {
			System.out.println("Planilha não existe");
			sheet = workbook.createSheet(nomePlanilha);
		}
	}

	public void inserirDadosExcel(List<Ponto> pontos) {
		CellStyle headerStyle = workbook.createCellStyle();
		headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		headerStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		headerStyle.setAlignment(CellStyle.ALIGN_CENTER);
		headerStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
		CellStyle horasTotais = workbook.createCellStyle();

		int rowNum = 0;

		for (Ponto ponto : pontos) {
			List<String> textoCelulas = new ArrayList<String>();
			Row row = sheet.getRow(rowNum);
			if (row == null) {
				row = sheet.createRow(rowNum);
			}
			//Row row = sheet.createRow(rowNum);
			if (rowNum == 0) {
				textoCelulas.add("Data");
				textoCelulas.add("Ponto 1");
				textoCelulas.add("Ponto 2");
				textoCelulas.add("Ponto 3");
				textoCelulas.add("Ponto 4");
				textoCelulas.add("Ponto 5");
				textoCelulas.add("Ponto 6");
				textoCelulas.add("Cálculos");
				textoCelulas.add("Obs");
				textoCelulas.add(" ");
				textoCelulas.add("Total");
			} else {
				if(ponto.getData().contains("Total")) {
					break;
				}
				textoCelulas.add(ponto.getData());
				for (Horario horario : ponto.getMarcacao()) {						
					textoCelulas.add(horario.getHorario());
				}
				textoCelulas.add(ponto.getCalculos());
				textoCelulas.add(ponto.getObservacoes());
			}
			int numColuna = 0;
			for (String textoCelula : textoCelulas) {
				Cell cell = row.createCell(numColuna);
				if (rowNum == 0) {
					cell.setCellStyle(headerStyle);
				}
				cell.setCellValue(textoCelula);
				numColuna = numColuna + 1;
			}
			 rowNum = rowNum + 1;
			 if (!(rowNum == 1)) {
			 //Cell cell = row.createCell(numColuna + 1);
			 Cell cell = row.getCell(numColuna + 1);
			 if (cell == null) {
				 cell = row.createCell(numColuna + 1);
			 }
			 //cell.setCellFormula("C" + rowNum + " - B" + rowNum + " + E" + rowNum + " - D" + rowNum + "");
			 //cell.setCellFormula("SE(E(C" + rowNum + "<>\" \";C" + rowNum + "<>\" \");(C" + rowNum + "-B" + rowNum + "); 0) + SE(E(E" + rowNum + "<>\" \";D" + rowNum + "<>\" \");(E" + rowNum + "-D" + rowNum + "); 0)");
			 String formula = "(SE(E(C" + rowNum + "<> N3 , C" + rowNum + "<>N3),(C" + rowNum + "-B" + rowNum + "), 0) + SE(E(E" + rowNum + "<>N3,D" + rowNum + "<>N3),(E" + rowNum + "-D" + rowNum + "), 0) + SE(E(G" + rowNum + "<>N3,F" + rowNum + "<>N3),(G" + rowNum + "-F" + rowNum + "), 0))";
			 String formula2 = "SEERRO(O2-N2,0)";
			 String formula3 = "C" + rowNum + " - B" + rowNum + " + E" + rowNum + " - D" + rowNum + "";
			 cell.setCellType(CellType.FORMULA);
			 cell.setAsActiveCell();
			 cell.setCellFormula(formula2);
			
			 }
			 
		}
		
	}
	
	public void salvarArquivoExcel(File arquivo) {
		try {
			//triggerFormula(workbook);
			workbook.setForceFormulaRecalculation(true);
			FileOutputStream outputStream = new FileOutputStream(arquivo);
			workbook.write(outputStream);
			workbook.close();
			
			FileInputStream inputStream = new FileInputStream(arquivo);
			workbook = new XSSFWorkbook(inputStream);
			XSSFFormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
			evaluator.evaluateAll();
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	public static void triggerFormula(XSSFWorkbook workbook){      

		XSSFFormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        XSSFSheet sheet = workbook.getSheetAt(0);
        int lastRowNo=sheet.getLastRowNum();        

        for(int rownum=0;rownum<=lastRowNo;rownum++){
        Row row;
         if (sheet.getRow(rownum)!=null){
                 row= sheet.getRow(rownum);

              int lastCellNo=row.getLastCellNum();

                  for(int cellnum=0;cellnum<lastCellNo;cellnum++){  
                          Cell cell;
                          if(row.getCell(cellnum)!=null){
                             cell = row.getCell(cellnum);   
                             if(CellType.FORMULA == cell.getCellTypeEnum()) {
                            	 evaluator.evaluateFormulaCell(cell);
                             }
                            if(Cell.CELL_TYPE_FORMULA==cell.getCellType()){
                            evaluator.evaluateFormulaCell(cell);
                        }
                    }
                 }
         }
        }


    }
	public void gerarArquivoExcel(List<Ponto> pontos, String nomeAnalista) throws IOException {
		// Criando o arquivo e uma planilha chamada "Product"

		File file = new File("C:\\Users\\Public\\RegistroPonto\\PontoEquipe.xlsx");

		int posicao = file.toString().lastIndexOf("\\");
		String caminho = file.toString().substring(0, posicao);
		System.out.println(caminho);
		
		// verifica se o arquivo existe
		if (file.exists()) {
			System.out.println("Arquivo existe");
			FileInputStream inputStream = new FileInputStream(file);
			workbook = new XSSFWorkbook(inputStream);
			// vefirifica se planilha existe
			if (workbook.getSheet(nomeAnalista) != null) {
				System.out.println("Planilha existe");
				sheet = workbook.getSheet(nomeAnalista);
			} else {
				System.out.println("Planilha não existe");
				sheet = workbook.createSheet(nomeAnalista);
				// addIntoCell(excelData, excelSheet, "0", style);
			}
		} else {
			System.out.println("Arquivo nãoexiste");
			workbook = new XSSFWorkbook();
			sheet = workbook.createSheet(nomeAnalista);
		}

		// XSSFWorkbook workbook = new XSSFWorkbook();
		// XSSFSheet sheet = workbook.createSheet(nomeAnalista);

		CellStyle headerStyle = workbook.createCellStyle();
		headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		headerStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		headerStyle.setAlignment(CellStyle.ALIGN_CENTER);
		headerStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);

		CellStyle horasTotais = workbook.createCellStyle();

		Object[][] datatypes = { { "Datatype", "Type", "Size(in bytes)" }, { "int", "Primitive", 2 },
				{ "float", "Primitive", 4 }, { "double", "Primitive", 8 }, { "char", "Primitive", 1 },
				{ "String", "Non-Primitive", "No fixed size" } };

		int rowNum = 0;

		for (Ponto ponto : pontos) {
			List<String> textoCelulas = new ArrayList<String>();
			Row row = sheet.createRow(rowNum++);
			if (rowNum == 1) {
				textoCelulas.add("Data");
				textoCelulas.add("Ponto 1");
				textoCelulas.add("Ponto 2");
				textoCelulas.add("Ponto 3");
				textoCelulas.add("Ponto 4");
				textoCelulas.add("Ponto 5");
				textoCelulas.add("Ponto 6");
				textoCelulas.add("Cálculos");
				textoCelulas.add("Obs");
				textoCelulas.add(" ");
				textoCelulas.add("Total");
			} else {
				textoCelulas.add(ponto.getData());
				for (Horario horario : ponto.getMarcacao()) {
					textoCelulas.add(horario.getHorario());
					System.out.println(horario.getHorario());
				}
				textoCelulas.add(ponto.getCalculos());
				textoCelulas.add(ponto.getObservacoes());
			}
			int numColuna = 0;
			for (String textoCelula : textoCelulas) {
				Cell cell = row.createCell(numColuna++);
				if (rowNum == 1) {
					cell.setCellStyle(headerStyle);
				}
				cell.setCellValue(textoCelula);
			}
			// if (!(rowNum == 1)) {
			// Cell cell = row.createCell(numColuna + 1);
			// cell.setCellFormula("B2 +C2");
			// cell.setCellFormula("SE(E(C" + rowNum + "<>\" \";C" + rowNum + "<>\" \");(C"
			// + rowNum + "-B" + rowNum + "); 0) + SE(E(E" + rowNum + "<>\" \";D" + rowNum +
			// "<>\" \");(E" + rowNum + "-D" + rowNum + "); 0)");
			// }
		}
		try {
			FileOutputStream outputStream = new FileOutputStream(FILE_NAME);
			workbook.write(outputStream);
			workbook.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

		System.out.println("Done");

		/*
		 * Workbook workbook = new HSSFWorkbook(); HSSFSheet sheet =
		 * workbook.createSheet("Product"); // Definindo alguns padroes de layout
		 * sheet.setDefaultColumnWidth(15); sheet.setDefaultRowHeight((short)400);
		 * 
		 * //Carregando os produtos List products = getProducts();
		 * 
		 * int rownum = 0; int cellnum = 0; Cell cell; Row row;
		 * 
		 * //Configurando estilos de células (Cores, alinhamento, formatação, etc..)
		 * HSSFDataFormat numberFormat = workbook.createDataFormat();
		 * 
		 * CellStyle headerStyle = workbook.createCellStyle();
		 * headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		 * headerStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		 * headerStyle.setAlignment(CellStyle.ALIGN_CENTER);
		 * headerStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		 * 
		 * CellStyle textStyle = workbook.createCellStyle();
		 * textStyle.setAlignment(CellStyle.ALIGN_CENTER);
		 * textStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		 * 
		 * CellStyle numberStyle = workbook.createCellStyle();
		 * numberStyle.setDataFormat(numberFormat.getFormat(“#,##0.00”));
		 * numberStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		 * 
		 * // Configurando Header row = sheet.createRow(rownum++); cell =
		 * row.createCell(cellnum++); cell.setCellStyle(headerStyle);
		 * cell.setCellValue(“Code”);
		 * 
		 * cell = row.createCell(cellnum++); cell.setCellStyle(headerStyle);
		 * cell.setCellValue(“Name”);
		 * 
		 * cell = row.createCell(cellnum++); cell.setCellStyle(headerStyle);
		 * cell.setCellValue(“Price”);
		 * 
		 * // Adicionando os dados dos produtos na planilha for (Product product :
		 * products) { row = sheet.createRow(rownum++); cellnum = 0;
		 * 
		 * cell = row.createCell(cellnum++); cell.setCellStyle(textStyle);
		 * cell.setCellValue(product.getId());
		 * 
		 * cell = row.createCell(cellnum++); cell.setCellStyle(textStyle);
		 * cell.setCellValue(product.getName());
		 * 
		 * cell = row.createCell(cellnum++); cell.setCellStyle(numberStyle);
		 * cell.setCellValue(product.getPrice());
		 */
	}

}
