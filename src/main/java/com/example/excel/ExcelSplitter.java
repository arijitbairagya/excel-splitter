package com.example.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.DateFormatConverter;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelSplitter {

	private final String fileName;
	private int maxRows = 1000; // default to thousand
	private int wbCount = 1;
	private int totalColumns = 0;
	private List<String> headerValues = new ArrayList<String>();
	
	String excelFormatPattern = DateFormatConverter.convert(Locale.ENGLISH, "h:mm:ss AM/PM");
	
	private static Logger LOGGER = LogManager.getLogger(ExcelSplitter.class);


	public static void main(String... args) {
		LOGGER.debug("Argument Length - {}", args.length);
		if(args.length == 1) { // only file name
			
		}
		else if(args.length == 2) { // file name and chunk size
			
		}
		else {
			LOGGER.debug("Invalud Arguments Specified. Please provide file path and chunk size.");
			throw new AssertionError("Invalud Arguments Specified. Please provide file path and chunk size.");
		}
			
		LOGGER.debug("File Name : {}", args[0]);
		LOGGER.debug("Chunk size : {}", args[1]);
		/* Pass the file path and number of rows per sheet to split */
		new ExcelSplitter(String.valueOf(args[0]), Integer.parseInt(args[1]));
	}

	public ExcelSplitter(String fileName, final int maxRows) {

		this.fileName = fileName;
		this.maxRows = maxRows;

		try {
			/* Read in the original Excel file. */
			XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(fileName));
			XSSFSheet sheet = workbook.getSheetAt(0);

			/* Only split if there are more rows than the desired amount. */
			if (sheet.getPhysicalNumberOfRows() >= maxRows) {
				splitWorkbook(workbook);
			}
		}
		catch (Exception e) {
			e.printStackTrace();
		}
		LOGGER.debug("COMPLETED...");
	}


	private List<SXSSFWorkbook> splitWorkbook(XSSFWorkbook workbook) {

		SXSSFWorkbook newWb = new SXSSFWorkbook();
		Sheet newSheet = newWb.createSheet();
		

		Row newRow;
		int rowCount = 0;

		XSSFSheet sheet = workbook.getSheetAt(0);

		for (Row row : sheet) {

			// set number of cells per row as it is there in header
			if(row.getRowNum() == 0) {
				LOGGER.debug("Number of headers: "+ row.getPhysicalNumberOfCells());
				row.cellIterator().forEachRemaining( cell -> headerValues.add(cell.getStringCellValue()));
				LOGGER.debug("Headers:" + headerValues);
				totalColumns = row.getPhysicalNumberOfCells();
			}

			newRow = newSheet.createRow(rowCount++);

			// if number of rows reach the limit create another work
			if (rowCount == maxRows+1) {
				// write the workbook 
				writeWorkBooks(newWb);
				LOGGER.debug("Creating New Workbook");
				newWb = new SXSSFWorkbook();
				newSheet = newWb.createSheet();

				// create header in the first row
				copyHeader(newSheet);
				rowCount = 1;
			}

			copyCellValuebyRow(newWb, newRow, row);
		}

		/* Only add the last workbook if it has content */
		if (newWb.getSheetAt(0).getPhysicalNumberOfRows() > 0) {
			writeWorkBooks(newWb);
		}
		return null;
	}

	private void copyHeader(Sheet sheet) {
		int col =0;
		Row headerRow = sheet.createRow(0);
		for (String cellVal : headerValues) {
			headerRow.createCell(col++).setCellValue(cellVal);
		}
	}

	private void copyCellValuebyRow(SXSSFWorkbook newWb, Row newRow, Row row) {
		Cell newCell;
		
		for(int cellCount=0; cellCount < totalColumns; cellCount++) {
			newCell = newRow.createCell(cellCount);
			Cell cellVal = row.getCell(cellCount);
			setValue(newCell, cellVal);

			CellStyle newStyle = newWb.createCellStyle();
			newCell.setCellStyle(newStyle);
		}
	}

	/*
	 * Grabbing cell contents can be tricky. We first need to determine what
	 * type of cell it is.
	 */
	private void setValue(Cell newCell, Cell cell) {
		
		if(cell != null) {
			switch (cell.getCellType()) {
			case Cell.CELL_TYPE_STRING: 
				newCell.setCellValue(cell.getRichStringCellValue().getString());
				break;
			case Cell.CELL_TYPE_NUMERIC:
				if (DateUtil.isCellDateFormatted(cell)) {
					newCell.setCellValue(cell.getDateCellValue());
					newCell.setCellStyle(cell.getCellStyle());
				} else {
					newCell.setCellValue(cell.getNumericCellValue());
				}
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				newCell.setCellValue(cell.getBooleanCellValue());
				break;
			case Cell.CELL_TYPE_FORMULA:
				newCell.setCellFormula(cell.getCellFormula());
				break;
			case Cell.CELL_TYPE_BLANK:
				newCell.setCellValue("");
				break;
			default:
				LOGGER.debug("Could not determine cell type - " + cell.getCellType());
				LOGGER.debug("Could not determine cell value - " + cell.getStringCellValue());
			}
		}
	}

	/* Write all the workbooks to disk. */
	private void writeWorkBooks(SXSSFWorkbook wb) {
		FileOutputStream out = null;
		try {
			//            for (int i = 0; i < wbs.size(); i++) {
			String newFileName = fileName.substring(0, fileName.length() - 5);
			out = new FileOutputStream(new File(newFileName + "_" + (wbCount++) + ".xlsx"));
			wb.write(out);
		} catch (IOException e) {
			e.printStackTrace();
		}
		finally {
			if(out != null)
				try {
					out.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
		}
	}

}
