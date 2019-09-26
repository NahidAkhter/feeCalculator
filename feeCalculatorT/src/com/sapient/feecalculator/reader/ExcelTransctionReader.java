package com.sapient.feecalculator.reader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.sapient.feecalculator.Constant.FILETYPE;
import com.sapient.feecalculator.Transaction;


public class ExcelTransctionReader extends AbstractTransactionReader implements ITransactionManager {


	@Override
	public void readTransaction(File transactionFile) {
		FileInputStream fis = null;
		ArrayList<String> list = new ArrayList<>();
		try {
			fis = new FileInputStream(transactionFile);	

			Workbook wb = WorkbookFactory.create(fis);
			Sheet sheet = wb.getSheetAt(0);

			list = new ArrayList<>();
			for (Iterator<Row> rit = sheet.rowIterator(); rit.hasNext();) {
				Row row = rit.next();

				for (Iterator<Cell> cit = row.cellIterator(); cit.hasNext();) {
					Cell cell = cit.next();
					list.add(getCellValueAsString(cell));					
				}	
				String[] transactionAttributes = new String[list.size()];
				for(int i=0;i<list.size();i++) {
					transactionAttributes[i] = list.get(i);
				}
				Transaction transaction = getTransaction(transactionAttributes); 
				saveTransaction(transaction);
				list.clear();
			}		
			
			
		} catch (FileNotFoundException e) {			
			e.printStackTrace();
		} catch (InvalidFormatException e) {		
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				fis.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	} 

	public static String getCellValueAsString(Cell cell) {
		String strCellValue = null;
		if (cell != null) {
			switch (cell.getCellType()) {
			case Cell.CELL_TYPE_STRING:
				strCellValue = cell.toString();
				break;
			case Cell.CELL_TYPE_NUMERIC:
				if (DateUtil.isCellDateFormatted(cell)) {
					SimpleDateFormat dateFormat = new SimpleDateFormat(
							"dd/MM/yyyy");
					strCellValue = dateFormat.format(cell.getDateCellValue());
				} else {
					Double value = cell.getNumericCellValue();
					Long longValue = value.longValue();
					strCellValue = new String(longValue.toString());
				}
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				strCellValue = new String(new Boolean(cell.getBooleanCellValue()).toString());
				break;
			case Cell.CELL_TYPE_BLANK:
				strCellValue = "";
				break;
			}
		}
		return strCellValue;
	}

	@Override
	public ITransactionManager readFile(FILETYPE fileType) {
		if(fileType == FILETYPE.EXCEL){
			return TrasactionReader.getTrasactionReaderInstance().readExcelFile();
		}
		return null;
	}


}
