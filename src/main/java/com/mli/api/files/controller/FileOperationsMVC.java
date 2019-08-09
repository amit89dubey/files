package com.mli.api.files.controller;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.servlet.ModelAndView;

@Controller
public class FileOperationsMVC {

	@RequestMapping(value={"/"}, method = RequestMethod.GET)
	public ModelAndView login(){
		//processFileOperations();
		ReadWriteExcelFile.readXLSFile();
		
		ModelAndView modelAndView = new ModelAndView();
		modelAndView.setViewName("login");
		return modelAndView;
	}
	
	private void processFileOperations()
	{
		try
		{
			File myFile = new File("E:\\GITHUB_DATA\\25_series_unused_proposal_nos as on Jan. 08, 2014.xls");
	        FileInputStream fis = new FileInputStream(myFile);

	        // Finds the workbook instance for XLSX file
	        XSSFWorkbook myWorkBook = new XSSFWorkbook(fis);
	       
	        // Return first sheet from the XLSX workbook
	        XSSFSheet mySheet = myWorkBook.getSheetAt(0);
	       
	        // Get iterator to all the rows in current sheet
	        Iterator<Row> rowIterator = mySheet.iterator();
	       
	        // Traversing over each row of XLSX file
	        while (rowIterator.hasNext()) {
	            Row row = rowIterator.next();

	            //StringBuilder sb = new StringBuilder();
	            // For each row, iterate through each columns
	            Iterator<Cell> cellIterator = row.cellIterator();
	            while (cellIterator.hasNext()) {

	                Cell cell = cellIterator.next();

	                switch (cell.getCellType()) {
	                case Cell.CELL_TYPE_STRING:
	                    System.out.print(cell.getStringCellValue() + "\t");
	                	//sb.append(cell.getStringCellValue()).append("||");
	                    break;
	                case Cell.CELL_TYPE_NUMERIC:
	                    System.out.print(cell.getNumericCellValue() + "\t");
	                    //sb.append(cell.getNumericCellValue()).append("||");
	                    break;
	                case Cell.CELL_TYPE_BOOLEAN:
	                    System.out.print(cell.getBooleanCellValue() + "\t");
	                    //sb.append(cell.getBooleanCellValue()).append("||");
	                    break;
	                default :
	                	System.out.print("NA");
	                	//sb.append("AMIT");
	                }
	            }
	            System.out.println("");
	        }
		}
		catch(Exception ex)
		{
			System.out.println(ex);
		}
	}


}
