package com.ashish.parse.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelFile {
    public static void main(String[] args) 
    {
        try {

            FileInputStream file = new FileInputStream(new File(".//test1.xls"));//DATASAMPLExlsx.xlsx"));

//          MS Office 2007+ specific  
//            XSSFWorkbook workbook = new XSSFWorkbook(file);
            Workbook workbook = WorkbookFactory.create(file);

//			MS Office 2007+ specific
//            FormulaEvaluator objFormulaEvaluator = new XSSFFormulaEvaluator((XSSFWorkbook) workbook);
            
//            https://poi.apache.org/spreadsheet/user-defined-functions.html
            
            FormulaEvaluator objFormulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
            
// HSSF works for (before or on MS Office 2007) - xls
// XSSF works for (MS Office 2007+) - xlsx
// WorkbookFactory allows POI to auto-detect based on the filetype
// http://poi.apache.org/spreadsheet/converting.html
            
//			MS Office 2007+ specific            
//            XSSFSheet sheet = workbook.getSheetAt(0);
            
            Sheet sheet = workbook.getSheetAt(0);
            DataFormatter objDefaultFormat = new DataFormatter();
            Iterator<Row> rowIterator = sheet.iterator();
            rowIterator.next();
            while(rowIterator.hasNext())
            {
                Row row = rowIterator.next();
                //For each row, iterate through each columns
                Iterator<Cell> cellIterator = row.cellIterator();
                while(cellIterator.hasNext())
                {
                    Cell cell = cellIterator.next();
                    switch(cell.getCellType()) 
                    {
	                    case Cell.CELL_TYPE_BLANK:
	                    	System.out.println("blank===>>>");
	                    	break;
	                    case Cell.CELL_TYPE_ERROR:
	                    	System.out.println("error===>>>");
	                    	break;
	                    case Cell.CELL_TYPE_FORMULA:
	                    	System.out.println("formula===>>>" + cell.getCellFormula());
	                    	break;
                        case Cell.CELL_TYPE_BOOLEAN:
                            System.out.println("boolean===>>>"+cell.getBooleanCellValue() + "\t");
                            break;
                        case Cell.CELL_TYPE_NUMERIC:
                        	
//                        	Not recommended option as per Apache Documentation
//                        	cell.setCellType(Cell.CELL_TYPE_STRING);
//                        	System.out.println("String===>>>"+cell.getStringCellValue() + "\t");

                        	System.out.println("numeric===>>>"+
                        			String.valueOf(cell.getNumericCellValue()) + "\t" + 
                        			objDefaultFormat.formatCellValue(cell, objFormulaEvaluator));
                            break;
                        case Cell.CELL_TYPE_STRING:
                            System.out.println("String===>>>"+cell.getStringCellValue() + "\t");
                            break;
                    }
                }
                System.out.println("");
            }
                        
            file.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
        	e.printStackTrace();
        }
    }

}