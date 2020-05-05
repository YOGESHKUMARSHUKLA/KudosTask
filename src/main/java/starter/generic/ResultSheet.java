package starter.generic;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.xssf.usermodel.*;

import java.awt.Color;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;


public class ResultSheet {
	public  FileInputStream fis;
	 public  FileOutputStream fos;
	 public  XSSFWorkbook wb;
	 public  XSSFSheet sh;
	 private  XSSFCell cell;
	 private  XSSFRow row;
	 private  XSSFCellStyle cellstyle;
	 private  XSSFColor mycolor;
	 private Logger Log=null;
	 

	public void setLog(Logger log) {
		this.Log = log;
	}


	private  XSSFColor mycolorMatch;
	 private  XSSFColor mycolorMismatch;
	 
	 private  String ExcelPath =null;
	 public  File wbFileObj = null;
	 public String wbFullpath=null;
	 
	//Create report workbook and TC sheet
	/***
	 * Pass to constructor 
	 * - Path where workbook is to be saved, 
	 * - Name of workbook with extension e.g. xlsx
	 * - Name of Initial Sheet to create
	 * 
	 * @param reportPath
	 * @param wbName
	 * @param wsName
	 */
/*	public ResultSheet(String reportPath, String wbName, String wsName ) {		

		try{
			this.wbFullpath = reportPath+wbName;
			// Path pathAbsolute = Paths.get("wbFullpath");
			this.wbFileObj = new File(wbFullpath);
		       File f = new File(reportPath+wbName);
			if (!f.exists()) {
				//f.createNewFile();
				//this.ExcelPath = f.getAbsolutePath();
				
				wb = new XSSFWorkbook();
				sh = wb.createSheet("Sheet1");
				fos = new FileOutputStream(wbFullpath);
	            wb.write(fos);
	            wb.close();
	            
				fis=new FileInputStream(wbFullpath);
		        wb=new XSSFWorkbook(fis);
		        sh = wb.getSheet(wsName);
		        //sh = wb.getSheetAt(0); //0 - index of 1st sheet
		        if (sh == null) {
		        	this.sh = wb.createSheet(wsName);
		        }
		        fis.close();
			
				Log.info("File doesn't exist, so created!");
				
			} else {
				f.delete();
				//f.createNewFile();
				this.ExcelPath = f.getAbsolutePath();
				
				wb = new XSSFWorkbook();
				sh = wb.createSheet("Sheet1");
				FileOutputStream outputStream = new FileOutputStream(wbFullpath);
	            wb.write(outputStream);
	            wb.close();
	            
				fos = new FileOutputStream(wbFullpath);
	            wb.write(fos);
	            wb.close();
	            
				fis=new FileInputStream(wbFullpath);
		        wb=new XSSFWorkbook(fis);
		        sh = wb.getSheet(wsName);
		        //sh = wb.getSheetAt(0); //0 - index of 1st sheet
		        if (sh == null) {
		        	this.sh = wb.createSheet(wsName);
		        }
		        //fis.close();
				
				Log.info("File already exists, overwritten !");
			}
		        
		} catch (Exception e) {
			e.printStackTrace();
		}

	}
	
	*//***
	 * Pass to constructor 
	 * - Path where workbook is to be saved, 
	 * - Name of workbook with extension e.g. xlsx
	 * **Dummy Sheet (Sheet1) will be added to WB
	 * @param reportPath
	 * @param wbName
	 * @param wsName
	 *//*
	public ResultSheet(String reportPath, String wbName ) {		
		
		try{
			this.wbFullpath = reportPath+wbName;
			this.wbFileObj = new File(wbFullpath);
			Log.info("Result Work book location: wbFileObj= "+wbFileObj.getAbsolutePath());
			Log.info("wbFullpath= "+wbFullpath);
			//Log.info("Result sheet full path : "+wbFullpath);
		       File f = new File(reportPath+wbName);
			if (!f.exists()) {
				XSSFWorkbook wbTmp = new XSSFWorkbook();
				wbTmp.createSheet("Sheet1");
				FileOutputStream fileOut = new FileOutputStream(f);
				wbTmp.write(fileOut);
				fileOut.close();
				
				this.ExcelPath = f.getAbsolutePath();
				Log.info("File doesn't exist, so created!");
			} else {
				f.delete();
				//Create new
				XSSFWorkbook wbTmp = new XSSFWorkbook();
				wbTmp.createSheet("Sheet1");
				FileOutputStream fileOut = new FileOutputStream(f);
				wbTmp.write(fileOut);
				fileOut.close();
				
				this.ExcelPath = f.getAbsolutePath();
				Log.info("File already exists, overwritten !");
			}
		        fis=new FileInputStream(wbFileObj);
		         wb=new XSSFWorkbook(fis);
		         wb.getSheetAt(0); 
		         Log.info("wb 0th sheet name = "+wb.getSheetAt(0).getSheetName());
		         //0 - index of 1st sheet
				//sh = wb.createSheet("Sheet1");
		         fis.close();
		       
		       if (!f.exists()) {
					//f.createNewFile();
					//this.ExcelPath = f.getAbsolutePath();
					
					wb = new XSSFWorkbook();
					sh = wb.createSheet("Sheet1");
					FileOutputStream outputStream = new FileOutputStream(wbFullpath);
		            wb.write(outputStream);
		            wb.close();
		            
		            
					fis=new FileInputStream(wbFullpath);
			        wb=new XSSFWorkbook(fis);
			        sh = wb.getSheetAt(0);
			        if (sh == null) {
			        	Log.info("New sheet with name sheet 1 created");
			        	this.sh = wb.createSheet("Sheet1");
			        }
			       // fis.close();
				
					Log.info("File doesn't exist, so created!");
					
				} else {
					f.delete();
					//f.createNewFile();
					this.ExcelPath = f.getAbsolutePath();
					
					wb = new XSSFWorkbook();
					sh = wb.createSheet("Sheet1");
					FileOutputStream outputStream = new FileOutputStream(wbFullpath);
		            wb.write(outputStream);
		            wb.close();
		            
					fis=new FileInputStream(wbFullpath);
			        wb=new XSSFWorkbook(fis);
			        sh = wb.getSheetAt(0);
			        if (sh == null) {
			        	Log.info("New sheet with name sheet 1 created");
			        	this.sh = wb.createSheet("Sheet1");
			        }
				}
		} catch (Exception e) {
			e.printStackTrace();
		}


	}*/
	
	/***
	 * Pass to constructor 
	 * - Name of workbook with extension e.g. xlsx
	 * **Dummy Sheet (Sheet1) will be added to WB
	 * @param wbName
	 */
	public ResultSheet(String wbName ) {		
		
		try{
			this.wbFullpath = wbName;
			this.wbFileObj = new File(wbFullpath);
			System.out.println("wbFullpath= "+wbFullpath);
		       File f = new File(wbFullpath);
		     /*  
		       if (!f.exists()) {
					//f.createNewFile();
					//this.ExcelPath = f.getAbsolutePath();
		    	   
		    	   wb = new XSSFWorkbook();
					fos = new FileOutputStream(f);
					wb.write(fos);
					fos.close();
					
					
		    	   	fis = new FileInputStream(f);
					wb = new XSSFWorkbook(fis);
					sh = wb.createSheet("Sheet1");
					fis.close();
					fos = new FileOutputStream(f);
		            wb.write(fos);
		            fos.close();
		          //  wb.close();
		            
					Log.info("File doesn't exist, so created!");
					
				} 
 */
					//Replace delete with Archive functionality in future versions
					f.delete();
					
					//Create new sheet on drive
					wb = new XSSFWorkbook();
					fos = new FileOutputStream(f);
					wb.write(fos);
					fos.close();
					
					this.ExcelPath = f.getAbsolutePath();
					
					//Read file and add new dummy sheet
					fis = new FileInputStream(f);
					wb = new XSSFWorkbook(fis);
					sh = wb.createSheet("Sheet1");
					fis.close();
					fos = new FileOutputStream(f);
		            wb.write(fos);
		            fos.close();
		           // wb.close();
		            
		            System.out.println("Result Sheet created for Suite : "+f.getAbsolutePath());
				
		} catch (Exception e) {
			e.printStackTrace();
		}


	}
	
	
	
	public String createNewTCSheet(String wsName) throws IOException {
		loadWorkBookInWriteMode();
		Log.info("Inside create new sheet fn");
		sh = wb.getSheet(wsName);
		if (sh == null) {
			Log.info("Creating new sheet in wb with name "+wsName);
			sh = wb.createSheet(wsName);
		}
		saveReportOnDrive();
		return wsName;
		
	}
	
	private void loadWorkBookInWriteMode() throws IOException {
		Log.info("Loading file "+wbFullpath+" from drive ");
		File f = new File(this.wbFullpath);
		
		fis = new FileInputStream(f);
		wb = new XSSFWorkbook(fis);
		sh = wb.getSheetAt(0);
		Log.info("File loaded succesfully");
	}
	
	private void loadWorkBook() throws IOException {
		Log.info("Loading file "+wbFullpath+" from drive ");
		File f = new File(this.wbFullpath);
		
		fis = new FileInputStream(f);
		wb = new XSSFWorkbook(fis);
		sh = wb.getSheetAt(0);
		Log.info("File loaded succesfully");
	}
	
	public void saveReportOnDrive() throws IOException{
		Log.info("Trying to save file "+wbFullpath+" to drive");
		File f = new File(this.wbFullpath);
		if(!(fis == null))
			fis.close();
		fos = new FileOutputStream(f);
        wb.write(fos);
        fos.close();
        Log.info("File saved successfully");
	}

	public String createNewSheetToSaveTCResults(String TestCaseName, int runCount) throws IOException {
		loadWorkBookInWriteMode();
		
		String tcNameInWorkBook = "";
		if (runCount == 1) {
			tcNameInWorkBook=TestCaseName;
		} else {
			tcNameInWorkBook=TestCaseName + "_" + runCount;
		}
		
		sh = wb.createSheet(tcNameInWorkBook);
		
		saveReportOnDrive();
		return tcNameInWorkBook;
		
	}
	
	
	/**
	 * Add blank line in sheet
	 * 
	 * @param text
	 * @param rownum
	 * @param colnum
	 * @throws Exception
	 */
	public void addBlankRow(String wsname,int rownum) throws Exception {
		
		loadWorkBookInWriteMode();
		sh=wb.getSheet(wsname);
		Log.info("Adding blank row - in sheet-"+sh.getSheetName());
		Log.info("ROW = "+rownum);
		try {
			try{
			row = sh.getRow(rownum);
			Log.info("ROW got - "+row.getRowNum());
			}catch (Exception e) {
				Log.info("EXCEPTION Create row - ");
				row = sh.createRow(rownum);
				Log.info("NEW Create row #- "+row.getRowNum());
			}
			if (row == null) {
				Log.info("Row is still null creating again");
				row = sh.createRow(rownum);
			}
			
			try{
				cell = row.getCell(0);
				cell.setCellValue(" ");
				Log.info("Cell value set - ");
				}catch (Exception e) {
					Log.info("EXCEPTION Create Cell - ");
					cell = row.createCell(0);
					cell.setCellValue(" ");
				}
			
/*			if (cell == null) {
				Log.info("EXCEPTION2 Create Cell - ");
				cell = row.createCell(0);
				cell.setCellValue(" ");
			} */

			saveReportOnDrive();
			
		} catch (Exception e){
			saveReportOnDrive();
			throw (e);
		}
	}
	
	/**
	 * @param text
	 * @param rownum
	 * @param colnum
	 * @throws Exception
	 */
	public void setCellData(String wsName, String text, int rownum, int colnum) throws Exception
	{
		loadWorkBookInWriteMode();
		
		sh = wb.getSheet(wsName);
		Log.info("Adding new cell value to sheet " + sh.getSheetName());
		Log.info("ROW =  " + rownum);
		
		try{   
		 try{
			 row  = sh.getRow((short)rownum);
			 Log.info("Row retrived successfully to write cell value : "+row.getRowNum());
		    }catch(NullPointerException NPE)
		    {
		    	Log.info("Row value is null - creating new row : "+rownum);
		    	row = sh.createRow(rownum);
		    	//row  = sh.getRow(rownum);
		    }
		 
			try {
				cell = row.getCell(colnum);
				cell.setCellValue(text);
				Log.info("Cell value set successfully ");
			} catch (NullPointerException NPE) {
				Log.info("Cell valule is null - trying to create new cell and set value ");
				cell = row.createCell(colnum);
				//cell = row.getCell(colnum);
				Log.info("Adding value to ROW = "+row.getRowNum()+"----- Value="+text);
				cell.setCellValue(text);
				
			}

			saveReportOnDrive();
	 }catch(Exception e)
	 {
		 Log.info("Error setting value to Cell !!!");
		 throw (e);
	 } 
	 }
	
	public void setValidationStatus(String wsName, String text, int rownum, int colnum, String color) throws Exception
	{
		loadWorkBookInWriteMode();
		
		sh = wb.getSheet(wsName);
		Log.info("Adding new cell value to sheet " + sh.getSheetName());
		Log.info("ROW =  " + rownum);
		
		
		//Set cell style
	    XSSFCellStyle resultMsgStyle = wb.createCellStyle();
        Font font = wb.createFont();
        
		if(color.toUpperCase().equals("GREEN")){
			font.setColor(IndexedColors.GREEN.getIndex());
		}
		if(color.toUpperCase().equals("RED")){
			font.setColor(IndexedColors.RED.getIndex());
		}
			
		font.setBold(true);
		resultMsgStyle.setFont(font);

      
		
		try{   
		 try{
			 row  = sh.getRow((short)rownum);
			 Log.info("Row retrived successfully to write cell value : "+row.getRowNum());
		    }catch(NullPointerException NPE)
		    {
		    	Log.info("Row value is null - creating new row : "+rownum);
		    	row = sh.createRow(rownum);
		    	
		    }
		 
			try {
				cell = row.getCell(colnum);
				cell.setCellValue(text);
				cell.setCellStyle(resultMsgStyle);
				Log.info("Cell value set successfully ");
			} catch (NullPointerException NPE) {
				Log.info("Cell valule is null - trying to create new cell and set value ");
				cell = row.createCell(colnum);
				//cell = row.getCell(colnum);
				Log.info("Adding value to ROW = "+row.getRowNum()+"----- Value="+text);
				cell.setCellValue(text);
				cell.setCellStyle(resultMsgStyle);
				
			}

			saveReportOnDrive();
	 }catch(Exception e)
	 {
		 Log.info("Error setting value to Cell !!!");
		 throw (e);
	 } 
	 }
	
	//To write result In test data and test case list sheet.
		public boolean writeScrappedValuesToSheet(String wsName,ArrayList<String> Result, int rowNum) throws IOException{
			
			loadWorkBookInWriteMode();
			
			sh = wb.getSheet(wsName);
			XSSFCellStyle style = wb.createCellStyle();
			style.setBorderBottom(XSSFCellStyle.BORDER_THIN);
			style.setBorderTop(XSSFCellStyle.BORDER_THIN);
			style.setBorderRight(XSSFCellStyle.BORDER_THIN);
			style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
			
			try{
				int sheetIndex=wb.getSheetIndex(wsName);
				if(sheetIndex==-1)
					return false;			
				//int colNum = retrieveNoOfCols(wsName);
				//int rowNum = retrieveNoOfRows(wsName);
				XSSFRow row = sh.createRow(rowNum);
				
				Log.info("Writing Scrapped Values to Sheet");
				int colNo = 0;
				for (String element : Result) {
					//Log.info(element);
					XSSFCell cell = row.createCell(colNo++);
					cell.setCellStyle(style);
					cell.setCellValue(element);
				}
				
			}catch(Exception e){
				e.printStackTrace();
				Log.info("Error writing scrapped values to Result Sheet.");
				saveReportOnDrive();
				return false;
			}
			Log.info("Scrapped values successfully written to Result Sheet.");
			saveReportOnDrive();
			return true;
		}
		
		/**
		 * Save result sheet to file
		 * 
		 * @param prop
		 * @return
		 * @throws IOException
		 */
		public String saveReport() throws IOException {
			
			 saveReportOnDrive();
		     return this.wbFullpath;
			 
		}
		
		
	/***
	 * 
	 * @param rownum
	 * @param colnum
	 * @return
	 * @throws Exception
	 */
	public String getCellData(String wsName, int rownum, int colnum) throws Exception
	 {
		loadWorkBook();
		sh = wb.getSheet(wsName);
	  try{
	     cell = sh.getRow(rownum).getCell(colnum);
	     String CellData = null;         
	     switch (cell.getCellType()){
	     case Cell.CELL_TYPE_STRING:
	          CellData = cell.getStringCellValue();
	          break;
	     case Cell.CELL_TYPE_NUMERIC:
	          if (DateUtil.isCellDateFormatted(cell))
	          {
	             CellData = cell.getDateCellValue().toString();
	          }
	          else
	          {  
	             CellData = Double.toString(cell.getNumericCellValue());
	             if (CellData.contains(".0"))//removing the extra .0
	              {
	               CellData = CellData.substring(0, CellData.length()-2);
	              }
	           }
	           break;
	     case Cell.CELL_TYPE_BLANK:
	           CellData = "";
	           break;
	     case Cell.CELL_TYPE_BOOLEAN:
	           CellData = Boolean.toString(cell.getBooleanCellValue());
	           break;
	     }      
	        return CellData;
	        }catch (Exception e){return"";}
	 }
	 
	/***
	 * 
	 * @param color
	 * @param rownum
	 * @param colnum
	 * @throws Exception
	 */
	public void highlightCell(String wsName, String color, int rownum, int colnum)throws Exception
	 {
		loadWorkBookInWriteMode();
		sh =wb.getSheet(wsName);
		
	  try{
	//  cell = sh.getRow(rownum).getCell(colnum,Row.RETURN_BLANK_AS_NULL);
	  cell = sh.getRow(rownum).getCell(colnum, null);
	   }catch(Exception e){Log.info("cell is null");}
	  if (cell == null) 
	    {
	   cell = row.createCell(colnum);
	   cell = row.getCell(colnum);
	    }
	  cellstyle = wb.createCellStyle();
	  color = color.toUpperCase();
	  
	  switch(color)
	  {
	  case "GREEN":
	   mycolor = new XSSFColor(Color.GREEN);
	   break;
	  case "RED":
	   mycolor = new XSSFColor(Color.RED);
	  break;
	  default:
	   mycolor = new XSSFColor(Color.BLACK);
	  break;
	  }
	  cellstyle.setFillForegroundColor(mycolor); 
	  cellstyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
	  cell.setCellStyle(cellstyle);
	  
	  saveReportOnDrive();
	  System.out.print("Cell Highlighting done");  
	 }
	
	//To retrieve No Of Rows from excel file
		public int retrieveNoOfRows(String wsName){		
			int sheetIndex=wb.getSheetIndex(wsName);
			if(sheetIndex==-1)
				return 0;
			else{
				sh = wb.getSheetAt(sheetIndex);
				int rowCount=sh.getLastRowNum()+1;		
				return rowCount;		
			} 
		}
		
		//To retrieve No Of Columns
		public int retrieveNoOfCols(String wsName){
			int sheetIndex=wb.getSheetIndex(wsName);
			if(sheetIndex==-1)
				return 0;
			else{
				sh = wb.getSheetAt(sheetIndex);
				int colCount=sh.getRow(0).getLastCellNum();			
				return colCount;
			}
		}
		
		//To retrieve test data from test case data sheets.
		public Object[][] retrieveTestData(String wsName){
			int sheetIndex=wb.getSheetIndex(wsName);
			if(sheetIndex==-1){
				Log.info("Inside retrieveTestData - no sheet found");
				return null;
			}
			else{
					int rowNum = retrieveNoOfRows(wsName);
					int colNum = retrieveNoOfCols(wsName);
			
					Object data[][] = new Object[rowNum-1][colNum-2];
					//Log.info("rowN: "+rowNum);
					for (int i=0; i<rowNum-1; i++){
						XSSFRow row = sh.getRow(i+1);
						
						for(int j=0; j< colNum-2; j++){					
							if(row==null){
								data[i][j] = "";
							}
							else{
								XSSFCell cell = row.getCell(j);	
						
								if(cell==null){
									data[i][j] = "";							
								}
								else{
									cell.setCellType(Cell.CELL_TYPE_STRING);
									String value = cellToString(cell);
									data[i][j] = value;						
								}
							}
						}				
					}			
					return data;		
			}
		
		}	
		
		public  String cellToString(XSSFCell cell){
			int type;
			Object result;
			type = cell.getCellType();			
			switch (type){
				case 0 :
					result = cell.getNumericCellValue();
					break;
					
				case 1 : 
					result = cell.getStringCellValue();
					break;
					
				default :
					throw new RuntimeException("Unsupportd cell.");			
			}
			return result.toString();
		}
		
		/***
		 * Will save workbook to drive  - WILL NOT CLOSE the workBook Obj
		 * @throws IOException
		 */
		public void saveWorkbook() throws IOException {
			if(!(fis==null))
				fis.close();
			 fos = new FileOutputStream(this.wbFullpath);
			 wb.write(fos);
			 fos.flush();
			 fos.close();
		}
		
		/**
		 * Save and close workbook & input/output streams
		 * @throws IOException
		 */
		public void closeWorkBook() throws IOException {
			if(!(fis==null))
				fis.close();
			 fos = new FileOutputStream(this.wbFullpath);
			 wb.write(fos);
			 fos.flush();
			 fos.close();
			 wb.close();
		}
		
		
	/**
	 * Copy and append the row from source sheet to end of Sheet
	 * 
	 * @param wbSrc
	 * @param wsSrc
	 * @param rowSrc
	 * @throws Exception
	 */
	public  void copyRowFromSheet(XSSFWorkbook wbSrc, XSSFSheet wsSrc, int rowSrc) throws Exception {
		if (wbSrc == null) {
			throw new Exception("Workbook/Sheet is null");
		}
		int sheetIndex = wbSrc.getSheetIndex(wsSrc.getSheetName());
		if (sheetIndex == -1) {
			throw new Exception("Sheet not found");
		}

		XSSFRow newRow = sh.getRow(sh.getLastRowNum());
		XSSFRow sourceRow = wsSrc.getRow(rowSrc);

		// If the row exist in destination, push down all rows by 1 else create
		// a new row
		if (newRow == null) {
			newRow = sh.createRow(sh.getLastRowNum());
		}

		// Loop through source columns to add to new row
		for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
			// Grab a copy of the old/new cell
			XSSFCell oldCell = sourceRow.getCell(i);
			XSSFCell newCell = newRow.createCell(i);

			// If the old cell is null jump to next cell
			if (oldCell == null) {
				newCell = null;
				continue;
			}

			// Copy style from old cell and apply to new cell
			XSSFCellStyle newCellStyle = wb.createCellStyle();
			newCellStyle.cloneStyleFrom(oldCell.getCellStyle());

			newCell.setCellStyle(newCellStyle);

			// If there is a cell comment, copy
			if (oldCell.getCellComment() != null) {
				newCell.setCellComment(oldCell.getCellComment());
			}

			// If there is a cell hyperlink, copy
			if (oldCell.getHyperlink() != null) {
				newCell.setHyperlink(oldCell.getHyperlink());
			}

			// Set the cell data type
			newCell.setCellType(oldCell.getCellType());

			// Set the cell data value
			switch (oldCell.getCellType()) {
			case Cell.CELL_TYPE_BLANK:
				newCell.setCellValue(oldCell.getStringCellValue());
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				newCell.setCellValue(oldCell.getBooleanCellValue());
				break;
			case Cell.CELL_TYPE_ERROR:
				newCell.setCellErrorValue(oldCell.getErrorCellValue());
				break;
			case Cell.CELL_TYPE_FORMULA:
				newCell.setCellFormula(oldCell.getCellFormula());
				break;
			case Cell.CELL_TYPE_NUMERIC:
				newCell.setCellValue(oldCell.getNumericCellValue());
				break;
			case Cell.CELL_TYPE_STRING:
				newCell.setCellValue(oldCell.getRichStringCellValue());
				break;
			}

		}
		saveReportOnDrive();
	}
		
	
	
	/**
	 * Copy and paste the row from source sheet to destination
	 * to a given row number
	 * 
	 * @param wbSrc
	 * @param wsSrc
	 * @param rowSrc
	 * @throws Exception
	 */
	public  void copyRowFromSheet(XSSFWorkbook wbSrc, XSSFSheet wsSrc, int rowSrc, int rowDest) throws Exception {
		if (wbSrc == null) {
			throw new Exception("Workbook/Sheet is null");
		}
		int sheetIndex = wbSrc.getSheetIndex(wsSrc.getSheetName());
		if (sheetIndex == -1) {
			throw new Exception("Sheet not found");
		}

		XSSFRow newRow = sh.getRow(rowDest);
		XSSFRow sourceRow = wsSrc.getRow(rowSrc);

		// If the row exist in destination, push down all rows by 1 else create
		// a new row
		if (newRow == null) {
			newRow = sh.createRow(sh.getLastRowNum());
		}

		// Loop through source columns to add to new row
		for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
			// Grab a copy of the old/new cell
			XSSFCell oldCell = sourceRow.getCell(i);
			XSSFCell newCell = newRow.createCell(i);

			// If the old cell is null jump to next cell
			if (oldCell == null) {
				newCell = null;
				continue;
			}

			// Copy style from old cell and apply to new cell
			XSSFCellStyle newCellStyle = wb.createCellStyle();
			newCellStyle.cloneStyleFrom(oldCell.getCellStyle());

			newCell.setCellStyle(newCellStyle);

			// If there is a cell comment, copy
			if (oldCell.getCellComment() != null) {
				newCell.setCellComment(oldCell.getCellComment());
			}

			// If there is a cell hyperlink, copy
			if (oldCell.getHyperlink() != null) {
				newCell.setHyperlink(oldCell.getHyperlink());
			}

			// Set the cell data type
			newCell.setCellType(oldCell.getCellType());

			// Set the cell data value
			switch (oldCell.getCellType()) {
			case Cell.CELL_TYPE_BLANK:
				newCell.setCellValue(oldCell.getStringCellValue());
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				newCell.setCellValue(oldCell.getBooleanCellValue());
				break;
			case Cell.CELL_TYPE_ERROR:
				newCell.setCellErrorValue(oldCell.getErrorCellValue());
				break;
			case Cell.CELL_TYPE_FORMULA:
				newCell.setCellFormula(oldCell.getCellFormula());
				break;
			case Cell.CELL_TYPE_NUMERIC:
				newCell.setCellValue(oldCell.getNumericCellValue());
				break;
			case Cell.CELL_TYPE_STRING:
				newCell.setCellValue(oldCell.getRichStringCellValue());
				break;
			}

		}
		saveReportOnDrive();
	}
		
	
/**
 * Copy specific row from source sheet
 * from specific start column index to end index
 * 
 * 
 * @param wbSrc  - Copy from this workbook
 * @param wsSrc  - Copy from this specific sheet
 * @param rowSrc - Copy from this specific row from source sheet
 * @param rowDest- Copy to this specific row in destination sheet 
 * @param colBegSrc - Copy from this column index from src sheet ( 0 to n-1 )
 * @param colEndSrc - Copy till this specific index in src sheet (0 to n-1 )
 * @throws Exception
 */
	public void copyRowFromSheet(XSSFWorkbook wbSrc, XSSFSheet wsSrc, String wsDest, int rowSrc, int rowDest,int colBegSrc,int colEndSrc) throws Exception {
		loadWorkBookInWriteMode();
		Log.info("SRC ROW = "+rowSrc);
		Log.info("DEST ROW = "+rowDest);
		sh = wb.getSheet(wsDest);
		
		if (wbSrc == null) {
			throw new Exception("Workbook/Sheet is null");
		}
		int sheetIndex = wbSrc.getSheetIndex(wsSrc.getSheetName());
		if (sheetIndex == -1) {
			throw new Exception("Sheet not found");
		}

		XSSFRow newRow = sh.getRow(rowDest);
		XSSFRow sourceRow = wsSrc.getRow(rowSrc);

		if (newRow == null) {
			newRow = sh.createRow(sh.getLastRowNum()+1);
		}

		// Loop through source columns to add to new row
		for (int i = colBegSrc; i < colEndSrc; i++) {
			// Grab a copy of the old/new cell
			XSSFCell oldCell = sourceRow.getCell(i);
			XSSFCell newCell = newRow.createCell(i);

			// If the old cell is null jump to next cell
			if (oldCell == null) {
				newCell = null;
				continue;
			}

			// Copy style from old cell and apply to new cell
		//	CellStyle newCellStyle = wb.createCellStyle();
		//	newCellStyle.cloneStyleFrom(oldCell.getCellStyle());

			//////////////////////////
		//	CellStyle origStyle = oldCell.getCellStyle(); // Or from a cell
		//	newCellStyle.cloneStyleFrom(origStyle);
		//	newCell.setCellStyle(newCellStyle);
			//////////////////
			
			
			//newCell.setCellStyle(newCellStyle);

			// If there is a cell comment, copy
			if (oldCell.getCellComment() != null) {
				newCell.setCellComment(oldCell.getCellComment());
			}

			// If there is a cell hyperlink, copy
			if (oldCell.getHyperlink() != null) {
				newCell.setHyperlink(oldCell.getHyperlink());
			}

			// Set the cell data type
			newCell.setCellType(oldCell.getCellType());

			// Set the cell data value
			switch (oldCell.getCellType()) {
			case Cell.CELL_TYPE_BLANK:
				newCell.setCellValue(oldCell.getStringCellValue());
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				newCell.setCellValue(oldCell.getBooleanCellValue());
				break;
			case Cell.CELL_TYPE_ERROR:
				newCell.setCellErrorValue(oldCell.getErrorCellValue());
				break;
			case Cell.CELL_TYPE_FORMULA:
				newCell.setCellFormula(oldCell.getCellFormula());
				break;
			case Cell.CELL_TYPE_NUMERIC:
				newCell.setCellValue(oldCell.getNumericCellValue());
				break;
			case Cell.CELL_TYPE_STRING:
				newCell.setCellValue(oldCell.getRichStringCellValue());
				break;
			}

		}
		saveReportOnDrive();
	}

public Boolean setActiveSheet(String wsName) {
	
	if(wb==null){
		Log.info("Workbook is null !");
		return false;
	}
	int sheetIndex=wb.getSheetIndex(wsName);
	if(sheetIndex==-1){
		Log.info("Sheet not found to activate");
		return false;
	}
	else{
		sh = wb.getSheetAt(sheetIndex);
		Log.info("in activate sheet fn : "+sh.getSheetName());
		return true;
	} 
	
	
	//sh = wb.getSheet(tcNameInWB);

}


public Boolean validateResultInSheet(String wsName, int expectedRowNum,int actualRowNum) throws Exception{
	Log.info("Validating sheet-"+wsName);
	loadWorkBookInWriteMode();
	int sheetIndex=wb.getSheetIndex(wsName);
	if(sheetIndex==-1)
		return false;
	else{
		sh = wb.getSheetAt(sheetIndex);
		XSSFRow expectedRow = null;
		XSSFRow actualRow = null;
		
		try{
			expectedRow = sh.getRow(expectedRowNum);
		}catch (Exception e){
			e.printStackTrace();
			Log.info("err expectedrow");
			expectedRow = sh.createRow(expectedRowNum);
		}
		try{
			actualRow = sh.getRow(actualRowNum);
		}catch (Exception e){
			e.printStackTrace();
			Log.info("err actualrow");
			actualRow = sh.createRow(actualRowNum);
		}
		
		
		
		int numOfCol1= 0;
		int numOfCol2= 0;
		if(expectedRow==null){
			 numOfCol1= 0;
		}else{
			 numOfCol1= expectedRow.getLastCellNum();
		}
		
		if(expectedRow==null){
			numOfCol2= 0;
		}else{
			numOfCol2=actualRow.getLastCellNum();
		}
		
		Log.info("EXP COLS1="+numOfCol1);
		Log.info("EXP COLS2="+numOfCol2);
		String strcellExpected ="";
		String strcellActual ="";
		
		Boolean isEqual=true;
		//Cell style to highlight in RED on mismatch - move to separate func
		XSSFCellStyle styleForMatch = wb.createCellStyle();
		XSSFCellStyle styleForMismatch = wb.createCellStyle();
		styleForMatch = wb.createCellStyle();
		styleForMismatch = wb.createCellStyle();
		mycolorMatch = new XSSFColor(Color.GREEN);
		mycolorMismatch = new XSSFColor(Color.RED);
		styleForMatch.setFillForegroundColor(mycolorMatch); 
		styleForMismatch.setFillForegroundColor(mycolorMismatch); 
		
		styleForMatch.setFillPattern(CellStyle.SOLID_FOREGROUND);
		styleForMismatch.setFillPattern(CellStyle.SOLID_FOREGROUND);
		
		XSSFCell cellExpected =null;
		XSSFCell cellActual = null;
		int maxVal = numOfCol1 >numOfCol2 ? numOfCol1 : numOfCol2;
		if (maxVal<2)
			maxVal=2;
		
			for(int c=0;c<=maxVal-1;c++){   
				Log.info("c="+c);
				try{
					 cellExpected  = expectedRow.getCell(c, MissingCellPolicy.RETURN_NULL_AND_BLANK);
				}catch (Exception e){
					e.printStackTrace();
					Log.info("err cellExpected");
					 cellExpected  = expectedRow.createCell(c);
				}
				try{
					 cellActual = actualRow.getCell(c, MissingCellPolicy.RETURN_NULL_AND_BLANK);
				}catch (Exception e){
					e.printStackTrace();
					Log.info("err cellActual");
					 cellActual = actualRow.createCell(c);
				}
				
				
				// cellActual = actualRow.getCell(c, MissingCellPolicy.RETURN_NULL_AND_BLANK);
				//Log.info("E="+cellExpected.getStringCellValue()+" ---- A="+cellActual.getStringCellValue());
				try{
					strcellExpected = cellExpected.getStringCellValue();
				}catch (Exception e){
					Log.info("err strcellExpected");
					cellExpected  = expectedRow.createCell(c);
					cellExpected.setCellValue("");
					strcellExpected = "";
				}
				 
				try{
					strcellActual =cellActual.getStringCellValue();
				}catch (Exception e){
					Log.info("err strcellActual");
					cellActual  = actualRow.createCell(c);
					cellActual.setCellValue("");
					strcellActual = "";
				}
				 
				 
				if(!(strcellExpected.trim().equals(strcellActual.trim()))){
					isEqual=false;
					cellExpected.setCellStyle(styleForMismatch);
					cellActual.setCellStyle(styleForMismatch);
					//Log.info("NOT MATCHING =  E="+cellExpected.getStringCellValue()+" ---- A="+cellActual.getStringCellValue());
				}
				else{
					//Log.info("MATCHING =  E="+cellExpected.getStringCellValue()+" ---- A="+cellActual.getStringCellValue());
					cellExpected.setCellStyle(styleForMatch);
					cellActual.setCellStyle(styleForMatch);
				}
			}
		
		if(!isEqual){
			Log.info("VALIDATION FAILED");
			saveReportOnDrive();
			String validationResultMsg= "Status : FAIL - Some Mismatches found in Expected vs Actual values";
			setValidationStatus(wsName, validationResultMsg, 10, 0, "RED");
			//saveReportOnDrive();
			return false;	
		}
		else{
			Log.info("VALIDATION PASS");
			saveReportOnDrive();
			String validationResultMsg= "Status : PASS - Expected & Actual Values match as expected";
			setValidationStatus(wsName, validationResultMsg, 10, 0, "GREEN");
			//saveReportOnDrive();
			return true;	
		}
	}	
	
}



}
