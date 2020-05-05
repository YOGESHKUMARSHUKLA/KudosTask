package starter.generic;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.xssf.usermodel.*;

import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;



public class Read_XLS {	
	public String filelocation;
	public  FileInputStream ipstr = null;
	public  FileOutputStream opstr =null;
	public XSSFWorkbook wb = null;
	public XSSFSheet ws = null;	
	public File wbFileObj = null;
	
	 private  XSSFCell cell;
	 private  XSSFCellStyle cellstyle;
	 private  XSSFColor mycolorMatch;
	 private  XSSFColor mycolorMismatch;
	 private  XSSFRow row;

		
	public Read_XLS(String filelocation) {		
		this.filelocation=filelocation;
		try {
			wbFileObj = new File(filelocation);
			ipstr = new FileInputStream(filelocation);
			//System.out.println("fl"+filelocation);
			wb = new XSSFWorkbook(ipstr);
			ws = wb.getSheetAt(0);
			//System.out.println("sheet Name--"+ws.getSheetName());
			ipstr.close();
		} catch (Exception e) {			
			e.printStackTrace();
		} 
		
	}
	

	//To retrieve No Of Rows from excel file
	public int retrieveNoOfRows(String wsName){		
		int sheetIndex=wb.getSheetIndex(wsName);
		if(sheetIndex==-1)
			return 0;
		else{
			ws = wb.getSheetAt(sheetIndex);
			int rowCount=ws.getLastRowNum()+1;		
			return rowCount;		
		} 
	}
	
	//To retrieve No Of Columns
	public int retrieveNoOfCols(String wsName){
		int sheetIndex=wb.getSheetIndex(wsName);
		if(sheetIndex==-1)
			return 0;
		else{
			ws = wb.getSheetAt(sheetIndex);
			int colCount=ws.getRow(0).getLastCellNum();			
			return colCount;
		}
	}
	
	//To retrieve Run flag of test suite and test case.
	public String retrieveToRunFlag(String wsName, String colName, String rowName){
		
		int sheetIndex=wb.getSheetIndex(wsName);
		if(sheetIndex==-1)
			return null;
		else{
			int rowNum = retrieveNoOfRows(wsName);
			int colNum = retrieveNoOfCols(wsName);
			int colNumber=-1;
			int rowNumber=-1;			
			
			XSSFRow Suiterow = ws.getRow(0);				
			
			for(int i=0; i<colNum; i++){
				if(Suiterow.getCell(i).getStringCellValue().equals(colName.trim())){
					colNumber=i;					
				}					
			}
			
			if(colNumber==-1){
				return "";				
			}
			
			
			for(int j=0; j<rowNum; j++){
				XSSFRow Suitecol = ws.getRow(j);				
				if(Suitecol.getCell(0).getStringCellValue().equals(rowName.trim())){
					rowNumber=j;	
				}					
			}
			
			if(rowNumber==-1){
				return "";				
			}
			
			XSSFRow row = ws.getRow(rowNumber);
			XSSFCell cell = row.getCell(colNumber);
			if(cell==null){
				return "";
			}
			String value = cellToString(cell);
			return value;			
		}			
	}
	
	//To retrieve DataToRun flag of test data.
	public String[] retrieveToRunFlagTestData(String wsName, String colName){
		
		int sheetIndex=wb.getSheetIndex(wsName);
		if(sheetIndex==-1)
			return null;
		else{
			int rowNum = retrieveNoOfRows(wsName);
			int colNum = retrieveNoOfCols(wsName);
			int colNumber=-1;
					
			
			XSSFRow Suiterow = ws.getRow(0);				
			String data[] = new String[rowNum-1];
			for(int i=0; i<colNum; i++){
				if(Suiterow.getCell(i).getStringCellValue().equals(colName.trim())){
					colNumber=i;					
				}					
			}
			
			if(colNumber==-1){
				return null;				
			}
			
			for(int j=0; j<rowNum-1; j++){
				XSSFRow Row = ws.getRow(j+1);
				if(Row==null){
					data[j] = "";
				}
				else{
					XSSFCell cell = Row.getCell(colNumber);
					if(cell==null){
						data[j] = "";
					}
					else{
						String value = cellToString(cell);
						data[j] = value;	
					}	
				}
			}
			
			return data;			
		}			
	}
	
	//To retrieve test data from test case data sheets.
	public Object[][] retrieveTestData(String wsName){
		int sheetIndex=wb.getSheetIndex(wsName);
		if(sheetIndex==-1){
			System.out.println("Inside retrieveTestData - no sheet found");
			return null;
		}
		else{
				int rowNum = retrieveNoOfRows(wsName);
				int colNum = retrieveNoOfCols(wsName);
		
				Object data[][] = new Object[rowNum-1][colNum-2];
				//System.out.println("rowN: "+rowNum);
				for (int i=0; i<rowNum-1; i++){
					XSSFRow row = ws.getRow(i+1);
					
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
	
	//Return Data from excel sheet as string
	public String[][] getDataAsStringArray(String wsName){
		
		int sheetIndex=wb.getSheetIndex(wsName);
		if(sheetIndex==-1){
			System.out.println("Inside retrieveTestData - no sheet found");
			return null;
		}
		else{
				int rowNum = retrieveNoOfRows(wsName);
				int colNum = 3; //HARDCODED
				
				//System.out.println("rowNum - "+ rowNum);
				//System.out.println("colNum - "+ colNum);
				String data[][] = new String[rowNum][colNum];
				//System.out.println("rowN: "+rowNum);
				for (int i=0; i<rowNum; i++){
					XSSFRow row = ws.getRow(i);
					
					for(int j=0; j< colNum; j++){					
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
								//int value1=Integer.parseInt(value);
								data[i][j] = value;	
								//System.out.println(value);
							}
						}
					}				
				}			
				return data;		
		}
	
	}		
	
	
	public static String cellToString(XSSFCell cell){
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
	
	//To write result In test data and test case list sheet.
	public boolean writeResult(String wsName, String colName, int rowNumber, String Result){
		try{
			int sheetIndex=wb.getSheetIndex(wsName);
			if(sheetIndex==-1)
				return false;			
			int colNum = retrieveNoOfCols(wsName);
			int colNumber=-1;
					
			
			XSSFRow Suiterow = ws.getRow(0);			
			for(int i=0; i<colNum; i++){				
				if(Suiterow.getCell(i).getStringCellValue().equals(colName.trim())){
					colNumber=i;					
				}					
			}
			
			if(colNumber==-1){
				return false;				
			}
			
			XSSFRow Row = ws.getRow(rowNumber);
			XSSFCell cell = Row.getCell(colNumber);
			if (cell == null)
		        cell = Row.createCell(colNumber);			
			
			cell.setCellValue(Result);
			
			opstr = new FileOutputStream(filelocation);
			wb.write(opstr);
			opstr.close();
			
			
		}catch(Exception e){
			e.printStackTrace();
			return false;
		}
		return true;
	}
	
	//To write result In test suite list sheet.
	public boolean writeResult(String wsName, String colName, String rowName, String Result){
		try{
			int rowNum = retrieveNoOfRows(wsName);
			int rowNumber=-1;
			int sheetIndex=wb.getSheetIndex(wsName);
			if(sheetIndex==-1)
				return false;			
			int colNum = retrieveNoOfCols(wsName);
			int colNumber=-1;
					
			
			XSSFRow Suiterow = ws.getRow(0);			
			for(int i=0; i<colNum; i++){				
				if(Suiterow.getCell(i).getStringCellValue().equals(colName.trim())){
					colNumber=i;					
				}					
			}
			
			if(colNumber==-1){
				return false;				
			}
			
			for (int i=0; i<rowNum-1; i++){
				XSSFRow row = ws.getRow(i+1);				
				XSSFCell cell = row.getCell(0);	
				cell.setCellType(Cell.CELL_TYPE_STRING);
				String value = cellToString(cell);	
				if(value.equals(rowName)){
					rowNumber=i+1;
					break;
				}
			}		
			
			XSSFRow Row = ws.getRow(rowNumber);
			XSSFCell cell = Row.getCell(colNumber);
			if (cell == null)
		        cell = Row.createCell(colNumber);			
			
			cell.setCellValue(Result);
			
			opstr = new FileOutputStream(filelocation);
			wb.write(opstr);
			opstr.close();
			
			
		}catch(Exception e){
			e.printStackTrace();
			return false;
		}
		return true;
	}
	
	//To write result In test data and test case list sheet.
	public boolean writeScrappedValuesToSheet(String wsName,ArrayList<String> Result){
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
			int rowNum = retrieveNoOfRows(wsName);
			XSSFRow row = ws.createRow(rowNum);
			
			System.out.println("Writing Scrapped Values to Sheet");
			int colNo = 0;
			for (String element : Result) {
				System.out.println(element);
				XSSFCell cell = row.createCell(colNo++);
				cell.setCellStyle(style);
				cell.setCellValue(element);
			}
			
		/*	opstr = new FileOutputStream(prop.getProperty("ExtentReportResults"+));
			wb.write(opstr);
			opstr.close();
			*/
			
		}catch(Exception e){
			e.printStackTrace();
			return false;
		}
		return true;
	}

	public String saveReport() throws IOException {
		// TODO Auto-generated method stub
		//String reportPath = prop.getProperty("ExtentReportResults") + wbFileObj.getName();
		String reportPath =  wbFileObj.getAbsolutePath();
		System.out.println(reportPath);
		//Delete if file exists, (to be replaced with archive function)
		 try{
             File file = new File(reportPath);
             if(file.delete()){
                 System.out.println(file.getName() + " Was deleted!");
             }else{
                 System.out.println("Delete Operation Failed. Check: " + file);
             }
         }catch(Exception e1){
             e1.printStackTrace();
         }
		 
		opstr = new FileOutputStream(reportPath);
		wb.write(opstr);
		opstr.close();
		return reportPath;
	}
	
	public Boolean validateResultInSheet(String wsName, int expectedRowNum,int actualRowNum) throws Exception{
		System.out.println("Validating sheet-"+wsName);
		int sheetIndex=wb.getSheetIndex(wsName);
		if(sheetIndex==-1)
			return false;
		else{
			ws = wb.getSheetAt(sheetIndex);
			XSSFRow expectedRow = ws.getRow(expectedRowNum);
			XSSFRow actualRow = ws.getRow(actualRowNum);
			int numOfCol1= expectedRow.getLastCellNum();
			int numOfCol2=actualRow.getLastCellNum();
			
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
			
			int maxVal = numOfCol1 > numOfCol2 ? numOfCol1 : numOfCol2;
			
				for(int c=0;c<maxVal-2;c++){                                //Ignoring last 2 columns
					XSSFCell cellExpected  = expectedRow.getCell(c, MissingCellPolicy.RETURN_NULL_AND_BLANK);
					XSSFCell cellActual = actualRow.getCell(c, MissingCellPolicy.RETURN_NULL_AND_BLANK);
					//System.out.println("E="+cellExpected.getStringCellValue()+" ---- A="+cellActual.getStringCellValue());
					
					if(!(cellExpected.getStringCellValue().equals(cellActual.getStringCellValue()))){
						isEqual=false;
						cellExpected.setCellStyle(styleForMismatch);
						cellActual.setCellStyle(styleForMismatch);
						System.out.println("NOT MATCHING =  E="+cellExpected.getStringCellValue()+" ---- A="+cellActual.getStringCellValue());
					}
					else{
						System.out.println("MATCHING =  E="+cellExpected.getStringCellValue()+" ---- A="+cellActual.getStringCellValue());
						cellExpected.setCellStyle(styleForMatch);
						cellActual.setCellStyle(styleForMatch);
					}
				}
			
			if(!isEqual){
				System.out.println("VALIDATION FAILED");
				setCellData("Test Case : FAILED - Some Mismatches found in Expected vs Actual values", 5,0);
				return false;	
			}
			else
				return true;	
		}	
		
	}
	
	/*public void highlightCell(String color, int rownum, int colnum)throws Exception
	 {
	  try{
	//  cell = sh.getRow(rownum).getCell(colnum,Row.RETURN_BLANK_AS_NULL);
	  cell = ws.getRow(rownum).getCell(colnum, null);
	   }catch(Exception e){System.out.println("cell is null");}
	  if (cell == null) 
	    {
	   row=ws.getRow(rownum);
	   cell = row.createCell(colnum);
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
	 }
	 
	 
	 
	 
	public String getCellData(int rownum, int colnum) throws Exception
	 {
	  try{
	     cell = ws.getRow(rownum).getCell(colnum);
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
	 }*/
	 
	

	
	
	/***
	 * @param text
	 * @param rownum
	 * @param colnum
	 * @throws Exception
	 */
	public void setCellData(String text, int rownum, int colnum) throws Exception
	{
	 try{   
	    row  = ws.getRow(rownum);
	    if(row ==null)
	    {
	       row = ws.createRow(rownum);
	    }
	    cell = row.getCell(colnum);
	   if (cell != null) 
	    {
	        cell.setCellValue(text);
	    } 
	    else 
	    {
	         cell = row.createCell(colnum);
	         cell.setCellValue(text);  
	    }
	 }catch(Exception e){throw (e);} }

	public int getTotalSuites() {
		// TODO Auto-generated method stub
		return 0;
	}
}
