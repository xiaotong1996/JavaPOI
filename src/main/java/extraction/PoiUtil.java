package extraction;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

public class PoiUtil {
	
	private static String info;
	
	public static String getInfo()
	{
		return info;
	}
	
	/********************************************************
	 * function : copy a cell (except the merged cell) in the same sheet
	 ********************************************************/
	public static boolean copyCell(XSSFSheet xssfSheet,int rowIn,int colIn,int rowOut,int colOut)
	{
		XSSFRow rowI=xssfSheet.getRow(rowIn-1);
		XSSFCell cellI=rowI.getCell(colIn-1);
		XSSFCellStyle cellStyleIn=cellI.getCellStyle();
		
		XSSFRow rowO=xssfSheet.getRow(rowOut-1);
		XSSFCell cellO=rowO.getCell(colOut-1);
		
		cellO.setCellValue(getCellValue(cellI));
		cellO.setCellStyle(cellStyleIn);
		
		return true;
	}
	/********************************************************
	 * function : copy a part of cells in the same sheet
	 ********************************************************/
	public static boolean copyCells( XSSFWorkbook wb,int indexSheet,int rowInStart,int colInStart,int rowInEnd,int colINEnd,int rowOutStart,int colOutStart )
	{
		XSSFSheet sheet=wb.getSheetAt(indexSheet);
		XSSFRow rowIn;
		XSSFCell cellIn;
		XSSFCellStyle cellStyleIn;
		String cellInValue;
		
		
		XSSFRow rowOut;
		XSSFCell cellOut;
		XSSFCellStyle cellStyleOut;
		
		int rowNum=rowInEnd-rowInStart;
		int colNum=colINEnd-colInStart;
		for(int i=0;i<rowNum+1;i++)
	    {  	    	
	    	
	    	rowIn=sheet.getRow(rowInStart+i-1); 		    		
	    	rowOut = sheet.getRow(rowOutStart+i-1);
	    	
			//set the copied row's height
//	    	rowOut.setHeight(rowIn.getHeight());
	    	
	    	for(int j=0;j<colNum+1;j++) {
	    		cellIn=rowIn.getCell(colInStart+j-1);
	    		if(cellIn!=null) {
	    		cellInValue=getCellValue(cellIn);
	    	    cellOut=rowOut.createCell(colOutStart+j-1);
	    	    cellStyleIn=cellIn.getCellStyle();
	    	    cellStyleOut=wb.createCellStyle();
	    	    cellStyleOut.cloneStyleFrom(cellStyleIn);
	    	    cellOut.setCellStyle(cellStyleOut);
	    	    cellOut.setCellValue(cellInValue);
	    		}
	    	}	
	    }
		
		
		//in order to deal with merged regions
		java.util.List<CellRangeAddress> regions=sheet.getMergedRegions();
		
		for(CellRangeAddress cellRangeAddress : regions)
		{
			if(cellRangeAddress.getFirstColumn()>=colInStart-1&&
					cellRangeAddress.getLastColumn()<=colINEnd-1&&
					cellRangeAddress.getFirstRow()>=rowInStart-1&&
					cellRangeAddress.getLastRow()<=rowInEnd-1)
			{
				int diffrow=rowOutStart-rowInStart;
				int diffcol=colOutStart-colInStart;
			int firstRow=cellRangeAddress.getFirstRow()+diffrow;
			int firstCol=cellRangeAddress.getFirstColumn()+diffcol;
			int lastRow=firstRow+cellRangeAddress.getLastRow()-cellRangeAddress.getFirstRow();
			int lastCol=firstCol+cellRangeAddress.getLastColumn()-cellRangeAddress.getFirstColumn();
			CellRangeAddress cellRangeAddressNew=new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);

			sheet.addMergedRegion(cellRangeAddressNew);
			}
		}
		
		//set the copied column's width
		int diffCol=colOutStart-colInStart;
		for(int columnIndex=colInStart-1;columnIndex<=colINEnd-1;columnIndex++) {
			int width=sheet.getColumnWidth(columnIndex);
			sheet.setColumnWidth(columnIndex+diffCol, width);
		}
		
	    return true;
	}
	
	
	/********************************************************
	 * function : just copy the cells' values and styles (not include merge cells)
	 * *******************************************************/
	public static boolean copySimCells( XSSFWorkbook wb,int indexSheet,int rowInStart,int colInStart,int rowInEnd,int colINEnd,int rowOutStart,int colOutStart )
	{
		XSSFSheet sheet=wb.getSheetAt(indexSheet);
		XSSFRow rowIn;
		XSSFCell cellIn;
		XSSFCellStyle cellStyleIn;
		String cellInValue;
		
		
		XSSFRow rowOut;
		XSSFCell cellOut;
		XSSFCellStyle cellStyleOut;
		
		int rowNum=rowInEnd-rowInStart;
		int colNum=colINEnd-colInStart;
		for(int i=0;i<rowNum+1;i++)
	    {  	    	
	    	
	    	rowIn=sheet.getRow(rowInStart+i-1); 		    		
	    	rowOut = sheet.getRow(rowOutStart+i-1);
	    	
	    	for(int j=0;j<colNum+1;j++) {
	    		cellIn=rowIn.getCell(colInStart+j-1);
	    		if(cellIn!=null) {
	    		cellInValue=getCellValue(cellIn);
	    	    cellOut=rowOut.createCell(colOutStart+j-1);
	    	    cellStyleIn=cellIn.getCellStyle();
//	    	    cellStyleOut=wb.createCellStyle();
//	    	    cellStyleOut.cloneStyleFrom(cellStyleIn);
	    	    cellOut.setCellStyle(cellStyleIn);
	    	    cellOut.setCellValue(cellInValue);
	    		}
	    	}	
	    }
	    return true;
	}
	
	
	/********************************************************
	 * function : copy a part of cells in different workbooks
	 * *******************************************************
	 * @param:
	 * xbIn : the resource workbook
	 * xbOut : the destination workbook
	 * indexSheetIn : the index of the sheet in the xbIn (start from 0)
	 * indexSheetOut : the index of the sheet in the xbOut (start from 0)
	 * rowInStart : the number of the first row that we need in the indexSheetIn ( start from 1 )
	 * rowInEnd : the number of the last row that we need in the indexSheetIn ( start from 1 )
	 * colInStart : the number of the first column that we need in the indexSheetIn ( start from 1, represent A in the sheet )
	 * colInEnd : the number of the last column that we need in the indexSheetIn ( start from 1, represent A in the sheet )
	 * rowOutStart : the number of the first row that we want to put cells in the indexSheetOut ( start from 1 )
	 * colOutStart : the number of the first column that we want to put cells in the indexSheetOut ( start from 1, represent A in the sheet )
	 * *******************************************************/
	public static boolean copyCells(XSSFWorkbook wbIn,XSSFWorkbook wbOut,int indexSheetIn,int indexSheetOut, 
			int rowInStart,int colInStart,int rowInEnd,int colINEnd,int rowOutStart,int colOutStart )
	{
		XSSFSheet sheetIn=wbIn.getSheetAt(indexSheetIn);
		XSSFRow rowIn;
		XSSFCell cellIn;
		XSSFCellStyle cellStyleIn;
		String cellInValue;
		
		
		XSSFSheet sheetOut=wbOut.getSheetAt(indexSheetOut);
		XSSFRow rowOut;
		XSSFCell cellOut;
		XSSFCellStyle cellStyleOut;
		
		int rowNum=rowInEnd-rowInStart;
		int colNum=colINEnd-colInStart;
		for(int i=0;i<rowNum+1;i++)
	    {  	    	
	    	
	    	rowIn=sheetIn.getRow(rowInStart+i-1); 		    		
	    	rowOut = sheetOut.createRow(rowOutStart+i-1);
			//set the copied row's height
	    	rowOut.setHeight(rowIn.getHeight());
	    	for(int j=0;j<colNum+1;j++) {
	    		cellIn=rowIn.getCell(colInStart+j-1);
	    		if(cellIn!=null) {
	    		cellInValue=getCellValue(cellIn);
	    	    cellOut=rowOut.createCell(colOutStart+j-1);
	    	    cellStyleIn=cellIn.getCellStyle();
	    	    cellStyleOut=wbOut.createCellStyle();
	    	    cellStyleOut.cloneStyleFrom(cellStyleIn);
	    	    cellOut.setCellStyle(cellStyleOut);
	    	    cellOut.setCellValue(cellInValue);
	    		}
	    	}	
	    }
		
		
		//in order to deal with merged regions
		java.util.List<CellRangeAddress> regions=sheetIn.getMergedRegions();
		
		for(CellRangeAddress cellRangeAddress : regions)
		{
			if(cellRangeAddress.getFirstColumn()>=colInStart-1&&
					cellRangeAddress.getLastColumn()<=colINEnd-1&&
					cellRangeAddress.getFirstRow()>=rowInStart-1&&
					cellRangeAddress.getLastRow()<=rowInEnd-1)
			{
				int diffrow=rowOutStart-rowInStart;
				int diffcol=colOutStart-colInStart;
			int firstRow=cellRangeAddress.getFirstRow()+diffrow;
			int firstCol=cellRangeAddress.getFirstColumn()+diffcol;
			int lastRow=firstRow+cellRangeAddress.getLastRow()-cellRangeAddress.getFirstRow();
			int lastCol=firstCol+cellRangeAddress.getLastColumn()-cellRangeAddress.getFirstColumn();
			CellRangeAddress cellRangeAddressNew=new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);

			sheetOut.addMergedRegion(cellRangeAddressNew);
			}
		}
		//set the copied column's width
		for(int columnIndex=colInStart-1;columnIndex<=colINEnd-1;columnIndex++) {
			int width=sheetIn.getColumnWidth(columnIndex);
			sheetOut.setColumnWidth(columnIndex, width);
		}
		
	    return true;
	}
	
	public static void setCellValue(XSSFWorkbook wb,int indexSheet,int rowIn,int colIn,String value)
	{
		XSSFSheet sheet=wb.getSheetAt(indexSheet);
		XSSFRow row=sheet.getRow(rowIn-1);
		XSSFCell cell=row.getCell(colIn-1);
		cell.setCellValue(value);     
	}
	/********************************************************
	 * function : merge a region
	 ********************************************************/
	public static void mergeRegion(XSSFWorkbook wb,int indexSheet,int rowStart,int colStart,int rowEnd,int colEnd)
	{
		XSSFSheet sheet=wb.getSheetAt(indexSheet);
		sheet.addMergedRegion(new CellRangeAddress(
	    		rowStart-1, //first row (0-based)
	            rowEnd-1, //last row  (0-based)
	            colStart-1, //first column (0-based)
	            colEnd-1  //last column  (0-based)
	    ));
	}
	/********************************************************
	 * function : set Borders of a cell
	 ********************************************************/
	public static void setBorder(Cell cell)
	{
		XSSFCellStyle style = (XSSFCellStyle) cell.getCellStyle();     
		style.setBorderBottom(BorderStyle.THIN);
	    style.setBorderLeft(BorderStyle.THIN);
	    style.setBorderRight(BorderStyle.THIN);
	    style.setBorderTop(BorderStyle.THIN);
	    cell.setCellStyle(style);
	}
	
	/********************************************************
	 * function : set Borders of a cell
	 ********************************************************/
	public static void setBorder(XSSFWorkbook wb,int indexSheet,int rowIn,int colIn)
	{
		XSSFSheet sheet=wb.getSheetAt(indexSheet);
		XSSFRow row=sheet.getRow(rowIn-1);
		XSSFCell cell=row.getCell(colIn-1);
		XSSFCellStyle style = (XSSFCellStyle) cell.getCellStyle();     
		style.setBorderBottom(BorderStyle.THIN);
	    style.setBorderLeft(BorderStyle.THIN);
	    style.setBorderRight(BorderStyle.THIN);
	    style.setBorderTop(BorderStyle.THIN);
	    cell.setCellStyle(style);
	}
	/********************************************************
	 * function : set Borders of a region
	 ********************************************************/
	public static void setRegionBorder(XSSFWorkbook wb,int indexSheet,int rowStart,int colStart,int rowEnd,int colEnd)
	{
  
		for(int i=colStart;i<=colEnd;i++)
		{
			setSingleBorder(wb, indexSheet, rowStart, i, 'T');
			setSingleBorder(wb, indexSheet, rowEnd, i, 'B');
		}
		for(int j=rowStart;j<=rowEnd;j++)
		{
			setSingleBorder(wb, indexSheet, j, colStart, 'L');
			setSingleBorder(wb, indexSheet, j, colEnd, 'R');
		}
					

	}
	
	
	

	/********************************************************
	 * function : set a Border of a cell
	 ********************************************************/
	public static void setSingleBorder(XSSFWorkbook wb,int indexSheet,int rowIn,int colIn,char top)
	{
		XSSFSheet sheet=wb.getSheetAt(indexSheet);
		XSSFRow row=sheet.getRow(rowIn-1);
		XSSFCell cell=row.getCell(colIn-1);
		XSSFCellStyle style = cell.getCellStyle(); 
		switch (top) {
		case 'B':
			style.setBorderBottom(BorderStyle.THIN);
			break;
		case 'L':
		    style.setBorderLeft(BorderStyle.THIN);
		    break;
		case 'T':
			style.setBorderTop(BorderStyle.THIN);
			break;
		case 'R':
		    style.setBorderRight(BorderStyle.THIN);
		    break;   
		default:
			break;
		}
	    cell.setCellStyle(style);
			
	}
	
	/********************************************************
	 * function : get the projet's path
	 ********************************************************/
	public static String getPath() {      
	    File file=new File("");
        String abspath=file.getAbsolutePath();
        return abspath;
	}
	
	
	/********************************************************
	 * function : get an instance of a workbook
	 ********************************************************/
	public static XSSFWorkbook openWorkbook(String pathIn) throws IOException
	{
		File file = new File(pathIn);
	      FileInputStream fIP = new FileInputStream(file);
	      //Get the workbook instance for XLSX file 
	      XSSFWorkbook workbook = new XSSFWorkbook(fIP);
	      if(file.isFile() && file.exists())
	      {
	         info= pathIn+" file open successfully.\n";
	      }
	      else
	      {
	         info="Error to open workbook.xlsx file.\n";
	      }
		return workbook;
	}
	
	
	
	/********************************************************
	 * function : set a sheet's title
	 * @throws IOException 
	 ********************************************************/
	public static boolean setTitle(XSSFCell cellIn, XSSFWorkbook workbook,int indexSheet,String title) throws IOException
	{
		if (title == null || indexSheet <=0 || workbook==null) {
			return false;
		}
		else
		{
//			String path=getPath();
//			String pathIn=path+"/src/main/resources/FichierEntree_parJury.xlsx";
//			//get the source cell style
//		 	XSSFWorkbook workbookIn =openWorkbook(pathIn);
//		    XSSFSheet sheetIn=workbookIn.getSheetAt(0);
//		    XSSFRow rowIn=sheetIn.getRow(0);
//		    XSSFCell cellIn=rowIn.getCell(0);
		    XSSFCellStyle cellStyleIn=cellIn.getCellStyle();
		    
		    XSSFSheet sheet=workbook.getSheetAt(indexSheet);
		    //create the title cell
		    Row row0 = sheet.createRow(0);
		    row0.setHeight((short) 400);
		    Cell cell0 = row0.createCell(0);
		    
		    //merge the region a1-d1
		    sheet.addMergedRegion(new CellRangeAddress(
		    		0, //first row (0-based)
		            0, //last row  (0-based)
		            0, //first column (0-based)
		            3  //last column  (0-based)
		    ));
		    
		    
		    //set the title cell's style according to the source cell style
		    XSSFCellStyle cellStyleOut = workbook.createCellStyle();
		    cellStyleOut.cloneStyleFrom(cellStyleIn);
		    cell0.setCellStyle(cellStyleOut);
		    
		    //set the title cell's value
		    cell0.setCellValue(title);
		    
		    
		    return true;
		}
		
	}
	
	public static void setFormula(XSSFWorkbook wb,int indexSheet,int rowIn,int colIn,String index)
	{
		XSSFSheet sheet=wb.getSheetAt(indexSheet);
		XSSFRow row=sheet.getRow(rowIn-1);
		XSSFCell cell=row.getCell(colIn-1);
		if(cell!=null)
		{
			cell.setCellType(XSSFCell.CELL_TYPE_FORMULA);
			cell.setCellFormula(index);
		}
		else {
			cell=row.createCell(colIn-1);
			cell.setCellFormula(index);
			setBorder(cell);
		}

	}
	
	
	
	 /**********************************************************
     * Excel column index begin 1
     * @param colStr
     * @param length
     * @return
     **********************************************************/
    public static int excelColStrToNum(String colStr, int length) {
        int num = 0;
        int result = 0;
        for(int i = 0; i < length; i++) {
            char ch = colStr.charAt(length - i - 1);
            num = (int)(ch - 'A' + 1) ;
            num *= Math.pow(26, i);
            result += num;
        }
        return result;
    }

    /**
     * Excel column index begin 1
     * @param columnIndex
     * @return
     */
    public static String excelColIndexToStr(int columnIndex) {
        if (columnIndex <= 0) {
            return null;
        }
        String columnStr = "";
        columnIndex--;
        do {
            if (columnStr.length() > 0) {
                columnIndex--;
            }
            columnStr = ((char) (columnIndex % 26 + (int) 'A')) + columnStr;
            columnIndex = (int) ((columnIndex - columnIndex % 26) / 26);
        } while (columnIndex > 0);
        return columnStr;
    }
	
	/********************************************************
	 * function : get cell's value, this is used for the function "copyCells"
	 * ******************************************************* */	
	public static String getCellValue(XSSFCell cell) {
		String cellValue="";
		switch (cell.getCellType()) {
		case XSSFCell.CELL_TYPE_STRING:
			cellValue=cell.getStringCellValue();
			if(cellValue.trim().equals("")||cellValue.trim().length()<=0)
				cellValue="";
			break;
		case XSSFCell.CELL_TYPE_NUMERIC:
			cellValue=String.valueOf(cell.getNumericCellValue());
			break;
		case XSSFCell.CELL_TYPE_FORMULA:
			cell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
			cellValue=String.valueOf(cell.getNumericCellValue());
			break;
		case XSSFCell.CELL_TYPE_BLANK:
			cellValue="";
			break;

		default:
			break;
		}
		return cellValue;
	}
	
	/********************************************************
	 * function : count the number of valid rows in a given sheet 
	 * *******************************************************
	 * parameters:
	 * wb : workbook which contains the sheet we want
	 * indexSheet : the index of the sheet in the xb (start from 0)
	 *************************************************************** */
	public static int getSheetRowNumber(XSSFWorkbook wb,int indexSheet)
	{
		XSSFSheet sheet=wb.getSheetAt(indexSheet);
		int count=0;
		int begin = sheet.getFirstRowNum();  
		  
	    int end = sheet.getLastRowNum();  
	  
	    for (int i = begin; i <= end; i++) {  
	        if (null == sheet.getRow(i)|| getCellValue(sheet.getRow(i).getCell(0)) == "" || null==sheet.getRow(i).getCell(0)) {  
	            continue;  
	        }  
	        else count++;
	    }
	    
	    return count;
	}
	
	/********************************************************
	 * function : count the number of valid rows in a given sheet 
	 * ********************************************************/
	public static int getSheetRowNumber(XSSFSheet sheet)
	{
		int count=0;
		int begin = sheet.getFirstRowNum();  
		  
	    int end = sheet.getLastRowNum();  
	  
	    for (int i = begin; i <= end; i++) {  
	        if (null == sheet.getRow(i)||  null==sheet.getRow(i).getCell(0) || getCellValue(sheet.getRow(i).getCell(0)) == "") {  
	            continue;  
	        }  
	        else count++;
	    }
	    return count;
	}
	
	/********************************************************
	 * function : set Row's height
	 * ******************************************************* */	
	public static boolean setRowHeight(XSSFWorkbook wb,int indexSheet,int rowNum,short height)
	{
		XSSFSheet sheet=wb.getSheetAt(indexSheet);
		XSSFRow row=sheet.getRow(rowNum-1);
		row.setHeight(height);
		return true;
	}
	
	/********************************************************
	 * function : set Row's height
	 * ******************************************************* */	
	public static boolean setRowHeight(XSSFSheet sheet,int rowNum,short height)
	{
		XSSFRow row=sheet.getRow(rowNum-1);
		row.setHeight(height);
		return true;
	}
	
	/********************************************************
	 * function : set column's width
	 * ******************************************************* */	
	public static boolean setColWidth(XSSFWorkbook wb,int indexSheet,short colNum,int width)
	{
		XSSFSheet sheet=wb.getSheetAt(indexSheet);
		sheet.setColumnWidth(colNum-1, width);
		return true;
	}
	
	
}
