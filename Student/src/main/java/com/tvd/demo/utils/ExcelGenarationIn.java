package com.tvd.demo.utils;

import java.io.File;
import java.io.IOException;
import java.security.GeneralSecurityException;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public interface ExcelGenarationIn {
	public static XSSFFont setFont(int val, XSSFWorkbook workbook) {
		XSSFFont style = workbook.createFont();
		style.setFontHeightInPoints((short) val);

		return style;
	} 
	
		
	
	public XSSFCellStyle titleStyle(int fontsize, XSSFWorkbook workbook);
	
	public XSSFCellStyle titleStyle(int fontsize, XSSFWorkbook workbook,boolean wraptext);
	
	public XSSFCellStyle totalStyle(int fontsize, XSSFWorkbook workbook);
	
	public XSSFCellStyle genaralStyle(XSSFWorkbook workbook);
	
	public XSSFCellStyle genaralStyle(int fontsize, XSSFWorkbook workbook);
	
	public XSSFCellStyle genaralStyleWithDate(XSSFWorkbook workbook);
	
	public CellStyle genaralStyleWithdotDate(XSSFWorkbook workbook);
	
	public XSSFCellStyle genaralStyle(int fontsize, XSSFWorkbook workbook,boolean wraptext);
	
	public XSSFCellStyle tableHead(XSSFWorkbook workbook, boolean wraptext);
	
	public POIFSFileSystem Setsecurity(String fileName,String password,XSSFWorkbook workbook)throws InvalidFormatException, IOException, GeneralSecurityException;
	
	public void DownloadNonSecuredDocument(HttpServletResponse response,String fileName,XSSFWorkbook workbook)throws IOException;
	
	public void DownloadProtectedDocument(POIFSFileSystem poiFileSystem,HttpServletResponse response,String fileName,XSSFWorkbook workbook) throws IOException;
	
	public void SetPicture(XSSFWorkbook workbook,XSSFSheet worksheet,int col1,int col2,int row1,int row2,String ImageFileName)throws IOException;
	
	public XSSFCellStyle MergedRegionWithStyle(int FirstRow,int LastRow,int FirstCol, int LastCol,int fontsize, XSSFSheet worksheet,XSSFWorkbook workbook);
	
	public File getPath(String fileName) throws IOException;
	
	public void setheadder(XSSFSheet worksheet,String Stage,String Date);
	
	public void setfotter(XSSFSheet worksheet,String Center,String Username);
	
	public void setproperty(XSSFWorkbook workbook,String Creator,String Title);
	
	public XSSFCellStyle genaralStylewithdot(XSSFWorkbook workbook);

	public  XSSFCellStyle setvalue(XSSFRow row,Integer cell,Integer value,XSSFWorkbook workbook);
	
	public  XSSFCellStyle setvalue(XSSFRow row,Integer cell,String value,XSSFWorkbook workbook);
	
	public  XSSFCellStyle setvalue(XSSFRow row,Integer cell,Double value,XSSFWorkbook workbook);
	
	public  XSSFCellStyle setvalue(XSSFRow row,Integer cell,Long value,XSSFWorkbook workbook);

	public  XSSFCellStyle setformula(XSSFRow row,Integer cell,String formula,XSSFWorkbook workbook);
	
	public  XSSFCellStyle setTitle(XSSFRow row,Integer cell,String value,XSSFWorkbook workbook);
	
	public  XSSFCellStyle setheadding(XSSFRow row,Integer cell,String value,XSSFWorkbook workbook);
}
