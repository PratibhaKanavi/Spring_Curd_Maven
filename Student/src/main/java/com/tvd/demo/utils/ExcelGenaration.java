package com.tvd.demo.utils;

import java.awt.Color;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.security.GeneralSecurityException;

import javax.servlet.http.HttpServletResponse;
import javax.swing.plaf.synth.Region;

import org.apache.poi.POIXMLProperties;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;


@Component
public class ExcelGenaration {

	
	
	
	
	
	
	
	private short cellstyleborder=CellStyle.BORDER_THIN;
		private short color=IndexedColors.GREY_40_PERCENT.getIndex();
		public XSSFCellStyle cellStyle;
		
		
		
		public ExcelGenaration() {
			super();
			// TODO Auto-generated constructor stub
		}
	
		public ExcelGenaration(XSSFWorkbook workbook) {
			 this.cellStyle = workbook.createCellStyle();
		}
	
	
	
	
	
	

	public XSSFCellStyle titleStyle(int fontsize, XSSFWorkbook workbook) {

		XSSFCellStyle cellStyle=workbook.createCellStyle();
		XSSFFont font = setFont(fontsize, workbook);
		font.setBold(true);
		cellStyle.setFont(font);

		cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
		cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		return cellStyle;

	}

	public XSSFCellStyle titleStyle(int fontsize, XSSFWorkbook workbook, boolean wraptext) {

		XSSFCellStyle cellStyle=workbook.createCellStyle();
		XSSFFont font = setFont(fontsize, workbook);
		font.setBold(true);
		cellStyle.setFont(font);
		cellStyle.setWrapText(wraptext);
		cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
		cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		return cellStyle;

	}

	public XSSFCellStyle totalStyle(XSSFRow row,Integer cell,String value, XSSFWorkbook workbook,int fontsize) {
		XSSFCellStyle cellStyle=workbook.createCellStyle();
		XSSFFont font = setFont(fontsize, workbook);
		font.setBold(true);
		font.setFontName("Calibri");
		font.setFontHeightInPoints((short)11);
		
		//XSSFColor color=workbook.getCustomPalette().
		cellStyle.setFillForegroundColor(new XSSFColor(new Color(196,188,150)));
		cellStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
		cellStyle.setFont(font);
		
		cellStyle.setBorderTop(cellstyleborder);
		cellStyle.setBorderBottom(cellstyleborder);
		cellStyle.setBorderLeft(cellstyleborder);
		cellStyle.setBorderRight(cellstyleborder);
		cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
		cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		
		cellStyle.setBottomBorderColor(color);
		cellStyle.setTopBorderColor(color);
		cellStyle.setLeftBorderColor(color);
		cellStyle.setRightBorderColor(color);
		
		if (value != null) {
			row.createCell(cell).setCellValue(value);
		} else {
			row.createCell(cell).setCellValue(0);
		}
		row.getCell(cell).setCellStyle(cellStyle);
		return cellStyle;

	}
	
	public XSSFCellStyle totalStyle2(XSSFRow row,Integer cell,String value, XSSFWorkbook workbook,int fontsize) {
		XSSFCellStyle cellStyle=workbook.createCellStyle();
		XSSFFont font = setFont(fontsize, workbook);
		font.setBold(true);
		font.setFontName("Calibri");
		font.setFontHeightInPoints((short)11);
		//XSSFColor color=workbook.getCustomPalette().
		cellStyle.setFillForegroundColor(new XSSFColor(new Color(196,188,150)));
		cellStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
		cellStyle.setFont(font);
		cellStyle.setBorderTop(cellstyleborder);
		cellStyle.setBorderBottom(cellstyleborder);
		cellStyle.setBorderLeft(cellstyleborder);
		cellStyle.setBorderRight(cellstyleborder);
		cellStyle.setAlignment(CellStyle.ALIGN_RIGHT); 
		cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		
		cellStyle.setBottomBorderColor(color);
		cellStyle.setTopBorderColor(color);
		cellStyle.setLeftBorderColor(color);
		cellStyle.setRightBorderColor(color);
		
		if (value != null) {
			row.createCell(cell).setCellValue(value);
		} else {
			row.createCell(cell).setCellValue(0);
		}
		row.getCell(cell).setCellStyle(cellStyle);
		return cellStyle;

	}
	

	public XSSFCellStyle genaralStyle(XSSFWorkbook workbook) {
		XSSFCellStyle cellStyle=workbook.createCellStyle();
		XSSFFont font = setFont(9, workbook);
		cellStyle.setWrapText(true);
		cellStyle.setBorderTop(cellstyleborder);
		cellStyle.setBorderBottom(cellstyleborder);
		cellStyle.setBorderLeft(cellstyleborder);
		cellStyle.setBorderRight(cellstyleborder);
		cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
		cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellStyle.setBottomBorderColor(color);
		cellStyle.setTopBorderColor(color);
		cellStyle.setLeftBorderColor(color);
		cellStyle.setRightBorderColor(color);
		font.setFontName("Calibri");
		cellStyle.setFont(font);
		return cellStyle;

	}
	public XSSFCellStyle setvalue(XSSFRow row, Integer cell, String value, XSSFWorkbook workbook) {
		XSSFFont font = setFont(9, workbook);
		
		XSSFCellStyle cellStyle=workbook.createCellStyle();
		font.setFontName("Calibri");
		cellStyle.setFont(font);
		cellStyle.setWrapText(true);
		cellStyle.setBorderTop(cellstyleborder);
		cellStyle.setBorderBottom(cellstyleborder);
		cellStyle.setBorderLeft(cellstyleborder);
		cellStyle.setBorderRight(cellstyleborder);
		cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
		cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		
		cellStyle.setBottomBorderColor(color);
		cellStyle.setTopBorderColor(color);
		cellStyle.setLeftBorderColor(color);
		cellStyle.setRightBorderColor(color);

		if (value != null) {
			row.createCell(cell).setCellValue(value);
		} else {
			row.createCell(cell).setCellValue(0);
		}
		row.getCell(cell).setCellStyle(cellStyle);
		return cellStyle;

	}
	
	
	public XSSFCellStyle setvalue1(XSSFRow row, Integer cell, String value, XSSFWorkbook workbook) {
		XSSFFont font = setFont(9, workbook);
		
		XSSFCellStyle cellStyle=workbook.createCellStyle();
		font.setFontName("Calibri");
		cellStyle.setFont(font);
		cellStyle.setWrapText(true);
		cellStyle.setBorderTop(cellstyleborder);
		cellStyle.setBorderBottom(cellstyleborder);
		cellStyle.setBorderLeft(cellstyleborder);
		cellStyle.setBorderRight(cellstyleborder);
		cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
		cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		
		cellStyle.setBottomBorderColor(color);
		cellStyle.setTopBorderColor(color);
		cellStyle.setLeftBorderColor(color);
		cellStyle.setRightBorderColor(color);

		if (value != null) {
			row.createCell(cell).setCellValue(value);
		} else {
			row.createCell(cell).setCellValue(0);
		}
		row.getCell(cell).setCellStyle(cellStyle);
		return cellStyle;

	}
	
	public XSSFCellStyle SetSubDetails(XSSFRow row, Integer cell, String value, XSSFWorkbook workbook) {
		XSSFFont font = setFont(9, workbook);
		XSSFCellStyle cellStyle=workbook.createCellStyle();
		font.setFontName("Calibri");
		cellStyle.setFont(font);
		cellStyle.setWrapText(true);
		font.setFontName("Calibri");
		cellStyle.setAlignment(CellStyle.ALIGN_LEFT);
		cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		if (value != null) {
			row.createCell(cell).setCellValue(value);
		} else {
			row.createCell(cell).setCellValue(0);
		}
		row.getCell(cell).setCellStyle(cellStyle);
		return cellStyle;
	}

	public XSSFCellStyle setvalue(XSSFRow row, Integer cell, Integer value, XSSFWorkbook workbook) {
		XSSFCellStyle cellStyle=workbook.createCellStyle();
		XSSFFont font = setFont(9, workbook);
		cellStyle.setWrapText(true);
		font.setFontName("Calibri");
		cellStyle.setFont(font);

		cellStyle.setBorderTop(cellstyleborder);
		cellStyle.setBorderBottom(cellstyleborder);
		cellStyle.setBorderLeft(cellstyleborder);
		cellStyle.setBorderRight(cellstyleborder);
		cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
		cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		
		cellStyle.setBottomBorderColor(color);
		cellStyle.setTopBorderColor(color);
		cellStyle.setLeftBorderColor(color);
		cellStyle.setRightBorderColor(color);
		if (value != null) {
			row.createCell(cell).setCellValue(value);
		} else {
			row.createCell(cell).setCellValue(0);
		}
		row.getCell(cell).setCellStyle(cellStyle);
		return cellStyle;

	}

	public XSSFCellStyle setvalue(XSSFRow row, Integer cell, Double value, XSSFWorkbook workbook) {
		XSSFCellStyle cellStyle=workbook.createCellStyle();
		XSSFFont font = setFont(9, workbook);
		cellStyle.setWrapText(true);
		font.setFontName("Calibri");
		cellStyle.setFont(font);

		cellStyle.setBorderTop(cellstyleborder);
		cellStyle.setBorderBottom(cellstyleborder);
		cellStyle.setBorderLeft(cellstyleborder);
		cellStyle.setBorderRight(cellstyleborder);
		cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
		cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);

		cellStyle.setBottomBorderColor(color);
		cellStyle.setTopBorderColor(color);
		cellStyle.setLeftBorderColor(color);
		cellStyle.setRightBorderColor(color);
		if (value != null) {
			row.createCell(cell).setCellValue(value);
		} else {
			row.createCell(cell).setCellValue(0);
		}
		row.getCell(cell).setCellStyle(cellStyle);
		return cellStyle;

	}

	public XSSFCellStyle setvalue(XSSFRow row, Integer cell, Long value, XSSFWorkbook workbook) {
		XSSFCellStyle cellStyle=workbook.createCellStyle();
		XSSFFont font = setFont(9, workbook);
		cellStyle.setWrapText(true);
		font.setFontName("Calibri");
		cellStyle.setFont(font);

		cellStyle.setBorderTop(cellstyleborder);
		cellStyle.setBorderBottom(cellstyleborder);
		cellStyle.setBorderLeft(cellstyleborder);
		cellStyle.setBorderRight(cellstyleborder);
		cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
		cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);

		cellStyle.setBottomBorderColor(color);
		cellStyle.setTopBorderColor(color);
		cellStyle.setLeftBorderColor(color);
		cellStyle.setRightBorderColor(color);
		if (value != null) {
			row.createCell(cell).setCellValue(value);
		} else {
			row.createCell(cell).setCellValue(0);
		}
		row.getCell(cell).setCellStyle(cellStyle);
		return cellStyle;

	}

	public XSSFCellStyle setTitle(XSSFRow row, Integer cell, String value, XSSFWorkbook workbook) {
		XSSFCellStyle cellStyle=workbook.createCellStyle();
		XSSFFont font = setFont(12, workbook);
		cellStyle.setWrapText(true);
		font.setFontName("Calibri");
		font.setBold(true);
		cellStyle.setFont(font);
		cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
		cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		row.createCell(cell).setCellValue(value);
		row.getCell(cell).setCellStyle(cellStyle);
		return cellStyle;

	}

	public XSSFCellStyle setheadding(XSSFRow row, Integer cell, String value, XSSFWorkbook workbook) {
		XSSFCellStyle cellStyle=workbook.createCellStyle();
		XSSFFont font = setFont(9, workbook);
		cellStyle.setWrapText(true);
		font.setFontName("Calibri");
		font.setBold(true);
		cellStyle.setFont(font);

		cellStyle.setBorderTop(cellstyleborder);
		cellStyle.setBorderBottom(cellstyleborder);
		cellStyle.setBorderLeft(cellstyleborder);
		cellStyle.setBorderRight(cellstyleborder);
		cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
		cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellStyle.setBottomBorderColor(color);
		cellStyle.setTopBorderColor(color);
		cellStyle.setLeftBorderColor(color);
		cellStyle.setRightBorderColor(color);

		row.createCell(cell).setCellValue(value);

		row.getCell(cell).setCellStyle(cellStyle);
		return cellStyle;

	}

	public XSSFCellStyle setformula(XSSFRow row, Integer cell, String formula, XSSFWorkbook workbook) {
		XSSFCellStyle cellStyle=workbook.createCellStyle();
		XSSFFont font = setFont(9, workbook);
		cellStyle.setWrapText(true);
		font.setFontName("Calibri");
		cellStyle.setFont(font);

		cellStyle.setBorderTop(cellstyleborder);
		cellStyle.setBorderBottom(cellstyleborder);
		cellStyle.setBorderLeft(cellstyleborder);
		cellStyle.setBorderRight(cellstyleborder);
		cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
		cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);

		cellStyle.setBottomBorderColor(color);
		cellStyle.setTopBorderColor(color);
		cellStyle.setLeftBorderColor(color);
		cellStyle.setRightBorderColor(color);
		row.createCell(cell).setCellFormula(formula);

		row.getCell(cell).setCellStyle(cellStyle);
		return cellStyle;

	}

	public XSSFCellStyle genaralStylewithdot(XSSFWorkbook workbook) {
		XSSFCellStyle cellStyle=workbook.createCellStyle();
		cellStyle.setWrapText(true);
		cellStyle.setBorderTop(cellstyleborder);
		cellStyle.setBorderBottom(cellstyleborder);
		cellStyle.setBorderLeft(cellstyleborder);
		cellStyle.setBorderRight(cellstyleborder);
		cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
		cellStyle.setFillPattern(CellStyle.LESS_DOTS);

		cellStyle.setBottomBorderColor(color);
		cellStyle.setTopBorderColor(color);
		cellStyle.setLeftBorderColor(color);
		cellStyle.setRightBorderColor(color);
		cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		return cellStyle;

	}

	public XSSFCellStyle genaralStyle(int fontsize, XSSFWorkbook workbook) {
		XSSFCellStyle cellStyle=workbook.createCellStyle();
		XSSFFont font = setFont(fontsize, workbook);
		cellStyle.setFont(font);
		cellStyle.setBorderTop(cellstyleborder);
		cellStyle.setBorderBottom(cellstyleborder);
		cellStyle.setBorderLeft(cellstyleborder);
		cellStyle.setBorderRight(cellstyleborder);
		cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
		cellStyle.setBottomBorderColor(color);
		cellStyle.setTopBorderColor(color);
		cellStyle.setLeftBorderColor(color);
		cellStyle.setRightBorderColor(color);
		cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		return cellStyle;
	}

	public XSSFCellStyle genaralStyle(int fontsize, XSSFWorkbook workbook, boolean wraptext) {
		XSSFCellStyle cellStyle=workbook.createCellStyle();
		XSSFFont font = setFont(fontsize, workbook);
		cellStyle.setFont(font);
		cellStyle.setWrapText(wraptext);
		cellStyle.setBorderTop(cellstyleborder);
		cellStyle.setBorderBottom(cellstyleborder);
		cellStyle.setBorderLeft(cellstyleborder);
		cellStyle.setBorderRight(cellstyleborder);
		cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
		cellStyle.setBottomBorderColor(color);
		cellStyle.setTopBorderColor(color);
		cellStyle.setLeftBorderColor(color);
		cellStyle.setRightBorderColor(color);
		cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		return cellStyle;
	}

	public XSSFCellStyle tableHead(XSSFWorkbook workbook, boolean wraptext) {
		XSSFCellStyle cellStyle=workbook.createCellStyle();
		cellStyle.setWrapText(wraptext);
		cellStyle.setBorderTop(cellstyleborder);
		cellStyle.setBorderBottom(cellstyleborder);
		cellStyle.setBorderLeft(cellstyleborder);
		cellStyle.setBorderRight(cellstyleborder);
		cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
		cellStyle.setBottomBorderColor(color);
		cellStyle.setTopBorderColor(color);
		cellStyle.setLeftBorderColor(color);
		cellStyle.setRightBorderColor(color);
		cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		return cellStyle;
	}

	public XSSFCellStyle genaralStyleWithDate(XSSFWorkbook workbook) {
		XSSFCellStyle cellStyle=workbook.createCellStyle();
		cellStyle.setWrapText(true);
		cellStyle.setBorderTop(cellstyleborder);
		cellStyle.setBorderBottom(cellstyleborder);
		cellStyle.setBorderLeft(cellstyleborder);
		cellStyle.setBorderRight(cellstyleborder);
		cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
		cellStyle.setBottomBorderColor(color);
		cellStyle.setTopBorderColor(color);
		cellStyle.setLeftBorderColor(color);
		cellStyle.setRightBorderColor(color);
		cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		CreationHelper createHelper = workbook.getCreationHelper();
		cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("yyyy/mm/dd"));

		return cellStyle;

	}

	public CellStyle genaralStyleWithdotDate(XSSFWorkbook workbook) {
		XSSFCellStyle cellStyle=workbook.createCellStyle();
		cellStyle.setWrapText(true);
		cellStyle.setBorderTop(cellstyleborder);
		cellStyle.setBorderBottom(cellstyleborder);
		cellStyle.setBorderLeft(cellstyleborder);
		cellStyle.setBorderRight(cellstyleborder);
		cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
		cellStyle.setFillPattern(CellStyle.LESS_DOTS);
		cellStyle.setBottomBorderColor(color);
		cellStyle.setTopBorderColor(color);
		cellStyle.setLeftBorderColor(color);
		cellStyle.setRightBorderColor(color);
		cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		CreationHelper createHelper = workbook.getCreationHelper();
		cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("yyyy/mm/dd"));

		return cellStyle;
	}

	public XSSFCellStyle MergedRegionWithStyle(int FirstRow, int LastRow, int FirstCol, int LastCol, int fontsize,
			XSSFSheet worksheet, XSSFWorkbook workbook) {
		XSSFCellStyle cellStyle=workbook.createCellStyle();
		CellRangeAddress cellRangeAddress = new CellRangeAddress(FirstRow, LastRow, FirstCol, LastCol);
		worksheet.addMergedRegion(cellRangeAddress);

		XSSFFont font = setFont(fontsize, workbook);
		font.setBold(true);
	
		RegionUtil.setBorderTop(cellstyleborder, cellRangeAddress, worksheet, workbook);
		RegionUtil.setBorderBottom(cellstyleborder, cellRangeAddress, worksheet, workbook);
		RegionUtil.setBorderLeft(cellstyleborder, cellRangeAddress, worksheet, workbook);
		RegionUtil.setBorderRight(cellstyleborder, cellRangeAddress, worksheet, workbook);
		
		RegionUtil.setBottomBorderColor(color, cellRangeAddress, worksheet, workbook);
		RegionUtil.setTopBorderColor(color, cellRangeAddress, worksheet, workbook);
		RegionUtil.setLeftBorderColor(color, cellRangeAddress, worksheet, workbook);
		RegionUtil.setRightBorderColor(color, cellRangeAddress, worksheet, workbook);
		return cellStyle;
	}

	public POIFSFileSystem Setsecurity(String fileName, String password, XSSFWorkbook workbook)
			throws InvalidFormatException, IOException, GeneralSecurityException {
		FileOutputStream fileOut = new FileOutputStream(fileName);
		workbook.write(fileOut);
		POIFSFileSystem fs = new POIFSFileSystem();
		EncryptionInfo info = new EncryptionInfo(fs, EncryptionMode.agile);

		Encryptor enc = info.getEncryptor();
		enc.confirmPassword(password);

		OPCPackage opc = OPCPackage.open(new File(fileName), PackageAccess.READ_WRITE);
		OutputStream os = enc.getDataStream(fs);
		opc.save(os);
		opc.close();
		return fs;
	}

	public void DownloadNonSecuredDocument(HttpServletResponse response, String fileName, XSSFWorkbook workbook)
			throws IOException {

		response.setContentType("application/vnd.ms-excel");
		response.setHeader("Content-Disposition", "attachment; filename=" + fileName);
		// ...
		// Now populate workbook the usual way.
		// ...
		workbook.write(response.getOutputStream()); // Write workbook to
													// response.
		workbook.close();

	}

	public void DownloadProtectedDocument(POIFSFileSystem poiFileSystem, HttpServletResponse response, String fileName,
			XSSFWorkbook workbook) throws IOException {

		response.setContentType("application/vnd.ms-excel");
		response.setHeader("Content-Disposition", "attachment; filename=" + fileName);
		// ...
		// Now populate workbook the usual way.
		// ...
		poiFileSystem.writeFilesystem(response.getOutputStream()); // Write workbook to
		// response.
		poiFileSystem.close();

	}

	public void SetPicture(XSSFWorkbook workbook, XSSFSheet worksheet, int col1, int col2, int row1, int row2,
			String ImageFileName) throws IOException {
		worksheet.addMergedRegion(new CellRangeAddress(row1, row2, col1, col2));

		File myClass = new File(getClass().getProtectionDomain().getCodeSource().getLocation().getFile());
		File f = new File(myClass.getCanonicalPath().replaceFirst("classes", "") + "report/" + ImageFileName);

		/// picture/////
		InputStream inputStream = new FileInputStream(f);
		// Get the contents of an InputStream as a byte[].
		byte[] bytes = IOUtils.toByteArray(inputStream);
		// Adds a picture to the workbook
		int pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
		// close the input stream
		inputStream.close();

		// Returns an object that handles instantiating concrete classes
		CreationHelper helper = workbook.getCreationHelper();

		// Creates the top-level drawing patriarch.
		Drawing drawing = worksheet.createDrawingPatriarch();

		// Create an anchor that is attached to the worksheet
		ClientAnchor anchor = helper.createClientAnchor();
		// set top-left corner for the image
		anchor.setCol1(col1);
		anchor.setRow1(row1);
		anchor.setCol2(col2);
		anchor.setRow2(row2);

		// Creates a picture

		Picture pict = drawing.createPicture(anchor, pictureIdx);
		// Reset the image to the original size
		pict.resize();
	}

	public File getPath(String fileName) throws IOException {
		File myClass = new File(getClass().getProtectionDomain().getCodeSource().getLocation().getFile());
		File f = new File(myClass.getCanonicalPath().replaceFirst("classes", "") + "report/" + fileName);
		return f;

	}

	public void setheadder(XSSFSheet worksheet, String Stage, String Date) {
		worksheet.getHeader().setLeft(Stage);
		worksheet.getHeader().setRight("Date:" + Date);

	}

	public void setfotter(XSSFSheet worksheet, String Center, String Username) {

		worksheet.getFooter().setLeft("Center:" + Center);
		worksheet.getFooter().setCenter("UserName:" + Username);

	}

	public void setproperty(XSSFWorkbook workbook, String Creator, String Title) {
		POIXMLProperties xmlProps = workbook.getProperties();
		POIXMLProperties.CoreProperties coreProps = xmlProps.getCoreProperties();
		coreProps.setCreator(Creator);
		coreProps.setTitle(Title);

	}
	public  XSSFFont setFont(int val, XSSFWorkbook workbook) {
		XSSFFont style = workbook.createFont();
		style.setFontHeightInPoints((short) val);

		return style;
	}

	

	public XSSFCellStyle MergedRegionWithStyle1(int FirstRow,int LastRow,int FirstCol, int LastCol,int fontsize, XSSFSheet worksheet,XSSFWorkbook workbook) {
		CellRangeAddress cellRangeAddress = new CellRangeAddress(FirstRow, LastRow, FirstCol, LastCol);
		
			
			worksheet.addMergedRegion(cellRangeAddress);
			
			XSSFFont font = ExcelGenarationIn.setFont(fontsize, workbook);
			font.setBold(true);
			XSSFCellStyle cellStyle = workbook.createCellStyle();
			cellStyle.setFont(font);
			
		/*	RegionUtil.setBorderTop(CellStyle.BORDER_MEDIUM, cellRangeAddress, worksheet, workbook);
			RegionUtil.setBorderBottom(CellStyle.BORDER_MEDIUM, cellRangeAddress, worksheet, workbook);
			RegionUtil.setBorderLeft(CellStyle.BORDER_MEDIUM, cellRangeAddress, worksheet, workbook);
			RegionUtil.setBorderRight(CellStyle.BORDER_MEDIUM, cellRangeAddress, worksheet, workbook);*/
			
			cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
			return cellStyle;
		}
	
	public  XSSFCellStyle setMainTitleWithMarged(XSSFRow row,Integer cell,String value,XSSFWorkbook workbook) {
		XSSFFont font = ExcelGenarationIn.setFont(12, workbook);
		XSSFCellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setWrapText(true);
		font.setFontName("Arial");
		font.setBold(true);
		cellStyle.setFont(font);
		/*cellStyle.setBorderTop(CellStyle.BORDER_MEDIUM);
		cellStyle.setBorderBottom(CellStyle.BORDER_MEDIUM);
		cellStyle.setBorderLeft(CellStyle.BORDER_MEDIUM);
		cellStyle.setBorderRight(CellStyle.BORDER_MEDIUM);
		cellStyle.setAlignment(CellStyle.ALIGN_CENTER);*/
		cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		
		cellStyle.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
		cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		
		row.createCell(cell).setCellValue(value);
		row.getCell(cell).setCellStyle(cellStyle);
		return cellStyle;

	}

	public void cellStyle(CellStyle cellStyle2) {
		// TODO Auto-generated method stub
		short a=1;
		cellStyle2.setBorderBottom(a);
		cellStyle2.setBorderLeft(a);
		cellStyle2.setBorderRight(a);
		cellStyle2.setBorderTop(a);
		cellStyle2.setAlignment(CellStyle.ALIGN_CENTER);
		cellStyle2.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		
		cellStyle2.setWrapText(true);
	}

	public void cellStyle2(CellStyle cellStyle2) {
		short a=1;
		cellStyle2.setBorderBottom(a);
		cellStyle2.setBorderLeft(a);
		cellStyle2.setBorderRight(a);
		cellStyle2.setBorderTop(a);
		cellStyle2.setAlignment(CellStyle.ALIGN_RIGHT);
		cellStyle2.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		
		cellStyle2.setWrapText(true);
	}
	
	public XSSFCellStyle setDate(XSSFRow row, Integer cell, String value, XSSFWorkbook workbook) {
		XSSFFont font = setFont(9, workbook);
		
		XSSFCellStyle cellStyle=workbook.createCellStyle();
		font.setFontName("Calibri");
		cellStyle.setFont(font);
		font.setBold(true);
		cellStyle.setWrapText(true);
		cellStyle.setBorderTop(cellstyleborder);
		cellStyle.setBorderBottom(cellstyleborder);
		cellStyle.setBorderLeft(cellstyleborder);
		cellStyle.setBorderRight(cellstyleborder);
		cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
		cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		
		cellStyle.setBottomBorderColor(color);
		cellStyle.setTopBorderColor(color);
		cellStyle.setLeftBorderColor(color);
		cellStyle.setRightBorderColor(color);

		if (value != null) {
			row.createCell(cell).setCellValue(value);
		} else {
			row.createCell(cell).setCellValue(0);
		}
		row.getCell(cell).setCellStyle(cellStyle);
		return cellStyle;

	}
	
}
