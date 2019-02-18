package com.tvd.demo.controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.*;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.ModelMap;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.servlet.ModelAndView;

import com.tvd.demo.dto.StudentDTO;
import com.tvd.demo.service.StudentServ;
import com.tvd.demo.utils.ExcelGenaration;

@Controller
public class StudentController {
	
	@Autowired
	private StudentServ service;
	
	@Autowired
	private ExcelGenaration excel;
	
	@RequestMapping("/")
	public ModelAndView home(ModelMap map) {
		ModelAndView mv=new ModelAndView("home");
		List<StudentDTO> students=service.getAllStudents();
		for (StudentDTO s : students) {
			System.err.println("name is "+s.getName());
		}
		mv.addObject("students",students);
		return mv;
	}
	
	@RequestMapping("/saveStudent")
	public String saveStudent(@ModelAttribute StudentDTO dto) {
		System.err.println("inside save "+dto.getName());
		service.saveStudent(dto);
		return "redirect:/";
	}
	
	@RequestMapping("/editStudent")
	@ResponseBody
	public StudentDTO editStudent(@RequestParam int id) {
		StudentDTO student=service.editStudent(id);
		return student;
	}
	
	@RequestMapping("/deleteStudent")
	public String deleteStudent(@RequestParam int id) {
		service.deleteStudent(id);
		return "redirect:/";
	}

	@RequestMapping("/excelrep")
	public void excelReport(HttpServletResponse resp) throws FileNotFoundException, IOException {
		
		/* System.out.println("inside the excel report..."); */
		
		int row=2;
		int cell=0;
		
		Row row1;
		List<StudentDTO> students=service.getAllStudents();
		/*
		 * for (StudentDTO studentDTO : students) {
		 * System.out.println("name--->"+studentDTO.getName()); }
		 */
		XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(excel.getPath("StudentReport.xlsx")));
		
		XSSFSheet worksheet=workbook.getSheet("AAAAAA");
	
		
		  CellStyle cellStyle = worksheet.getWorkbook().createCellStyle();
		  excel.cellStyle(cellStyle); 
		  CellStyle cellStyle2 =worksheet.getWorkbook().createCellStyle(); 
		  excel.cellStyle2(cellStyle2);
		  
		  for (StudentDTO s : students) { 
		 row1 = worksheet.createRow((int) row);
		  
		  Cell cell001 = row1.createCell(0); 
		  cell001.setCellValue(s.getId());
		  cell001.setCellStyle(cellStyle);
		  
		  Cell cell002 = row1.createCell(1); 
		  cell002.setCellValue(s.getName());
		  cell002.setCellStyle(cellStyle);
		  
		  Cell cell003 = row1.createCell(2); 
		  cell003.setCellValue(s.getPlace());
		  cell003.setCellStyle(cellStyle);
		  
		  Cell cell004 = row1.createCell(3); 
		  cell004.setCellValue(s.getEmail());
		  cell004.setCellStyle(cellStyle);
		  
		  row++; }
		 
		 
		
		
		
		excel.DownloadNonSecuredDocument(resp, "StudentReport.xlsx", workbook);
		
		
	}
	
	@RequestMapping("responseExcel")
	public void responseexcel(HttpServletResponse resp) throws IOException {
		
		List<StudentDTO> students=service.getAllStudents();
		
		String filename="../report/sample.xlsx";
		
		PrintWriter out=resp.getWriter();
		out.write("<Table>");
		out.write("<tr>");
		out.write("<th>SLNO</th>");
		out.write("<th>NAME</th>");
		out.write("<th>PLACE</th>");
		out.write("<th>EMAIL</th>");
		out.write("</tr>");
		for (StudentDTO studentDTO : students) {
			
			out.write("<tr>");
			out.write("<td>"+studentDTO.getId()+"</td>");
			out.write("<td>"+studentDTO.getName()+"</td>");
			out.write("<td>"+studentDTO.getPlace()+"</td>");
			out.write("<td>"+studentDTO.getEmail()+"</td>");
			out.write("</tr>");
		}
		out.write("</Table>");
		resp.setHeader("Content-Disposition","attachment;filename=\""+filename+"\"");
 		out.close();
		
		
	}
	
	
}
