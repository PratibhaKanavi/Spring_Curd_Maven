package com.tvd.demo.service;

import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import com.tvd.demo.dao.StudentDao;
import com.tvd.demo.dto.StudentDTO;

@Component
public class StudentServImpl implements StudentServ {
	
	@Autowired
	private StudentDao dao;

	@Override
	public void saveStudent(StudentDTO dto) {
		dao.saveStudent(dto);
	}

	@Override
	public List<StudentDTO> getAllStudents() {
		return dao.getAllStudents();
	}

	@Override
	public StudentDTO editStudent(int id) {
		return dao.editStudent(id);
	}

	@Override
	public void deleteStudent(int id) {
		 dao.deleteStudent(id);
	}
	
	

}
