package com.tvd.demo.service;

import java.util.List;

import com.tvd.demo.dto.StudentDTO;

public interface StudentServ {

	void saveStudent(StudentDTO dto);

	List<StudentDTO> getAllStudents();

	StudentDTO editStudent(int id);

	void deleteStudent(int id);

}
