package com.tvd.demo.dao;

import java.util.List;

import com.tvd.demo.dto.StudentDTO;

public interface StudentDao {

	void saveStudent(StudentDTO dto);

	List<StudentDTO> getAllStudents();

	StudentDTO editStudent(int id);

	void deleteStudent(int id);

}
