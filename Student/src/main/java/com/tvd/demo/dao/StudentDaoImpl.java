package com.tvd.demo.dao;

import java.util.List;

import org.hibernate.HibernateException;
import org.hibernate.SessionFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;
import org.springframework.transaction.annotation.Transactional;

import com.tvd.demo.dto.StudentDTO;

@Component
@Transactional
public class StudentDaoImpl implements StudentDao{
	
	@Autowired
	private SessionFactory sf;

	@Override
	public void saveStudent(StudentDTO dto) {
		sf.getCurrentSession().saveOrUpdate(dto);
	}

	@Override
	public List<StudentDTO> getAllStudents() {
		try {
			return  sf.getCurrentSession().createQuery("FROM StudentDTO",StudentDTO.class).getResultList();
		} catch (HibernateException e) {
			e.printStackTrace();
			return null;
		}
		
	}

	@Override
	public StudentDTO editStudent(int id) {
		try {
			return sf.getCurrentSession().get(StudentDTO.class, id);
		} catch (HibernateException e) {
			e.printStackTrace();
			return null;
		}
	}

	@Override
	public void deleteStudent(int id) {
		sf.getCurrentSession().delete(sf.getCurrentSession().get(StudentDTO.class, id));
	}

}
