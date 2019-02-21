package com.tvd.demo.controller;

import java.util.ArrayList;
import java.util.List;

import javax.transaction.Transaction;

import org.hibernate.Session;
import org.hibernate.SessionFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import com.tvd.demo.dto.Answer;
import com.tvd.demo.dto.Question;

@Controller
public class MappingController {
	
	@Autowired
	private SessionFactory sf;	
	
@RequestMapping("/onetomany")	
public void insert() {
	
	System.out.println("inside insert");
	
	Session ses = sf.openSession();
	org.hibernate.Transaction tx =  ses.beginTransaction();
	
	
	Answer a1 = new Answer();
	a1.setAnsname("java  is Programming Language");
	a1.setPostedby("Games Gosling");

	Answer a2 = new Answer();
	a2.setAnsname("Java is a platform independent object oriented programming language");
	a2.setPostedby("patrickD");
	
	List<Answer> list1 = new ArrayList<>();
	list1.add(a1);
	list1.add(a2);
	
	Question que= new Question();
	que.setQues("What is java?");
	que.setAns(list1);

	ses.persist(que);
	tx.commit();
	ses.close();
}

}
