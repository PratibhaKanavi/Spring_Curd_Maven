package com.tvd.demo.dto;

import java.util.List;

import javax.persistence.CascadeType;
import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;
import javax.persistence.JoinColumn;
import javax.persistence.OneToMany;
import javax.persistence.OrderColumn;
import javax.persistence.Table;

@Entity
@Table(name="Question")
public class Question {
	@Id
	@GeneratedValue(strategy=GenerationType.TABLE)
	private int id;
	private String ques;
	
	@OneToMany(cascade=CascadeType.ALL)
	@JoinColumn(name="qid")
	@OrderColumn(name="type")  
	private List<Answer> ans;

	public int getId() {
		return id;
	}

	public void setId(int id) {
		this.id = id;
	}

	public String getQues() {
		return ques;
	}

	public void setQues(String ques) {
		this.ques = ques;
	}

	public List<Answer> getAns() {
		return ans;
	}

	public void setAns(List<Answer> ans) {
		this.ans = ans;
	}
	
	
	
	
	

}
