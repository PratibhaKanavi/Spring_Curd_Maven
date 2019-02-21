package com.tvd.demo.dto;

import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;
import javax.persistence.Table;

@Entity
@Table(name="Answer")
public class Answer {

@Id
@GeneratedValue(strategy=GenerationType.TABLE)

private int id;

private String ansname;
private String postedby;
public int getId() {
	return id;
}
public void setId(int id) {
	this.id = id;
}
public String getAnsname() {
	return ansname;
}
public void setAnsname(String ansname) {
	this.ansname = ansname;
}
public String getPostedby() {
	return postedby;
}
public void setPostedby(String postedby) {
	this.postedby = postedby;
}


}
