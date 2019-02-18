package com.tvd.demo.dto;

import javax.persistence.*;

import org.hibernate.annotations.GenericGenerator;

@Entity
public class StudentDTO {
	
	@Id
	@GenericGenerator(name="id",strategy="increment")
	@GeneratedValue(generator="id")
	private int id;
	
	private String name;
	private String place;
	private String email;
	public int getId() {
		return id;
	}
	public void setId(int id) {
		this.id = id;
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public String getPlace() {
		return place;
	}
	public void setPlace(String place) {
		this.place = place;
	}
	public String getEmail() {
		return email;
	}
	public void setEmail(String email) {
		this.email = email;
	}
	
	
	
	
	

}
