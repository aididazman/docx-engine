package com.docx.test.model;

import java.util.List;

public class User {

	private String name;
	private int age;
	private List<Phone> phones;
	private List<Children> childs;

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public int getAge() {
		return age;
	}

	public void setAge(int age) {
		this.age = age;
	}

	public List<Phone> getPhones() {
		return phones;
	}

	public void setPhones(List<Phone> phones) {
		this.phones = phones;
	}

	public List<Children> getChilds() {
		return childs;
	}

	public void setChilds(List<Children> childs) {
		this.childs = childs;
	}

}
