package com.docx.test.model;

import java.util.List;

public class Phone {

	private String phoneNo;
	private String provider;
	private List<Children> downlines;

	public String getPhoneNo() {
		return phoneNo;
	}

	public void setPhoneNo(String phoneNo) {
		this.phoneNo = phoneNo;
	}

	public String getProvider() {
		return provider;
	}

	public void setProvider(String provider) {
		this.provider = provider;
	}

	public List<Children> getDownlines() {
		return downlines;
	}

	public void setDownlines(List<Children> downlines) {
		this.downlines = downlines;
	}

}
