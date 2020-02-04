package com.deloitte.exceltojson.pojo;

public class MetaDatCode {

	public MetaDatCode (String code, String color, String description)
	{
		this.code = code;
		this.description = description;
		this.color = color;
	}
	public String code;
	public String getCode() {
		return code;
	}
	public void setCode(String code) {
		this.code = code;
	}
	public String getColor() {
		return color;
	}
	public void setColor(String color) {
		this.color = color;
	}
	public String getDescription() {
		return description;
	}
	public void setDescription(String description) {
		this.description = description;
	}
	public String color;
	public String description;
}
