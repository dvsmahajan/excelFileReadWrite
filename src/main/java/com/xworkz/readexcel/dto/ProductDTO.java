package com.xworkz.readexcel.dto;

public class ProductDTO {
	
	private Integer id;
	private String name;
	private String type;
	private Double cost;
	public Integer getId() {
		return id;
	}
	public void setId(Integer id) {
		this.id = id;
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public String getType() {
		return type;
	}
	public void setType(String type) {
		this.type = type;
	}
	public Double getCost() {
		return cost;
	}
	public void setCost(Double cost) {
		this.cost = cost;
	}
	public ProductDTO(Integer id, String name, String type, Double cost) {
		super();
		this.id = id;
		this.name = name;
		this.type = type;
		this.cost = cost;
	}
	
	public ProductDTO() {
		
	}
	
	@Override
	public String toString() {
		return "ProductDTO [id=" + id + ", name=" + name + ", type=" + type + ", cost=" + cost + "]";
	}

	
	
}
