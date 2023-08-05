package com.infogen.dto;

import lombok.Data;

@Data
public class LeanProjectsDto {
	
	private String projectName;
	private String itemName;
	private String uom;
	private float start;
	private float target;
	private double previousData;
	private double currentData;
}
