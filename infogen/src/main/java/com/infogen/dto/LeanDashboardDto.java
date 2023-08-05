package com.infogen.dto;

import lombok.Data;

@Data
public class LeanDashboardDto {

	private String metric;
	private String uom;
	private float previous;
	private float current;
	private float target;
	private float a;
	private float ytd;
	private float percentage;
}
