package com.infogen.restcontroller;

import java.awt.Color;
import java.awt.Rectangle;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.sl.usermodel.TableCell.BorderEdge;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.sl.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import com.infogen.dto.LeanDashboardDto;
import com.infogen.dto.LeanProjectsDto;

import jakarta.servlet.http.HttpServletResponse;

@RestController
public class PptController {

	public List<LeanDashboardDto> getLeanDashboardDataFromExcel() throws IOException {
		String filelocation = "report//New Microsoft Excel Worksheet.xlsx";
		FileInputStream fis = new FileInputStream(new File(filelocation));
		Workbook workbook = new XSSFWorkbook(fis);
		Sheet sheet = workbook.getSheetAt(1);
		DataFormatter dataFormatter = new DataFormatter();
		List<LeanDashboardDto> leanDashboardDtos = new ArrayList<>();
		int rowCount = sheet.getLastRowNum() + 1;
		for (int k = 2; k <= rowCount; k++) {
			Row row = sheet.getRow(k);
			LeanDashboardDto leanDashboardDto = new LeanDashboardDto();
			if (row != null) {
				for (Cell cell : row) {
					String cellValue = dataFormatter.formatCellValue(cell);
					if (cell.getColumnIndex() == 1) {
						if (cellValue != null && !cellValue.isEmpty()) {
							leanDashboardDto.setMetric(cellValue);
							;
						}
					} else if (cell.getColumnIndex() == 2) {
						if (cellValue != null && !cellValue.isEmpty()) {
							leanDashboardDto.setUom(cellValue);
						}
					} else if (cell.getColumnIndex() == 3) {
						if (cellValue != null && !cellValue.isEmpty()) {
							leanDashboardDto.setPrevious(Float.parseFloat(cellValue));
						}
					} else if (cell.getColumnIndex() == 4) {
						if (cellValue != null && !cellValue.isEmpty()) {
							leanDashboardDto.setCurrent(Float.parseFloat(cellValue));
						}
					} else if (cell.getColumnIndex() == 5) {
						if (cellValue != null && !cellValue.isEmpty()) {
							leanDashboardDto.setTarget(Float.parseFloat(cellValue));
						}
					} else if (cell.getColumnIndex() == 6) {
						if (cellValue != null && !cellValue.isEmpty()) {
							leanDashboardDto.setA(Float.parseFloat(cellValue));
						}
					} else if (cell.getColumnIndex() == 18) {
						if (cellValue != null && !cellValue.isEmpty()) {
							leanDashboardDto.setYtd(Float.parseFloat(cellValue));
						}
					}
				}
				leanDashboardDtos.add(leanDashboardDto);
			}
		}
		return leanDashboardDtos;
	}

	@GetMapping("/ppt")
	public void generatePresentation(HttpServletResponse response) throws Exception {
		String fileLocation = "report//presentation.pptx";
		// Create presentation
		XMLSlideShow samplePPT = new XMLSlideShow();

		XSLFSlideMaster defaultSlideMaster = samplePPT.getSlideMasters().get(0);
		samplePPT.setPageSize(new java.awt.Dimension(1800, 1100));
		// Retrieving the slide layout
		XSLFSlideLayout defaultLayout = defaultSlideMaster.getLayout(SlideLayout.TITLE_ONLY);

		// Creating the 1st slide
		XSLFSlide sampleSlide1 = samplePPT.createSlide(defaultLayout);
		XSLFTextShape sampleTitle = sampleSlide1.getPlaceholder(0);

		sampleTitle.clearText();
		sampleTitle.setAnchor(new Rectangle(400, 20, 800, 100));
		XSLFTextParagraph paragraph = sampleTitle.addNewTextParagraph();
		paragraph.setTextAlign(TextAlign.CENTER);
		XSLFTextRun r = paragraph.addNewTextRun();
		r.setText("Lean Projects Status - TPI");
		r.setFontColor(Color.RED);
		r.setFontFamily("TimesNewRoman");
		r.setFontSize(40.);
		r.setUnderlined(true);
		r.setBold(true);
		
		creationOfTable(sampleSlide1);

		XSLFSlide sampleSlide2 = samplePPT.createSlide(defaultLayout);
		XSLFTextShape sampleTitle1 = sampleSlide2.getPlaceholder(0);

		sampleTitle1.clearText();
		XSLFTextParagraph paragraph1 = sampleTitle1.addNewTextParagraph();
		paragraph1.setTextAlign(TextAlign.CENTER);
		XSLFTextRun r1 = paragraph1.addNewTextRun();
		r1.setText("Lean Dashboard - TPI");
		r1.setFontColor(Color.RED);
		r1.setFontFamily("TimesNewRoman");
		r1.setFontSize(36.);
		r1.setUnderlined(true);
		r1.setBold(true);
		
		creationOfTable(sampleSlide2);

		FileOutputStream outputStream = new FileOutputStream(fileLocation);
		samplePPT.write(outputStream);
		
		outputStream.close();
		// Closing presentation
		samplePPT.close();
	}

	private void creationOfTable(XSLFSlide slide) throws IOException {

		List<LeanProjectsDto> leanProjectsDtos = readProjectExcel();
		List<LeanDashboardDto> leanDashboardDtos = getLeanDashboardDataFromExcel();

		if (slide.getSlideNumber() == 1) {
			XSLFTable table = slide.createTable();
			table.setAnchor(new Rectangle(50, 120, 400, 400));
			
			List<String> headers = new ArrayList<>();
			headers.add("Sl.No");
			headers.add("Project Title");
			headers.add("Item");
			headers.add("UOM");
			headers.add("Start");
			headers.add("Target");
			headers.add("Mar'23");
			headers.add("Apr'23");
			XSLFTableRow headerRow = table.addRow();
			headerRow.setHeight(50);

			int r1 = 0;
			int r2 = 0;
			int r3 = 0;
			int r4 = 0;
			
			for (int i = 0; i < headers.size(); i++) {
				XSLFTableCell th = headerRow.addCell();
				th.setBorderColor(BorderEdge.top, Color.BLACK);
				th.setBorderColor(BorderEdge.right, Color.BLACK);
				th.setBorderColor(BorderEdge.bottom, Color.BLACK);
				th.setBorderColor(BorderEdge.left, Color.BLACK);
				th.setVerticalAlignment(VerticalAlignment.MIDDLE);
				XSLFTextParagraph p = th.addNewTextParagraph();
				p.setTextAlign(TextAlign.CENTER);
				XSLFTextRun r = p.addNewTextRun();
				r.setText(headers.get(i));
				r.setBold(true);
				r.setFontColor(Color.white);
				r.setFontFamily("TimesNewRoman");
				r.setFontSize(24.);
				th.setFillColor(new Color(79, 129, 189));
				if (i == 0) {
					table.setColumnWidth(i, 100);
				} else if (i == 1) {
					table.setColumnWidth(i, 350);
				} else if (i == 2) {
					table.setColumnWidth(i, 250);
				} else if (i == 3) {
					table.setColumnWidth(i, 200);
				} else {
					table.setColumnWidth(i, 150);
				}
			}
			int numColumns = 8;
			int no = 1;
			int row = 1;
			XSLFTableRow tr;
			for (LeanProjectsDto leanProjectsDto : leanProjectsDtos) {
				tr = table.addRow();
				tr.setHeight(35);
				for (int i = 0; i < numColumns; i++) {
					XSLFTableCell cell = tr.addCell();
					cell.setBorderColor(BorderEdge.top, Color.BLACK);
					cell.setBorderColor(BorderEdge.right, Color.BLACK);
					cell.setBorderColor(BorderEdge.bottom, Color.BLACK);
					cell.setBorderColor(BorderEdge.left, Color.BLACK);
					cell.setVerticalAlignment(VerticalAlignment.MIDDLE);
					XSLFTextParagraph p = cell.addNewTextParagraph();
					p.setTextAlign(TextAlign.LEFT);
					XSLFTextRun r = p.addNewTextRun();
					r.setFontFamily("TimesNewRoman");
					r.setFontSize(22.);

					if (i == 0) {
						if (leanProjectsDto.getProjectName() != null && !leanProjectsDto.getProjectName().isEmpty()) {
							r.setText(String.valueOf(no++));
							if (r1 == 0) {
								r1++;
							} else {
								r1 = r2;
								r1++;
							}
							r2 = r1;
						} else {
							r.setText(String.valueOf(0));
							r2++;
							table.mergeCells(r1, r2, i, i);
						}
					} else if (i == 1) {
						r.setText(leanProjectsDto.getProjectName());
						table.mergeCells(r1, r2, i, i);
					} else if (i == 2) {
						r.setText(leanProjectsDto.getItemName());
						setFillColorForCell(row, cell);
						if (leanProjectsDto.getItemName() != null && !leanProjectsDto.getItemName().isEmpty()) {
							if (r3 == 0) {
								r3++;
							} else {
								r3 = r4;
								r3++;
							}
							r4 = r3;
						} else {
							r4++;
							table.mergeCells(r3, r4, i, i);
						}
						
					} else if (i == 3) {
						r.setText(leanProjectsDto.getUom());
						setFillColorForCell(row, cell);
					} else if (i == 4) {
						r.setText(String.valueOf(leanProjectsDto.getStart()));
						setFillColorForCell(row, cell);
					} else if (i == 5) {
						r.setText(String.valueOf(leanProjectsDto.getTarget()));
						setFillColorForCell(row, cell);
					} else if (i == 6) {
						r.setText(String.valueOf(leanProjectsDto.getPreviousData()));
						setFillColorForCell(row, cell);
					} else if (i == 7) {
						r.setText(String.valueOf(leanProjectsDto.getCurrentData()));
						setFillColorForCell(row, cell);
					}
				}
				row++;
			}
			
//			table.mergeCells(1, 3, 0, 0);
//			table.mergeCells(4, 6, 0, 0);
//			table.mergeCells(7, 12, 0, 0);
//			table.mergeCells(13, 19, 0, 0);
//			table.mergeCells(20, 25, 0, 0);
//			table.mergeCells(1, 3, 1, 1);
//			table.mergeCells(4, 6, 1, 1);
//			table.mergeCells(7, 12, 1, 1);
//			table.mergeCells(13, 19, 1, 1);
//			table.mergeCells(20, 25, 1, 1);
//			table.mergeCells(7, 9, 2, 2);
//			table.mergeCells(17, 18, 2, 2);
//			table.mergeCells(24, 25, 2, 2);
		} else if (slide.getSlideNumber() == 2) {
			XSLFTable table = slide.createTable();
			table.setAnchor(new Rectangle(50, 200, 400, 400));
			
			List<String> headers = new ArrayList<>();
			headers.add("Metrics");
			headers.add("UOM");
			headers.add("2021-22 (A)");
			headers.add("2022-23 (A)");
			headers.add("Target 23-24");
			headers.add("A");
			headers.add("YTD");
			headers.add("%Imp. from 22-23");

			XSLFTableRow headerRow = table.addRow();
			headerRow.setHeight(80);

			for (int i = 0; i < headers.size(); i++) {
				XSLFTableCell th = headerRow.addCell();
				th.setBorderColor(BorderEdge.top, Color.BLACK);
				th.setBorderColor(BorderEdge.right, Color.BLACK);
				th.setBorderColor(BorderEdge.bottom, Color.BLACK);
				th.setBorderColor(BorderEdge.left, Color.BLACK);
				th.setVerticalAlignment(VerticalAlignment.MIDDLE);
				XSLFTextParagraph p = th.addNewTextParagraph();
				p.setTextAlign(TextAlign.CENTER);
				XSLFTextRun r = p.addNewTextRun();
				r.setText(headers.get(i));
				r.setBold(true);
				r.setFontColor(Color.BLACK);
				r.setFontFamily("TimesNewRoman");
				r.setFontSize(28.);
				th.setFillColor(Color.gray);
				if (i == 0) {
					table.setColumnWidth(i, 250);
				} else {
					table.setColumnWidth(i, 150);
				}
			}

			int numColumns = 8;
			int r1 = 0;
			int r2 = 1;
			XSLFTableRow tr;
			for (LeanDashboardDto leanDashboardDto : leanDashboardDtos) {
				tr = table.addRow();
				tr.setHeight(80);
				for (int i = 0; i < numColumns; i++) {
					XSLFTableCell cell = tr.addCell();
					cell.setBorderColor(BorderEdge.top, Color.BLACK);
					cell.setBorderColor(BorderEdge.right, Color.BLACK);
					cell.setBorderColor(BorderEdge.bottom, Color.BLACK);
					cell.setBorderColor(BorderEdge.left, Color.BLACK);
					cell.setVerticalAlignment(VerticalAlignment.MIDDLE);
					XSLFTextParagraph p = cell.addNewTextParagraph();
					p.setTextAlign(TextAlign.CENTER);
					XSLFTextRun r = p.addNewTextRun();
					r.setFontFamily("TimesNewRoman");
					r.setFontSize(24.);

					if (i == 0) {
						r.setText(leanDashboardDto.getMetric());
					} else if (i == 1) {
						r.setText(leanDashboardDto.getUom());
						if (leanDashboardDto.getUom() != null && !leanDashboardDto.getUom().isEmpty()) {
							if (r1 == 0) {
								r1++;
							} else {
								r1 = r2;
								r1++;
							}
							r2 = r1;
						} else {
							r2++;
							table.mergeCells(r1, r2, i, i);
						}
					} else if (i == 2) {
						r.setText(String.valueOf(leanDashboardDto.getPrevious()));
					} else if (i == 3) {
						r.setText(String.valueOf(leanDashboardDto.getCurrent()));
					} else if (i == 4) {
						r.setText(String.valueOf(leanDashboardDto.getTarget()));
					} else if (i == 5) {
						r.setText(String.valueOf(leanDashboardDto.getA()));
					} else if (i == 6) {
						r.setText(String.valueOf(leanDashboardDto.getYtd()));
					} else if (i == 7) {
						r.setText(String.valueOf((leanDashboardDto.getYtd() - leanDashboardDto.getCurrent())
								/ leanDashboardDto.getCurrent()));
						cell.setFillColor(new Color(50, 204, 140));
					}
				}
			}
//			table.mergeCells(3, 4, 1, 1);
//			table.mergeCells(5, 6, 1, 1);
		}
	}
	
	public void setFillColorForCell(int row, XSLFTableCell cell) {
		if (row == 1) {
			cell.setFillColor(new Color(210 , 180, 100));
		} else if (row == 2) {
			cell.setFillColor(new Color(210 , 180, 100));
		} else if (row == 5) {
			cell.setFillColor(new Color(210 , 180, 100));
		} else if (row == 7) {
			cell.setFillColor(new Color(210 , 180, 100));
		} else if (row == 8) {
			cell.setFillColor(new Color(210 , 180, 100));
		} else if (row == 9) {
			cell.setFillColor(new Color(210 , 180, 100));
		} else if (row == 10) {
			cell.setFillColor(new Color(210 , 180, 100));
		} else if (row == 14) {
			cell.setFillColor(new Color(210 , 180, 100));
		} else if (row == 15) {
			cell.setFillColor(new Color(210 , 180, 100));
		} else if (row == 16) {
			cell.setFillColor(new Color(210 , 180, 100));
		} else if (row == 21) {
			cell.setFillColor(new Color(210 , 180, 100));
		} else if (row == 23) {
			cell.setFillColor(new Color(210 , 180, 100));
		} else if (row == 24) {
			cell.setFillColor(new Color(210 , 180, 100));
		} else if (row == 25) {
			cell.setFillColor(new Color(210 , 180, 100));
		}
	}

	public List<LeanProjectsDto> readProjectExcel() throws IOException {
		String filelocation = "report//New Microsoft Excel Worksheet.xlsx";
		FileInputStream fis = new FileInputStream(new File(filelocation));
		Workbook workbook = new XSSFWorkbook(fis);
		Sheet sheet = workbook.getSheetAt(0);
		DataFormatter dataFormatter = new DataFormatter();
		List<LeanProjectsDto> leanProjects = new ArrayList<>();
		int rowCount = sheet.getLastRowNum() + 1;
		for (int k = 2; k <= rowCount; k++) {
			Row row = sheet.getRow(k);
			LeanProjectsDto leanProjectsDto = new LeanProjectsDto();
			if (row != null) {
				for (Cell cell : row) {
					String cellValue = dataFormatter.formatCellValue(cell);
					if (cell.getColumnIndex() == 4) {
						if (cellValue != null && !cellValue.isEmpty()) {
							leanProjectsDto.setProjectName(cellValue);
						}
					} else if (cell.getColumnIndex() == 5) {
						if (cellValue != null && !cellValue.isEmpty()) {
							leanProjectsDto.setItemName(cellValue);
						}
					} else if (cell.getColumnIndex() == 6) {
						if (cellValue != null && !cellValue.isEmpty()) {
							leanProjectsDto.setUom(cellValue);
						}
					} else if (cell.getColumnIndex() == 7) {
						if (cellValue != null && !cellValue.isEmpty()) {
							leanProjectsDto.setStart(Float.parseFloat(cellValue));
						}
					} else if (cell.getColumnIndex() == 8) {
						if (cellValue != null && !cellValue.isEmpty()) {
							leanProjectsDto.setTarget(Float.parseFloat(cellValue));
						}
					} else if (cell.getColumnIndex() == 20) {
						if (cellValue != null && !cellValue.isEmpty()) {
							leanProjectsDto.setPreviousData(Double.parseDouble(cellValue));
						}
					} else if (cell.getColumnIndex() == 21) {
						if (cellValue != null && !cellValue.isEmpty()) {
							leanProjectsDto.setCurrentData(Double.parseDouble(cellValue));
						}
					}
				}
				leanProjects.add(leanProjectsDto);
			}
		}
		workbook.close();
		return leanProjects;
	}
}
