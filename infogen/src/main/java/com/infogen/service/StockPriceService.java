package com.infogen.service;

import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.Rectangle;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.concurrent.TimeUnit;

import javax.imageio.ImageIO;

import org.apache.poi.sl.usermodel.VerticalAlignment;
import org.apache.poi.sl.usermodel.TableCell.BorderEdge;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
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
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import com.infogen.entity.StockPrice;
import com.infogen.repository.StockPriceRepository;
import com.xuggle.mediatool.IMediaWriter;
import com.xuggle.mediatool.ToolFactory;
import com.xuggle.xuggler.ICodec;

@Service
public class StockPriceService {

	@Autowired(required = false)
	private StockPriceRepository stockPriceRepository;

	public String generatePptForStockPrice(String spSymbol, String spInstrument) throws IOException {

		StockPrice stockPrice = stockPriceRepository.findBySpsymbolAndSpinstrument(spSymbol, spInstrument);
		if (stockPrice != null) {
			String fileLocation = "report//stockprice.pptx";
			File file = new File(fileLocation);
			FileInputStream is = new FileInputStream(file);
			XMLSlideShow samplePPT = new XMLSlideShow(is);

			XSLFSlideMaster defaultSlideMaster = samplePPT.getSlideMasters().get(0);
			samplePPT.setPageSize(new java.awt.Dimension(1500, 1000));
			XSLFSlideLayout defaultLayout = defaultSlideMaster.getLayout(SlideLayout.TITLE_ONLY);
			XSLFSlide sampleSlide1 = samplePPT.createSlide(defaultLayout);

			setTitleForPpt(spSymbol, spInstrument, sampleSlide1);

			creationOfTable(sampleSlide1, stockPrice);

			FileOutputStream outputStream = new FileOutputStream(fileLocation);
			samplePPT.write(outputStream);

			outputStream.close();
			// Closing presentation
			samplePPT.close();
			convertPptToImages();
			convertImageToVideo();
		}
		return "success";
	}

	private void setTitleForPpt(String spSymbol, String spInstrument, XSLFSlide slide) {
		XSLFTextShape sampleTitle = slide.getPlaceholder(0);
		sampleTitle.clearText();
		sampleTitle.setAnchor(new Rectangle(400, 20, 800, 100));
		XSLFTextParagraph paragraph = sampleTitle.addNewTextParagraph();
		paragraph.setTextAlign(TextAlign.CENTER);
		XSLFTextRun r = paragraph.addNewTextRun();
		r.setText(spInstrument + "-" + spSymbol);
		r.setFontColor(Color.RED);
		r.setFontFamily("TimesNewRoman");
		r.setFontSize(40.);
		r.setUnderlined(true);
		r.setBold(true);
	}

	private void creationOfTable(XSLFSlide slide, StockPrice stockPrice) throws IOException {
		XSLFTable table = slide.createTable();
		table.setAnchor(new Rectangle(50, 120, 400, 400));

		List<String> headers = new ArrayList<>();
		headers.add("Duration");
		headers.add("Open");
		headers.add("Low");
		headers.add("High");
		headers.add("Close");
		XSLFTableRow headerRow = table.addRow();
		headerRow.setHeight(50);

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
				table.setColumnWidth(i, 250);
			} else if (i == 1) {
				table.setColumnWidth(i, 250);
			} else if (i == 2) {
				table.setColumnWidth(i, 250);
			} else if (i == 3) {
				table.setColumnWidth(i, 250);
			} else {
				table.setColumnWidth(i, 250);
			}
		}

		int numColumns = 5;
		int row = 1;
		XSLFTableRow tr;

		for (int i = 0; i < 8; i++) {
			tr = table.addRow();
			tr.setHeight(35);
			for (int j = 0; j < numColumns; j++) {
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

				if (j == 0) {
					if (i == 0) {
						r.setText("To Day");
					} else if (i == 1) {
						r.setText("This Week");
					} else if (i == 2) {
						r.setText("This Month");
					} else if (i == 3) {
						r.setText("This Year");
					} else if (i == 4) {
						r.setText("Pre Day");
					} else if (i == 5) {
						r.setText("Pre Week");
					} else if (i == 6) {
						r.setText("Pre Month");
					} else if (i == 7) {
						r.setText("Pre Year");
					}
				} else if (j == 1) {
					if (i == 0) {
						r.setText(String.valueOf(stockPrice.getTdopen()));
					} else if (i == 1) {
						r.setText(String.valueOf(stockPrice.getTwopen()));
					} else if (i == 2) {
						r.setText(String.valueOf(stockPrice.getTmopen()));
					} else if (i == 3) {
						r.setText(String.valueOf(stockPrice.getTyopen()));
					} else if (i == 4) {
						r.setText(String.valueOf(stockPrice.getPdopen()));
					} else if (i == 5) {
						r.setText(String.valueOf(stockPrice.getPwopen()));
					} else if (i == 6) {
						r.setText(String.valueOf(stockPrice.getPmopen()));
					} else if (i == 7) {
						r.setText(String.valueOf(stockPrice.getPyopen()));
					}
				} else if (j == 2) {
					if (i == 0) {
						r.setText(String.valueOf(stockPrice.getTdlow()));
					} else if (i == 1) {
						r.setText(String.valueOf(stockPrice.getTwlow()));
					} else if (i == 2) {
						r.setText(String.valueOf(stockPrice.getTmlow()));
					} else if (i == 3) {
						r.setText(String.valueOf(stockPrice.getTylow()));
					} else if (i == 4) {
						r.setText(String.valueOf(stockPrice.getPdlow()));
					} else if (i == 5) {
						r.setText(String.valueOf(stockPrice.getPwlow()));
					} else if (i == 6) {
						r.setText(String.valueOf(stockPrice.getPmlow()));
					} else if (i == 7) {
						r.setText(String.valueOf(stockPrice.getPylow()));
					}
				} else if (j == 3) {
					if (i == 0) {
						r.setText(String.valueOf(stockPrice.getTdhigh()));
					} else if (i == 1) {
						r.setText(String.valueOf(stockPrice.getTwhigh()));
					} else if (i == 2) {
						r.setText(String.valueOf(stockPrice.getTmhigh()));
					} else if (i == 3) {
						r.setText(String.valueOf(stockPrice.getTyhigh()));
					} else if (i == 4) {
						r.setText(String.valueOf(stockPrice.getPdhigh()));
					} else if (i == 5) {
						r.setText(String.valueOf(stockPrice.getPwhigh()));
					} else if (i == 6) {
						r.setText(String.valueOf(stockPrice.getPmhigh()));
					} else if (i == 7) {
						r.setText(String.valueOf(stockPrice.getPyhigh()));
					}
				} else if (j == 4) {
					if (i == 0) {
						r.setText(String.valueOf(stockPrice.getTdclose()));
					} else if (i == 1) {
						r.setText(String.valueOf(stockPrice.getTwclose()));
					} else if (i == 2) {
						r.setText(String.valueOf(stockPrice.getTmclose()));
					} else if (i == 3) {
						r.setText(String.valueOf(stockPrice.getTyclose()));
					} else if (i == 4) {
						r.setText(String.valueOf(stockPrice.getPdclose()));
					} else if (i == 5) {
						r.setText(String.valueOf(stockPrice.getPwclose()));
					} else if (i == 6) {
						r.setText(String.valueOf(stockPrice.getPmclose()));
					} else if (i == 7) {
						r.setText(String.valueOf(stockPrice.getPyclose()));
					}
				}
			}
			row++;
		}
	}

	public void convertPptToImages() throws FileNotFoundException, IOException {
		String pptFilePath = "report//stockprice.pptx";
		String outputDir = "report//images//";
		XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(pptFilePath));
		List<XSLFSlide> slides = ppt.getSlides();

		for (int i = 0; i < slides.size(); i++) {
			BufferedImage image = new BufferedImage(ppt.getPageSize().width, ppt.getPageSize().height,
					BufferedImage.TYPE_INT_RGB);
			Graphics2D graphics = image.createGraphics();
			slides.get(i).draw(graphics);
			graphics.dispose();

			String imageFileName = outputDir + File.separator + "slide" + (i + 1) + ".png";
			ImageIO.write(image, "png", new FileOutputStream(imageFileName));
		}
		convertImageToVideo();
	}
	
	public void convertImageToVideo() throws IOException {
		String outputVideoPath = "report//video//stockpriceVdo.mp4";
		String imagesDir = "report//images//";
		int frameRate = 1;
		
		 IMediaWriter writer = ToolFactory.makeWriter(outputVideoPath);
	        writer.addVideoStream(0, 0, ICodec.ID.CODEC_ID_MPEG4, 1920, 1080);
	        
	        File[] images = new File(imagesDir).listFiles((dir, name) -> name.endsWith(".png"));
	        Arrays.sort(images);
	        int i = 1;
	        for (File image : images) {
	            BufferedImage frame = ImageIO.read(image);
	            writer.encodeVideo(0, frame, frameRate * i, TimeUnit.SECONDS);
	            i++;
	        }

	        writer.close();
	}
}