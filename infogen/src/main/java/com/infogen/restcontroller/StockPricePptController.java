package com.infogen.restcontroller;

import java.io.FileNotFoundException;
import java.io.IOException;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import com.infogen.service.StockPriceService;

@RestController
public class StockPricePptController {
	
	@Autowired
	StockPriceService stockPriceService;
	
	@PostMapping("/generateppt")
	public String generatePptForStockPrice(@RequestParam String spSymbol, @RequestParam String spInstrument) throws IOException {
		return stockPriceService.generatePptForStockPrice(spSymbol, spInstrument);
	}
	
	@GetMapping("/img")
	public String convertPptToImages() throws FileNotFoundException, IOException {
		stockPriceService.convertPptToImages();
		return "success";
	}
	
	@GetMapping("/vdo")
	public String convertImageToVideo() throws IOException {
		stockPriceService.convertImageToVideo();
		return "success";
	}
}
