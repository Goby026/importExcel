package com.gobydev.importexcel.controller;

import java.io.IOException;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import com.gobydev.importexcel.entity.Product;

@Controller
public class ImportExcelController {
	
	@RequestMapping(value = "/import-excel", method = RequestMethod.POST)
	public ResponseEntity<List<Product>> importExcelFile( @RequestParam("file") MultipartFile files) throws IOException, ParseException{
		
		HttpStatus status = HttpStatus.OK;
		List<Product> productsList = new ArrayList<>();
		
		XSSFWorkbook workbook = new XSSFWorkbook( files.getInputStream() );
		XSSFSheet worksheet = workbook.getSheetAt(0);
		
		for (int index = 0; index < worksheet.getPhysicalNumberOfRows(); index++) {
			
			if (index > 0) {
				Product product = new Product();
				
				XSSFRow row = worksheet.getRow(index);
				Integer id = (int) row.getCell(0).getNumericCellValue();
				
				product.setId(id);
				product.setName(row.getCell(1).getStringCellValue());
				product.setPrice(row.getCell(2).getNumericCellValue());
				product.setQuantity((int)row.getCell(3).getNumericCellValue());				
				product.setCreation_date(row.getCell(4).getDateCellValue());
				product.setStatus(new Boolean( row.getCell(5).getStringCellValue() ) );
				
				productsList.add(product);
				
			}
			
		}
		
		return new ResponseEntity<>(productsList, status);
		
	}

}
