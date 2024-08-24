package com.excel.controller;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.excel.service.ProductService;

@RestController
@RequestMapping("/api/products")
public class ProductController {

	@Autowired
	private ProductService productService;

	@PostMapping("/upload")
	public String uploadExcelFile(@RequestParam("file") MultipartFile file) {
		if (file.isEmpty()) {
			return "Please select a file!";
		}
		productService.saveExcelData(file);
		return "File uploaded successfully!";
	}
}
