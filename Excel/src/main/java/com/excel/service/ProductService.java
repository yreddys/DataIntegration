package com.excel.service;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import com.excel.entity.Product;
import com.excel.repository.ProductRepository;

@Service
public class ProductService {

	@Autowired
	private ProductRepository productRepository;

	public void saveExcelData(MultipartFile file) {
		try {
			List<Product> products = parseExcelFile(file.getInputStream());
			productRepository.saveAll(products);
		} catch (IOException e) {
			e.printStackTrace();
			// Handle the exception properly
		}
	}

	private List<Product> parseExcelFile(InputStream is) {
	    List<Product> products = new ArrayList<>();
	    try {
	        Workbook workbook = new XSSFWorkbook(is);
	        Sheet sheet = workbook.getSheetAt(0);
	        for (Row row : sheet) {
	            // Skip the header row
	            if (row.getRowNum() == 0) {
	                continue;
	            }
	            Product product = new Product();
	            
	            // Handling cell types
	            Cell nameCell = row.getCell(0);
	            Cell descriptionCell = row.getCell(1);
	            Cell priceCell = row.getCell(2);
	            
	            if (nameCell != null) {
	                switch (nameCell.getCellType()) {
	                    case STRING:
	                        product.setName(nameCell.getStringCellValue());
	                        break;
	                    case NUMERIC:
	                        product.setName(String.valueOf(nameCell.getNumericCellValue()));
	                        break;
	                    default:
	                        product.setName(""); // Default value or handle appropriately
	                }
	            }

	            if (descriptionCell != null) {
	                switch (descriptionCell.getCellType()) {
	                    case STRING:
	                        product.setDescription(descriptionCell.getStringCellValue());
	                        break;
	                    case NUMERIC:
	                        product.setDescription(String.valueOf(descriptionCell.getNumericCellValue()));
	                        break;
	                    default:
	                        product.setDescription(""); // Default value or handle appropriately
	                }
	            }

	            if (priceCell != null) {
	                if (priceCell.getCellType() == CellType.NUMERIC) {
	                    product.setPrice(priceCell.getNumericCellValue());
	                } else if (priceCell.getCellType() == CellType.STRING) {
	                    try {
	                        product.setPrice(Double.parseDouble(priceCell.getStringCellValue()));
	                    } catch (NumberFormatException e) {
	                        // Handle the exception or set a default value
	                        product.setPrice(0.0);
	                    }
	                }
	            }
	            
	            products.add(product);
	        }
	        workbook.close();
	    } catch (IOException e) {
	        e.printStackTrace();
	        // Handle the exception properly
	    }
	    return products;
	}

}
