package com.xworkz.readexcel.util;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.xworkz.readexcel.dto.ProductDTO;

public class WriteInExcelFile {

	public static void main(String[] args) {
		try {
			ReadExcelFileUTIL readExcelFileUTIL = new ReadExcelFileUTIL();
			List<ProductDTO> list =readExcelFileUTIL.readExcelFile();
			System.out.println("List is "+list);
			WriteInExcelFile writeInExcelFile = new WriteInExcelFile();
			writeInExcelFile.createExcelFile(list);
			System.out.println("done..........");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	private void createExcelFile(List<ProductDTO> list) {
		 //Blank workbook
        XSSFWorkbook workbook = new XSSFWorkbook(); 
         
        //Create a blank sheet
        XSSFSheet sheet = workbook.createSheet("Product Data");
        int rownum = 0;
		for(ProductDTO dto : list) {
			 Row row = sheet.createRow(rownum++);
//			 int cellnum = 0;
			 Cell cell = row.createCell(0);
			 cell.setCellValue(dto.getId());
			 cell = row.createCell(1);
			 cell.setCellValue(dto.getName());
			 cell = row.createCell(2);
			 cell.setCellValue(dto.getType());
			 cell = row.createCell(3);
			 cell.setCellValue(dto.getCost());
			 
		}
		
		 FileOutputStream out;
		try {
			out = new FileOutputStream(new File("C:\\Users\\dvsma\\OneDrive\\Desktop\\createProduct.xlsx"));
	         workbook.write(out);
	         out.close();
	         System.out.println("createProduct.xlsx written successfully on disk.");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
