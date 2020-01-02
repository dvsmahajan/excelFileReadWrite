package com.xworkz.readexcel.util;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.xworkz.readexcel.dto.ProductDTO;

public class ReadExcelFileUTIL {

		String path;
	    public FileInputStream fis = null;
	    private XSSFWorkbook workbook = null;
	    private XSSFSheet sheet = null;
	    private XSSFRow row   =null;
	    private XSSFCell cell = null;

	    public ReadExcelFileUTIL() throws IOException
	    {
	        fis = new FileInputStream("C:\\Users\\dvsma\\OneDrive\\Desktop\\product.xlsx"); 
	        workbook = new XSSFWorkbook(fis);
	        sheet = workbook.getSheetAt(0);
	    }
	    
	    public List<ProductDTO> readExcelFile()
	    {
	    	List<ProductDTO> list = new ArrayList<ProductDTO>();
	        int index = workbook.getSheetIndex("Sheet1");
	        sheet = workbook.getSheetAt(index);
	        int rownumber=sheet.getLastRowNum()+1;  

	        for (int i=0; i<rownumber; i++ )
	        {
	
	            row = sheet.getRow(i);
//	            int colnumber = row.getLastCellNum();
	
	        	ProductDTO dto = new ProductDTO(((int)row.getCell(0).getNumericCellValue()),row.getCell(1).getStringCellValue(),row.getCell(2).getStringCellValue(),((double)row.getCell(3).getNumericCellValue()) );
	
	        	list.add(dto);
//	            System.out.println("Id is "+row.getCell(0).getNumericCellValue()+
//	            		" Name is "+row.getCell(1).getStringCellValue()+
//	            		" Type is "+row.getCell(2).getStringCellValue()+
//	            		" Price is "+row.getCell(3).getNumericCellValue());
	        }
	        
	        return list;
	    } 

	    public static void main(String[] args) {
			try {
				
				ReadExcelFileUTIL excelFileUTIL = new  ReadExcelFileUTIL();
				excelFileUTIL.readExcelFile();
			} catch (IOException e) {
				System.out.println("Exception is comming.............");
				e.printStackTrace();
			}
			
		}
	    
}
