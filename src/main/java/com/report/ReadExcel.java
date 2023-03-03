package com.report;

import java.io.FileInputStream;
import java.util.ArrayList;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel 
{
	public static int rowNo = 1;
	public static void clearExcelData()
	{
		try
		{
			String userDirectory = System.getProperty("user.dir");
			FileInputStream fis = new FileInputStream(userDirectory + "/"+ "Report.xlsx");
		
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			XSSFSheet sheet = workbook.getSheetAt(0);
			
			for(int i = sheet.getLastRowNum(); i >= 1; i--)
			{
			  Row row = sheet.getRow(i);
			   sheet.removeRow(row);
			  
			}
			
		
		workbook.close();
		} 
		catch (Exception e) 
		{
			e.printStackTrace();
		}

	}
	@SuppressWarnings("resource")
	public static String readExcel(String data){
		try {
		String userDirectory = System.getProperty("user.dir");
		FileInputStream fis = new FileInputStream(userDirectory + "/"+ "Report.xlsx");
	
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheetAt(0);
		Row row = sheet.getRow(rowNo);
		{ 
			//row= sheet.getRow(1);
			if(data.equals(null))
			{
				System.out.println("projectId, cycleId and folderId is null");
			}
			else if(data.equals("projectId"))
				{
				Cell Idcellpid = row.getCell(1);
				int pid =  (int) Idcellpid.getNumericCellValue();
		
			      String pcode = Integer.toString(pid);
			    return pcode;	
				}
				
				
				else if(data.equals("cycleId"))
				{
					Cell Idcellcid = row.getCell(3);
					String cid = Idcellcid.getStringCellValue();
					return cid;
				}
			
				
				else if(data.equals("folderId"))
				{
					Cell Idcellfid = row.getCell(5);
					String fid = Idcellfid.getStringCellValue();
					return fid;
				}
			
				else if(data.equals("cycleName"))
				{
					Cell Idcellfid = row.getCell(2);
					String fid = Idcellfid.getStringCellValue();
					return fid;
				}
				else if(data.equals("folderName"))
				{
					Cell Idcellfid = row.getCell(4);
					String fid = Idcellfid.getStringCellValue();
					return fid;
				}
		}
		
		
		
		
		
		workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		return null;
		
	}


}
