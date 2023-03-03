package com.report;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Array;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.NotOfficeXmlFileException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import views.html.ac.internal.main;

public class Data2 {

public static void main(String[] args) throws NotOfficeXmlFileException, IOException {
	writeDataRowwise();
	readExcel1();
	}

public static Object writeDataRowwise() {

	int developmentCritical=0,openCritical=0,fixRejectedCritical=0,codeReviewCritical=0,testingCritical=0,readyforQACritical=0;
	int developmentLow=0,openLow=0,fixRejectedLow=0,codeReviewLow=0,testingLow=0,readyforQALow=0;
	int developmentMajor=0,openMajor=0,fixRejectedMajor=0,codeReviewMajor=0,testingMajor=0,readyforQAMajor=0;
	int developmentModerate=0,openModerate=0,fixRejectedModerate=0,codeReviewModerate=0,testingModerate=0,readyforQAModerate=0;
	int grandTotalRowCritical=0,grandTotalRowLow=0,grandTotalRowMajor=0,grandTotalRowModerate=0;
	int gTCellCodereview=0,gTCelldevelopment=0,gTCellfixRejected=0,gTCellopen=0,gTCellreadyforQA=0,gTCelltesting=0,gTCellGrandTotal=0;
	String severity="";
	String status="";

	try {
		File obj1= new File((System.getProperty("user.dir") +"/Report.xlsx"));
		FileInputStream creds;
		creds = new FileInputStream(obj1);
		XSSFWorkbook workbook = new XSSFWorkbook(creds);
		XSSFSheet sheet = workbook.getSheetAt(1);
		int row = sheet.getLastRowNum();
		System.out.println("row count :"+ row);
		int col = sheet.getRow(0).getLastCellNum();
		System.out.println("col no : "+ col);
		
		int a2=report_CycleName.defects_rowsize;

		for (int i=1;i<=a2;i++) {
				
			status=sheet.getRow(i).getCell(2).getStringCellValue().trim();
			severity=sheet.getRow(i).getCell(3).getStringCellValue().trim(); 

			if (status.equalsIgnoreCase("Code Review") && severity.equals("Critical")){codeReviewCritical++;}
			else if (status.equalsIgnoreCase("Development") && severity.equals("Critical")){developmentCritical++;}
			else if (status.equalsIgnoreCase("Fix Rejected") && severity.equals("Critical")){fixRejectedCritical++;}
			else if(status.equalsIgnoreCase("Open") && severity.equals("Critical")){openCritical++;}
			else if (status.equalsIgnoreCase("Ready for QA") && severity.equals("Critical")){readyforQACritical++;}
			else if (status.equalsIgnoreCase("Testing") && severity.equals("Critical")){testingCritical++;}
		
			else if (status.equalsIgnoreCase("Code Review") && severity.equals("Low")){codeReviewLow++;}
			else if (status.equalsIgnoreCase("Development") && severity.equals("Low")){developmentLow++;}
			else if (status.equalsIgnoreCase("Fix Rejected") && severity.equals("Low")){fixRejectedLow++;}
			else if(status.equalsIgnoreCase("Open") && severity.equals("Low")){openLow++;}
			else if (status.equalsIgnoreCase("Ready for QA") && severity.equals("Low")){readyforQALow++;}
			else if (status.equalsIgnoreCase("Testing") && severity.equals("Low")){testingLow++;}
			
			else if (status.equalsIgnoreCase("Code Review") && severity.equals("Major")){codeReviewMajor++;}
			else if (status.equalsIgnoreCase("Development") && severity.equals("Major")){developmentMajor++;}
			else if (status.equalsIgnoreCase("Fix Rejected") && severity.equals("Major")){fixRejectedMajor++;}
			else if(status.equalsIgnoreCase("Open") && severity.equals("Major")){openMajor++;}
			else if (status.equalsIgnoreCase("Ready for QA") && severity.equals("Major")){readyforQAMajor++;}
			else if (status.equalsIgnoreCase("Testing") && severity.equals("Major")){testingMajor++;}
			
			else if (status.equalsIgnoreCase("Code Review") && severity.equals("Moderate")){codeReviewModerate++;}
			else if (status.equalsIgnoreCase("Development") && severity.equals("Moderate")){developmentModerate++;}
			else if (status.equalsIgnoreCase("Fix Rejected") && severity.equals("Moderate")){fixRejectedModerate++;}
			else if(status.equalsIgnoreCase("Open") && severity.equals("Moderate")){openModerate++;}
			else if (status.equalsIgnoreCase("Ready for QA") && severity.equals("Moderate")){readyforQAModerate++;}
			else if (status.equalsIgnoreCase("Testing") && severity.equals("Moderate")){testingModerate++;}
		
			else if (status.equalsIgnoreCase("Code Review") && severity.equals("Null")){codeReviewModerate++;}
			else if (status.equalsIgnoreCase("Development") && severity.equals("Null")){developmentModerate++;}
			else if (status.equalsIgnoreCase("Fix Rejected") && severity.equals("Null")){fixRejectedModerate++;}
			else if (status.equalsIgnoreCase("Open") && severity.equals("Null")){openModerate++;}
			else if (status.equalsIgnoreCase("Ready for QA") && severity.equals("Null")){readyforQAModerate++;}
			else if (status.equalsIgnoreCase("Testing") && severity.equals("Null")){testingModerate++;}
		}	
	}catch (FileNotFoundException e) {
		e.printStackTrace();
	} 
	catch (NullPointerException e) {
		e.printStackTrace();
		return null;
		}
	catch (IOException e) {

		e.printStackTrace();
	}

	grandTotalRowCritical=codeReviewCritical+developmentCritical+fixRejectedCritical+openCritical+readyforQACritical+testingCritical;
	
	ArrayList<Integer>list1=new ArrayList<>();
	list1.add(codeReviewCritical);
	list1.add(developmentCritical);
	list1.add(fixRejectedCritical);
	list1.add(openCritical);
	list1.add(readyforQACritical);
	list1.add(testingCritical);
	list1.add(grandTotalRowCritical);
	
	for(int i=0;i<=list1.size()-1;i++) {

		String s1=String.valueOf(list1.get(i));
		WriteInExcelRow1(0, i, s1);
	}
	
	grandTotalRowLow=codeReviewLow+developmentLow+fixRejectedLow+openLow+readyforQALow+testingLow;
	ArrayList<Integer>list2=new ArrayList<>();
	list2.add(codeReviewLow);
	list2.add(developmentLow);
	list2.add(fixRejectedLow);
	list2.add(openLow);
	list2.add(readyforQALow);
	list2.add(testingLow);
	list2.add(grandTotalRowLow);
	
	for(int i=0;i<=list2.size()-1;i++) {
		String s2=String.valueOf(list2.get(i));
		WriteInExcelRow1(1, i, s2);		
	}
	
grandTotalRowMajor=codeReviewMajor+developmentMajor+fixRejectedMajor+openMajor+readyforQAMajor+testingMajor;
	
	ArrayList<Integer>list3=new ArrayList<>();
	list3.add(codeReviewMajor);
	list3.add(developmentMajor);
	list3.add(fixRejectedMajor);
	list3.add(openMajor);
	list3.add(readyforQAMajor);
	list3.add(testingMajor);
	list3.add(grandTotalRowMajor);
	
	for(int i=0;i<=list3.size()-1;i++) {
		String s3=String.valueOf(list3.get(i));
		WriteInExcelRow1(2, i, s3);
	}

grandTotalRowModerate=codeReviewModerate+developmentModerate+fixRejectedModerate+openModerate+readyforQAModerate+testingModerate;


	ArrayList<Integer>list4=new ArrayList<>();
	list4.add(codeReviewModerate);
	list4.add(developmentModerate);
	list4.add(fixRejectedModerate);
	list4.add(openModerate);
	list4.add(readyforQAModerate);
	list4.add(testingModerate);
	list4.add(grandTotalRowModerate);
	
	for(int i=0;i<=list4.size()-1;i++) {
		String s4=String.valueOf(list4.get(i));
		WriteInExcelRow1(3, i, s4);
	}
	
	gTCellCodereview=codeReviewModerate+codeReviewMajor+codeReviewLow+codeReviewCritical;
	gTCelldevelopment=developmentModerate+developmentMajor+developmentLow+developmentCritical;
	gTCellfixRejected=fixRejectedModerate+fixRejectedMajor+fixRejectedLow+fixRejectedCritical;
	gTCellopen=openModerate+openMajor+openLow+openCritical;
	gTCellreadyforQA=readyforQAModerate+readyforQAMajor+readyforQALow+readyforQACritical;
	gTCelltesting=testingModerate+testingMajor+testingLow+testingCritical;
	gTCellGrandTotal=grandTotalRowModerate+grandTotalRowMajor+grandTotalRowLow+grandTotalRowCritical;
	
	ArrayList<Integer>total=new ArrayList<>();
	total.add(gTCellCodereview);
	total.add(gTCelldevelopment);
	total.add(gTCellfixRejected);
	total.add(gTCellopen);
	total.add(gTCellreadyforQA);
	total.add(gTCelltesting);
	total.add(gTCellGrandTotal);
	
	for(int i=0;i<=total.size()-1;i++) {
		String s5=String.valueOf(total.get(i));
		WriteInExcelRow1(4, i, s5);
	}
	
	return null;
	
}

public static void WriteInExcelRow1(int RowNumber ,int ColNumber, String s1 ) {

			try   
			{  	   int rowNo = RowNumber+1;
				   int colNumber=ColNumber+1;
				String userDirectory = System.getProperty("user.dir");
				String filename = userDirectory + "/" +"Report.xlsx" ;  
				
				FileInputStream fis = new FileInputStream(filename); 
				XSSFWorkbook workbook = new XSSFWorkbook(fis);
				XSSFSheet sheet = workbook.getSheetAt(2);
				Cell cell = sheet.getRow(rowNo).getCell(colNumber);
				cell.setCellValue(s1);
				fis.close();
				FileOutputStream outFile = new FileOutputStream((filename));
				workbook.write(outFile);
				outFile.close();
			}   
			catch (Exception e)   
			{  
				e.printStackTrace();  
			}  
		}

public static void clearExcelData3()
{
	try
	{
		String userDirectory = System.getProperty("user.dir");
		FileInputStream fis = new FileInputStream(userDirectory + "/"+ "Report.xlsx");
	
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheetAt(2);
//		int a2=report_CycleName.defects_rowsize;
		
		
		for(int i =1;i<=5; i++)
		{
			    Cell cell0 = sheet.getRow(i).getCell(1);
			    cell0.setCellValue(" ");
			
				Cell cell1 = sheet.getRow(i).getCell(2);
				cell1.setCellValue(" ");
				
				Cell cell2 = sheet.getRow(i).getCell(3);
				cell2.setCellValue(" ");
				
				Cell cell3 = sheet.getRow(i).getCell(4);
				cell3.setCellValue(" ");
				
				Cell cell4 = sheet.getRow(i).getCell(5);
				cell4.setCellValue(" ");
				
				Cell cell5 = sheet.getRow(i).getCell(6);
				cell5.setCellValue(" ");
				
				Cell cell6 = sheet.getRow(i).getCell(7);
				cell6.setCellValue(" ");
		}
		  
		FileOutputStream output = new FileOutputStream(userDirectory + "/" +"Report.xlsx");
		workbook.write(output);
		output.close();
	
	workbook.close();
	} 
	catch (Exception e) 
	{
		e.printStackTrace();
	}

}
public static String readExcel1(){
	try {
	String userDirectory = System.getProperty("user.dir");
	FileInputStream fis = new FileInputStream(userDirectory + "/"+ "Report.xlsx");

	XSSFWorkbook workbook = new XSSFWorkbook(fis);
	XSSFSheet sheet = workbook.getSheetAt(2);
//	Row row = sheet.getRow(rowNo);
	int row = sheet.getLastRowNum();
	System.out.println("row count"+row);
	int col = sheet.getRow(0).getLastCellNum();
	System.out.println("cell count"+col);
	
	
	float codeReviewPercent=0, developmentPercent=0,fixRejectedPercent=0,openPercent=0,readyForQAPercent=0,TestingPercent=0;
	int grandTotal=0,codeReview=0,development=0,fixRejected=0,open=0, readyForQA=0, Testing=0;

       codeReview = Integer.parseInt(sheet.getRow(5).getCell(1).getStringCellValue());
       System.out.println("codeReview: "+ codeReview);
       development = Integer.parseInt(sheet.getRow(5).getCell(2).getStringCellValue());
       fixRejected = Integer.parseInt(sheet.getRow(5).getCell(3).getStringCellValue());
       open = Integer.parseInt(sheet.getRow(5).getCell(4).getStringCellValue());
       System.out.println("open :"+ open);
       readyForQA = Integer.parseInt(sheet.getRow(5).getCell(5).getStringCellValue());
       Testing = Integer.parseInt(sheet.getRow(5).getCell(6).getStringCellValue());
       grandTotal = Integer.parseInt(sheet.getRow(5).getCell(7).getStringCellValue());
       System.out.println("grandTotal :"+ grandTotal);
       
       
       if(codeReview>0 && grandTotal>0) {
   		codeReviewPercent=(float)(((int)(((codeReview*100)/grandTotal) *100.0))/100.0);
   		}
       else {
    	   codeReviewPercent=0;
       }
       
       if(development>0 && grandTotal>0) {
    	   developmentPercent=(float)(((int)(((development*100)/grandTotal) *100.0))/100.0);
      		}
          else {
        	  developmentPercent=0;
          }
       
       if(fixRejectedPercent>0 && grandTotal>0) {
    	   fixRejectedPercent=(float)(((int)(((fixRejected*100)/grandTotal) *100.0))/100.0);
      		}
          else {
        	  fixRejectedPercent=0;
          }
       if(openPercent>0 && grandTotal>0) {
    	   openPercent=(float)(((int)(((open*100)/grandTotal) *100.0))/100.0);
      		}
          else {
        	  openPercent=0;
          }
       if(readyForQAPercent>0 && grandTotal>0) {
    	   readyForQAPercent=(float)(((int)(((readyForQA*100)/grandTotal) *100.0))/100.0);
      		}
          else {
        	  readyForQAPercent=0;
          }
       if(TestingPercent>0 && grandTotal>0) {
           TestingPercent=(float)(((int)(((Testing*100)/grandTotal) *100.0))/100.0);
      		}
          else {
        	  TestingPercent=0;
          }      
      
       System.out.println("TestingPercent :"+ TestingPercent);
	

	workbook.close();
	} catch (Exception e) {
		e.printStackTrace();
	}
	return null;	
}

}