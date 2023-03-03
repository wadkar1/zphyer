package com.report;

import java.io.BufferedWriter;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

import org.apache.commons.text.CaseUtils;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class ReadData2 {
	public static FileInputStream fis;
	@SuppressWarnings("finally")
	public static void main (String[] args) throws InterruptedException, IOException
	{
		ExcelToHtml(null, null);
		
	}
	public static void ExcelToHtml(String info,String projectName) throws IOException
	{		
	String date =(LocalDateTime.now().format(DateTimeFormatter.ofPattern("MMM dd, yyyy")));
	
	File obj1= new File((System.getProperty("user.dir") +"/Report.xlsx"));
	FileInputStream creds = new FileInputStream(obj1);
	XSSFWorkbook workbook = new XSSFWorkbook(creds);
	XSSFSheet sheet = workbook.getSheetAt(0);
	int rowCount= sheet.getPhysicalNumberOfRows();
	int ActualCount=0;
	int a=report_CycleName.last_rowCount;
//	System.out.println(a);
	for (int i=1; i<=a;i++)
	{
		try {
		String total1=sheet.getRow(i).getCell(11).getStringCellValue().trim();
		if(total1.equals("") || total1.equals(" "))
		continue;
		else
			ActualCount++;
		}
		finally{
			continue;
		}
	}
	//reading header values
	String header2 = CaseUtils.toCamelCase((sheet.getRow(0).getCell(2).getStringCellValue()),true,' ');
	String header4 = CaseUtils.toCamelCase((sheet.getRow(0).getCell(4).getStringCellValue()),true,' ');
	String header6 = CaseUtils.toCamelCase((sheet.getRow(0).getCell(6).getStringCellValue()),true,' ');
	String header7 = CaseUtils.toCamelCase((sheet.getRow(0).getCell(7).getStringCellValue()),true,' ');
	String header8 = CaseUtils.toCamelCase((sheet.getRow(0).getCell(8).getStringCellValue()),true,' ');
	String header9 = CaseUtils.toCamelCase((sheet.getRow(0).getCell(9).getStringCellValue()),true,' ');
	String header10 = CaseUtils.toCamelCase((sheet.getRow(0).getCell(10).getStringCellValue()),true,' ');
	String header11 = CaseUtils.toCamelCase((sheet.getRow(0).getCell(11).getStringCellValue()),true,' ');
	String header12 = CaseUtils.toCamelCase((sheet.getRow(0).getCell(12).getStringCellValue()),true,' ');
	
	//displaying header values
	String html ="<html><head><title>Test Result</title>"
			
			+"<style>"
			+"body {background: #f6f6f6;}.map {overflow: auto;}"
			+"ul li {text-align: left;margin: 5px 0;color: #77787a;font-size: 14px;}"
			+".table {width: 100%;margin-bottom: 1rem;background-color: transparent;margin: 15px 0;border-collapse: collapse;font-size: 14px;}"
			+".table thead th {vertical-align: bottom;}"
			+".table td, .table th {padding: 0.75rem;vertical-align: top;text-align: center;}"
			+".table th {border-top: 1px solid #dee2e6;}"
			+ ".table-striped tbody tr:{border-bottom:1px solid #F2F2F2;}"
			+".primary {color: #007bff;}.success {color: #28a745;}.danger {color: #dc3545;}"
			+"* {color: #222;font-family: Arial, Helvetica, sans-serif;}"
			+".info {background: #F2F2F2!important;}"
			+"</style>"
			
			+"</head>"
			
			+"<body style=\"margin: 0\">"
			+"<table cellpadding=\"0\" cellspacing=\"0\" width=\"100%\" style=\"margin: 0 auto;font-family: Arial, Helvetica, sans-serif;color: #1f1f1f;\">"
			+"<tbody><tr><td>"
			+"<table cellpadding=\"0\" cellspacing=\"0\" width=\"100%\"><tbody><tr><td>"
			+"<table cellpadding=\"0\" cellspacing=\"0\" style=\"margin: 0px\" width=\"100%\">"
			+"<tbody><tr><td style=\"padding: 40px 5px 10px 5px;\">"
			+"<table width=\"100%\" cellpadding=\"cellspacing=\"0\">"
			+"<tbody><tr><td width=\"10%\" align=\"left\" style=\"padding: 0 16px 0 24px\"> <a href=\"#\" style=\"display: block\"> <img src=\"https://www.sourcefuse.com/wp-content/uploads/2021/09/SourceFuse-logo.png\" alt=\"sourcefuse\"width=\"136px\"/> </a> </td>"
			+"<td width=\"70%\" align=\"center\" style=\"padding: 0 16px 0 24px\">"
			+"<h3 style=\"margin: 0;\"> Daily Excecution Status </h3>"
			+"</td><td width=\"18%\" align=\"right\" style=\"padding: 0 30px 0 20px;line-height:"
			+"18px;font-size: 16px;\">"+date+"</td></tr></tbody></table></td></tr></tbody>"
			+"</table>"
			
			+"<table width=\"96%\" cellpadding=\"0\" cellspacing=\"0\" style=\"margin: 24px auto 25px; border: 1px solid #0ED2AF;\">"
			+"<tbody>"
			+"<tr><td style=\"padding: 35px;background: #fff;\">"
			+"<table cellpadding=\"0\" cellspacing=\"0\" width=\"100%\">"  
			+"<tbody><tr><td>"
			+"<p style=\"font-size: 15px;color: #77787a;margin: 3px 0;\"> Hello All, </p>"
			+"</td></tr>"
			+"<tr><td>"
			+ "<p style=\"font-size: 15px;color: #77787a;margin: 3px 0;\"> Please find the attached report below: </p>"
			+"</td></tr>"
			+"<tr><td>"
			+"<p style=\"font-size: 15px;color: #4a4b4c;margin: 15px 0 5px 0;\"> Project Name: "+projectName+" </p>"
			+"</td></tr>"
			+"<tr><td>"
			+"<p style=\"font-size: 15px;color: #4a4b4c;margin: 5px 0 5px 0;\"> Release - V: 0.1 </p>"
			+"<p style=\"font-size: 14px;color: #77787a;margin: 5px 0 5px 0px;\"> We have completed the Test Case creation activity till Sprint 62 including Peer and Product Review. </p>" 
			+"</td></tr>"
			+"<tr><td>" 
			+"<p style=\"font-size: 15px;color: #4a4b4c;margin: 10px 0 0px 0;font-weight: 600;\"> Test Execution Summary: </p>" 
			+"</td></tr></tbody>" 
			+"</table>"
			
			+"<table class=\"table table-striped custom-table\">"
			+"<thead><tr style=\"border-bottom:1px solid #d5cfcf;\">"
			+"<th>"+header2+"</th>"
			+"<th>"+header4+"</th>"
			+"<th>"+header6+"</th>"
			+"<th>"+header7+"</th>"
			+"<th>"+header8+"</th>"
			+"<th>"+header9+"</th>"
			+"<th>"+header10+"</th>"
			+"<th>"+header11+"</th>"
			+"<th>"+header12+"</th>"
			+"</tr></thead>";
			
	
	String defects="",wip="",blocked="",unexecuted="", fail="", pass="", total="";
	int total1=0, total2=0, pass1=0, pass2=0, fail2=0, fail1=0, unexecuted1=0, unexecuted2 = 0, wip1=0, wip2=0, blocked1=0, blocked2=0, defects1=0, defects2= 0;
	float passPercent=0, failPercent=0, unexePercent=0, wipPercent=0, blockedPercent=0;
	
	for (int i=1; i<=ActualCount;i++)
	{
	total=sheet.getRow(i).getCell(11).getStringCellValue().trim();
	if(total.equals(""))
	{
		total1 = 0; 
	}
	else
	{
		total1=Integer.parseInt(sheet.getRow(i).getCell(11).getStringCellValue().trim());
	}
	total2  = total2 + total1;
	pass=sheet.getRow(i).getCell(6).getStringCellValue().trim();
	if(pass.equals(""))
	{
		pass1 = 0; 
	}
	else
	{
		pass1=Integer.parseInt(sheet.getRow(i).getCell(6).getStringCellValue().trim());
	}
	pass2=pass2 + pass1;
	passPercent=(float)(((int)(((pass2*100)/total2) *100.0))/100.0);
	fail=sheet.getRow(i).getCell(7).getStringCellValue().trim();
	if(fail.equals(""))
	{
		fail1 = 0; 
	}
	else
	{
		fail1=Integer.parseInt(sheet.getRow(i).getCell(7).getStringCellValue().trim());
	}	
	fail2=fail2 + fail1;
	failPercent=(float)(((int)(((fail2*100)/total2) *100.0))/100.0);
	unexecuted=sheet.getRow(i).getCell(8).getStringCellValue().trim();
	if(unexecuted.equals(""))
	{
		unexecuted1 = 0; 
	}
	else
	{
		unexecuted1=Integer.parseInt(sheet.getRow(i).getCell(8).getStringCellValue().trim());
	}
	unexecuted2=unexecuted2 + unexecuted1;
	unexePercent=(float)(((int)(((unexecuted2*100)/total2) *100.0))/100.0);
	wip=sheet.getRow(i).getCell(9).getStringCellValue().trim();
	if(wip.equals(""))
	{
		wip1 = 0; 
	}
	else
	{
		wip1=Integer.parseInt(sheet.getRow(i).getCell(9).getStringCellValue().trim());
	}
	wip2=wip2+wip1;
	wipPercent=(float)(((int)(((wip2*100)/total2) *100.0))/100.0);
	blocked = sheet.getRow(i).getCell(10).getStringCellValue().trim();
	if(blocked.equals(""))
	{
		blocked1 = 0; 
	}
	else
	{
		blocked1=Integer.parseInt(sheet.getRow(i).getCell(10).getStringCellValue().trim());
	}
	blocked2=blocked2 + blocked1;
	blockedPercent=(float)(((int)(((blocked2*100)/total2) *100.0))/100.0);
	defects = sheet.getRow(i).getCell(12).getStringCellValue().trim();
	if(defects.equals(""))
	{
		defects1 = 0; 
	}
	else
	{
		defects1=Integer.parseInt(sheet.getRow(i).getCell(12).getStringCellValue());
	}
	defects2= defects2 + defects1;
	}
	
	//to read data till the rowCount
	for(int i =1; i <=ActualCount; i++ )
	{
		XSSFRow row=sheet.getRow(i);
		if(row!=null)
		{
		String data2 = sheet.getRow(i).getCell(2).getStringCellValue();
		String data4 = sheet.getRow(i).getCell(4).getStringCellValue();
		String data6 = sheet.getRow(i).getCell(6).getStringCellValue();
		String data7 = sheet.getRow(i).getCell(7).getStringCellValue();
		String data8 = sheet.getRow(i).getCell(8).getStringCellValue();
		String data9 = sheet.getRow(i).getCell(9).getStringCellValue();
		String data10 = sheet.getRow(i).getCell(10).getStringCellValue();
		String data11 = sheet.getRow(i).getCell(11).getStringCellValue();
		String data12 = sheet.getRow(i).getCell(12).getStringCellValue();
		

		html = html+"<tbody><tr style=\"border-bottom:1px solid #d5cfcf;\">"
				  +"<td>"+data2+"</td>"
				  +"<td class=\"primary\">"+data4+"</td>"
				  +"<td class=\"success\">"+data6+"</td>"
				  +"<td class=\"danger\">"+data7+"</td>" 
				  +"<td>"+data8+"</td>"
				  +"<td>"+data9+"</td>" 
				  +"<td class=\"danger\">"+data10+"</td>"
				  +"<td>"+data11+"</td>"
				  +"<td class=\"danger\">"+data12+"</td>" 
				  +"</tr>";
	
		}
	}
	//total count row
	
	html=html+"<tr>"
			+"<td colspan=\"2\"><b>Total Count</b></td>"
			
			+"<td><b>"+pass2+"</b></td>"
			+"<td><b>"+fail2+"</b></td>"
			+"<td><b>"+unexecuted2+"</b></td>"
			+"<td><b>"+wip2+"</b></td>"
			+"<td><b>"+blocked2+"</b></td>"
			+"<td><b>"+total2+"</b></td>"
			+"<td><b>"+defects2+"</b></td>"                                       
			+"</tr>"
			+"</tbody></table>"
			+ "</td></tr>"
			+ "</td></tr></tbody></table>";
	
	if(defects2!=0) 
	{
		File obj3= new File((System.getProperty("user.dir") +"/Report.xlsx"));
		FileInputStream creds3 = new FileInputStream(obj3);
		XSSFWorkbook workbook3 = new XSSFWorkbook(creds3);
		XSSFSheet sheet3 = workbook3.getSheetAt(2);
		int ActualCount2=0;
		int a2=report_CycleName.defects_rowsize;
		System.out.println("Defects Count: "+a2);
		int rowcount = sheet3.getLastRowNum();
		
		
		for (int i=1; i<=rowcount;i++)
		{
			try {
			String grandtotal2=sheet3.getRow(i).getCell(7).getStringCellValue().trim();
			if(grandtotal2.equals("") || grandtotal2.equals(" "))
			continue;
			else
				ActualCount2++;
			}
			finally{
				continue;
			}
		}
		//reading the defect header value 
		String dheader1 = CaseUtils.toCamelCase((sheet3.getRow(0).getCell(0).getStringCellValue()),true,' ');
		String dheader2 = CaseUtils.toCamelCase((sheet3.getRow(0).getCell(1).getStringCellValue()),true,' ');
		String dheader3 = CaseUtils.toCamelCase((sheet3.getRow(0).getCell(2).getStringCellValue()),true,' ');
		String dheader4 = CaseUtils.toCamelCase((sheet3.getRow(0).getCell(3).getStringCellValue()),true,' ');
		String dheader5 = CaseUtils.toCamelCase((sheet3.getRow(0).getCell(4).getStringCellValue()),true,' ');
		String dheader6 = CaseUtils.toCamelCase((sheet3.getRow(0).getCell(5).getStringCellValue()),true,' ');
		String dheader7 = CaseUtils.toCamelCase((sheet3.getRow(0).getCell(6).getStringCellValue()),true,' ');
		String dheader8 = CaseUtils.toCamelCase((sheet3.getRow(0).getCell(7).getStringCellValue()),true,' ');
		
		html=html+"<div style=\"padding:35px; background: #fff;margin: 0 24px;border: 1px solid #0ED2AF;\">"
				+"<p style=\"font-size: 15px;color: #4a4b4c;margin: 10px 0 0px 0;font-weight: 600;\"> Defect Severity VS Status : </p>"
				+"<table class=\"table table-striped custom-table\">"
				+"<thead>"
				+ "<tr style=\"border-bottom:1px solid #d5cfcf;\">"
				+"<th style=\"text-align:left;\">"+dheader1+"</th>"
				+"<th>"+dheader2+"</th>"
				+"<th>"+dheader3+"</th>"
				+"<th>"+dheader4+"</th>"
				+ "<th>"+dheader5+"</th>"
				+ "<th>"+dheader6+"</th>"
				+ "<th>"+dheader7+"</th>"
				+ "<th>"+dheader8+"</th>"
						+ "</tr></thead>";
				
//			System.out.println("ActualCount2 :"+ActualCount2);
			for (int i=1; i<=5;i++)
			{
					String defdata1 = sheet3.getRow(i).getCell(0).getStringCellValue();
					String defdata2 = sheet3.getRow(i).getCell(1).getStringCellValue();
					String defdata3 = sheet3.getRow(i).getCell(2).getStringCellValue();
					String defdata4 = sheet3.getRow(i).getCell(3).getStringCellValue();
					String defdata5 = sheet3.getRow(i).getCell(4).getStringCellValue();
					String defdata6 = sheet3.getRow(i).getCell(5).getStringCellValue();
					String defdata7 = sheet3.getRow(i).getCell(6).getStringCellValue();
					String defdata8 = sheet3.getRow(i).getCell(7).getStringCellValue();
					
					html = html+"<tbody><tr style=\"border-bottom:1px solid #d5cfcf;\">"
							+"<td style=\"text-align:left;\"><b>"+defdata1+"</b></td>"
							+"<td >"+defdata2+"</td>"
							+"<td >"+defdata3+"</td>"
							+ "<td >"+defdata4+"</td>"
							+ "<td >"+defdata5+"</td>"
							+ "<td >"+defdata6+"</td>"
							+ "<td>"+defdata7+"</td>"
							+ "<td>"+defdata8+"</td>"
							+"</tr>";
		}
			
	}
	  html=html+"</tbody></table></table>"
			+ "<div style=\"margin: 24px;border: 1px solid #0ED2AF;padding:35px;\">"
	+ "<p style=\"font-size: 15px;color: #4a4b4c;margin: 0px 0 5px 0;text-align: left; font-weight: 600;\"> Test Case Info:</p>"
	+"<p style=\"font-size: 14px;color: #77787a;;margin: 5px 0 5px 0px;text-align: left;\">"
    + "<ul><li>** Feature bugs are the bugs which are related to the Stories and Improvements with Label</li>"
    +"<li>"+info+"</li>"
	+ "</ul></p>"
	+"</div></div></td></tr></tbody></table></td></tr></tbody></table>";
			
    html= html+ "</body></html>";
	File fw = new File ("./result.html");
	BufferedWriter bw= new BufferedWriter(new FileWriter(fw));
	bw.write(html);
	bw.close();	
	workbook.close();
	}
}