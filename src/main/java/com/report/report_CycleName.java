package com.report;

import java.io.IOException;

import java.net.URISyntaxException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.*;
import org.json.simple.*;
import groovyjarjarasm.asm.tree.TryCatchBlockNode;
import io.restassured.RestAssured;
import io.restassured.path.json.JsonPath;
import io.restassured.response.Response;
import io.restassured.response.ResponseBodyExtractionOptions;
import io.restassured.specification.RequestSpecification;
import com.report.WriteExcel;

public class report_CycleName {

	public static JsonPath jp;
	public static String JWTID;
	public static String allCycleIdsUrl;
	public static String allFolderIdsUrl;
	public static String accessKey;
	
	public static String cycleId; 
	public static String cycleName ;
	public static String projectId;
	public static String folderId;
	public static String folderName;
	
	public static int defects_rowsize;

	public static String versionId; 
	public static String zephyrBaseUrl = "https://prod-api.zephyr4jiracloud.com/connect/";
	public static int folderCount = 0;
	public static int last_rowCount = 0;
	public static String fid_SummaryReportUrl; 
	public static String defectStatus = "Done";
	public static String defectStatus1 = "Scrap";
	public static   List<String> ListOfDefectNames = new ArrayList<>(); 
	public static   List<String> ListOfDefectNames1 = new ArrayList<>();
	public static  List<String> ListOfDefectStatus = new ArrayList<>();
	public static  List<String> ListOfDefectSummary = new ArrayList<>();
	public static  List<String> list_names = new ArrayList(); 
	public static  List<String> list_names1 = new ArrayList(); 
	public static List<String> finalListDataSummary = new ArrayList();
	public static List<String> list_statuses = new ArrayList<>();
	public static List<String> list_summaries = new ArrayList<>();
	public static List<String> masterSummaryList = new ArrayList();
	public static List<String> masterStatusList = new ArrayList();
	public static List<String> masterDefectList = new ArrayList();

	public static List<String> finalList = new ArrayList();
	public static List<String> finalList1 = new ArrayList();
	public static List<String> l_n1 =  new ArrayList();
	
//	static WriteExcel we=new WriteExcel();
	
	public static Auth getset = new Auth();
	public static List<String> SeverityList = new ArrayList();
	public static String jiraURL = "https://sourcefuse.atlassian.net/";
	public static String severity;
	public static String jiraID;
	public static void main(String[] args) 
	{
				new report_CycleName();
	}
	public static String getBasicAuth(String username, String password) {
		String valueToEncode = username + ":" + password;
		String auth="Basic " + Base64.getEncoder().encodeToString(valueToEncode.getBytes());
		return auth;
	}
	
	public static String getJiraID() throws IOException
	{   
		defects_rowsize = masterDefectList.size();
		
		for(int i = 0 ;i <masterDefectList.size() ;i++)
		{
			String jiraID=masterDefectList.get(i).toString();	
//			System.out.println(jiraID);
			String sev1= getAllSeverity(jiraID);
//			System.out.println("In getJiraID sev:"+sev1);
		 }
		return severity;
	}
	public static String getAllSeverity(String jiraID) throws IOException
	{         
		try {
      		JWTID = JWT_Token_Generator.key1(jiraID);

			RequestSpecification reportrequest = RestAssured.given().baseUri(jiraURL)
					.header("Content-Type", "application/json").header("Accept", "application/json")
//	    			.header("Authorization", JWTID)
					.header("Authorization", getBasicAuth(getset.getusername(), getset.getpassword()));
//					.header("Authorization", "shivraj.wadkar@sourcefuse.com:ATATT3xFfGF0BPzguUqsoDTw6j_DJXFmKGKYHU_RRlDyGI4Jm3HmRvYon8ZEYai1rEaHoPI6IcPFQhqD1DtGYNgAMV9xt9JUqxFJ9bPHGQYLv-SvX9mweGq8G33fHoKnqCQ_oNMd4Me8yQyWPaSxLvFn_b7plbK9uzLM0jiBt8zD6Z6VKiunIn8=67F09A2F")
//			        .header("zapiAccessKey", accessKey1);
			String issueEndpoint = "rest/api/3/issue/" + jiraID + "?expand=changelog&fields=*all";
			String getResponse = reportrequest.get(issueEndpoint).asString();
//			System.out.println("Response body : "+ getResponse);
//			Response response = reportrequest.body(getResponse).get();
//			System.out.println("Response Status Code is: "+response.getStatusCode());
//			String expandTag = reportrequest.get(issueEndpoint).getBody().jsonPath().getString("expand").toString();
//			System.out.println(expandTag);

			jp = new JsonPath(getResponse.toString()); 
			severity = jp.getString("fields.customfield_12819.value");
			
			if(severity!=null) 
			{		
				return severity;	
			}
			else 
			{
				return severity="Null";
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
		return severity;	
	}
	
	
	
	@SuppressWarnings("null")
public void GenerateZephyrReport_GetCycleReport_Zephyr(String PID, String VID ,String CN) {
	
	System.out.println("Fetching the report....");
	projectId = PID;
	cycleName = CN ;
	versionId = VID;

	WriteExcel we=new WriteExcel();
	we.writeExcel(1, VID, 0);
	we.writeExcel(1 ,PID, 1); 
	we.writeExcel(1 ,CN, 2); 

	 accessKey = "amlyYTpkZWM5MzA4OC02YmRhLTQ0ZDAtOWM0YS1hMjI5M2Q1MTc4OTQgNjI5MDZiZTQxZTRmNGIwMDY4MWRmNjdjIFVTRVJfREVGQVVMVF9OQU1F";

	allCycleIdsUrl = "https://prod-api.zephyr4jiracloud.com/connect/public/rest/api/1.0/cycles/search?versionId="+versionId+"&projectId="+projectId+"";
	
	getAllCycleIds();

}

public static void getAllCycleIds()
{

	try {
		JWTID = JWT_Token_Generator.key(allCycleIdsUrl);

		RequestSpecification reportrequest = RestAssured.given().baseUri(allCycleIdsUrl)
				.header("Content-Type", "application/json").header("Authorization", JWTID)
				.header("zapiAccessKey", accessKey);
		String getResponse = reportrequest.queryParam("versionId", versionId).queryParam("projectId", projectId).get()
				.asString();
		Response response = reportrequest.body(getResponse).get();

		jp = new JsonPath(response.asString()); 
	
	String cyclename = jp.getString("name");

	String[] cyclenameArray = cyclename.split(",");
	int size = cyclenameArray.length;

	String cycleID = jp.getString("id");
	String[] cycleIDArray = cycleID.split(",");
	 String[] newcycleIDArray = new String[size];
	 String[] newcycleNameArray = new String[size] ;
	for(int i=0 ; i<size ; i++)
	{
	 if(i == 0) {
	String val_ID = cycleIDArray[i].replace('[', ' ').trim().toString();
	 newcycleIDArray[i] =val_ID;
	 String val_Name = cyclenameArray[i].replace('[', ' ').trim().toString();
	 newcycleNameArray[i] = val_Name;
	 }
	 
	 else if(i == size-1)
	 {
		 String val_ID = cycleIDArray[i].replace(']', ' ').trim().toString();
		 newcycleIDArray[i] =val_ID;
		 String val_Name = cyclenameArray[i].replace(']', ' ').trim().toString();
		 newcycleNameArray[i] = val_Name;
	 }
	 
	 else
	 {
		 String val_ID = cycleIDArray[i].toString();
		 newcycleIDArray[i] = val_ID;
		 String val_Name = cyclenameArray[i].toString();
		 newcycleNameArray[i] = val_Name;
	
	 }
		
	}
	WriteExcel we=new WriteExcel();
	for(int i=0 ; i<size ; i++)
	{
	 if(newcycleNameArray[i].toString().trim().equals(cycleName))
	 {
		 we.writeExcel(1 , newcycleIDArray[i].toString().trim(), 3);
	 getAllFolderIdsFromOneCycleID(newcycleIDArray[i].toString().trim() , newcycleNameArray[i].toString() );
	 }
	}
	
	
	
	
	} catch (URISyntaxException e) {
		e.printStackTrace();
	}
	

}



public static void getAllFolderIdsFromOneCycleID(String cId ,  String cName)
{

	try {
		allFolderIdsUrl  = 	"https://prod-api.zephyr4jiracloud.com/connect/public/rest/api/1.0/folders?versionId="+versionId+"&cycleId="+cId+"&projectId="+projectId+"";

		JWTID = JWT_Token_Generator.key(allFolderIdsUrl);

		RequestSpecification reportrequest = RestAssured.given().baseUri( allFolderIdsUrl)
				.header("Content-Type", "application/json").header("Authorization", JWTID)
				.header("zapiAccessKey", accessKey);
		String getResponse = reportrequest.queryParam("versionId", versionId).queryParam("cycleId", cId).queryParam("projectId", projectId).get()
				.asString();
		Response response = reportrequest.body(getResponse).get();

		jp = new JsonPath(response.asString()); 
	
		String foldername = jp.getString("name").toString();
		String[] foldernameArray = foldername.split(",");

		 String folderID = jp.getString("id").toString();
		String[] folderIDArray = folderID.split(",");
		
		
		
		int size = folderIDArray.length;
		folderCount = size;
		
		 String[] newFolderIDArray = new String[size];
		 String[] newFolderNameArray = new String[size] ;
		 

		for(int i=0 ; i<size ; i++)
		{
		 
		 if(i == 0) {
			 String val_ID = folderIDArray[i].replace('[', ' ').trim().toString();
			 newFolderIDArray[i] =val_ID;
		 String val_Name = foldernameArray[i].replace('[', ' ').trim().toString();
		 newFolderNameArray[i] = val_Name;
		 }
		 
		 else if(i == size-1)
		 {
			 String val_ID = folderIDArray[i].replace(']', ' ').trim().toString();
			 newFolderIDArray[i] =val_ID;
			 String val_Name = foldernameArray[i].replace(']', ' ').trim().toString();
			 newFolderNameArray[i] = val_Name;
		 }
		 else
		 {
			 String val_ID = folderIDArray[i].toString();
			 newFolderIDArray[i] = val_ID;
			 String val_Name = foldernameArray[i].toString();
			 newFolderNameArray[i] = val_Name;
		
		 }
			
		}
		 int folder_rowNo = 0;
		for(int i=0 ; i<size ; i++)
		{
		 System.out.println("Folder Name is: "+newFolderNameArray[i].toString() );
		 
			try {
				WriteExcel we=new WriteExcel();
				we.writeExcel(i+1 , newFolderNameArray[i].toString().trim(), 4);
				we.writeExcel(i+1 , newFolderIDArray[i].toString().trim(), 5);
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
 
		  folder_rowNo = i+1;
		 getTestCasesResults_folderwise(newFolderNameArray[i].toString().trim(), newFolderIDArray[i].toString().trim() ,  cId , folder_rowNo);
		
		}
		
	} catch (URISyntaxException e) {
		e.printStackTrace();
	}
	

}

@SuppressWarnings("unchecked")
public static void getTestCasesResults_folderwise(String fName, String fId, String cId ,  int folder_rowNo)
{
	
	try {
	
		fid_SummaryReportUrl = "public/rest/api/2.0/executions/search/folder/"+fId+"?projectId="+projectId+"&versionId="+versionId+"&cycleId="+cId+"";
		JWTID = JWT_Token_Generator.key(zephyrBaseUrl + fid_SummaryReportUrl);
	
				RequestSpecification reportrequest = RestAssured.given().baseUri(zephyrBaseUrl + fid_SummaryReportUrl)
				.header("Content-Type", "application/json").header("Authorization", JWTID)
				.header("zapiAccessKey", accessKey);
		String getResponse = reportrequest.queryParam("versionId",versionId).queryParam("cycleId", cId).queryParam("projectId", projectId).get()
				.asString();
		Response response = reportrequest.body(getResponse).get();
	jp = new JsonPath(response.asString());
	String data = jp.getString("folderSummary.executionSummaries");
	  int res_size = jp.getInt("folderSummary.executionSummaries.size()");
	  int row  = folderCount ;
	  WriteExcel we=new WriteExcel();
	  for( int status_row= 6 ; status_row<=13 ; status_row++)
	  {
		  we.writeExcel(folder_rowNo , "0", status_row);
	  }
	

	  String totalTestcases = jp.getString("folderSummary.totalExecutions");
	  {
		  we.writeExcel(folder_rowNo ,totalTestcases, 11);
	  }
	  String totalDefects = jp.getString("folderSummary.totalDefects");
	  we.writeExcel(row ,totalDefects, 12);
	  System.out.println("TotalTestcases: "+totalTestcases);
	  for(int i = 1; i<=res_size; i++)
	  {
		  int s = i-1;
	  
	Map<Object, Object> res = response.jsonPath().getMap("folderSummary.executionSummaries["+s+"]");
	
	String status= res.get("executionStatusName").toString();
	
	
	String count= res.get("count").toString();
	 if(status.equals("PASS")) 
	{
			 we.writeExcel(folder_rowNo , count, 6);
		
	}
	
	
	else if(status.equals("FAIL"))
	{
		we.writeExcel( folder_rowNo , count, 7);	
	}
	
	else if(status.equals("UNEXECUTED"))
	{
		we.writeExcel(folder_rowNo ,count, 8);
	}

	else if(status.equals("WIP"))
	{
		we.writeExcel(folder_rowNo ,count, 9);
	}
	
	else if(status.equals("BLOCKED"))
	{
		we.writeExcel(folder_rowNo ,count, 10);
	}
	
	else
	{	
		System.out.println("Not Working");
	}
}
	  
		String def_data = jp.getString("searchResult.searchObjectList");
		
		  int def_data_res_size = jp.getInt("searchResult.searchObjectList.size()");
	  
		  String def_name_array = "";
		 int def_name_array_size = 0; 

		 int tcNo = 0;
	  for( int j = 0 ; j <def_data_res_size ; j++)
	  {
		  
		  Map<Object, Object> res = response.jsonPath().getMap("searchResult.searchObjectList["+j+"].execution");
	
		String defnames= res.get("defects").toString();
		 
		 int defnamesArray_size = jp.getInt("searchResult.searchObjectList["+j+"].execution.defects.size()");
		 List<String> l_n = null ;
		 l_n1 = null ;
		 @SuppressWarnings("unus;ed")

		 List<String> l_s = null ;
		 List<String> l_summary = null;
		if(defnamesArray_size>0 )
		{
			for (int q = 0 ; q <defnamesArray_size ; q++ )
			{
				
			
			  Map<Object, Object> resdef = response.jsonPath().getMap("searchResult.searchObjectList["+j+"].execution.defects["+q+"]");
		
			  String key_resdefnames= resdef.get("key").toString();

			

			  Map<Object, Object> resdef_status = response.jsonPath().getMap("searchResult.searchObjectList["+j+"].execution.defects["+q+"].status");
			  String status_resdefnames= resdef_status.get("name").toString();
			  
			  Map<Object, Object> resdef_summary = response.jsonPath().getMap("searchResult.searchObjectList["+j+"].execution.defects["+q+"]");
			  String summary_resdefsummary= resdef_summary.get("summary").toString();
			
			  l_n =  ListOfDefects_Names(key_resdefnames);

			  l_n1 =  ListOfDefects_Names1(summary_resdefsummary);
			  l_s =  ListOfDefects_Status(status_resdefnames);
			  l_summary = ListOfDefects_Summary(summary_resdefsummary);
			 folderwise_defects(l_n , l_n1 , l_s , l_summary, folder_rowNo);
			}
		}

	  }
	

	}
	
	
	catch (Exception e) {
		e.printStackTrace();
	}
	  System.out.println("Folder wise report is fetched for: "+fName);

	  last_rowCount = folder_rowNo;
	  System.out.println("");
	  list_names.clear();
	  finalList1.clear();
	  list_names1.clear();
	  list_statuses.clear();
	  list_summaries.clear();
	  ListOfDefectNames.clear();
	  ListOfDefectStatus.clear();
	  ListOfDefectSummary.clear();
	
	}

public static List<String> ListOfDefects_Names(String key_resdefnames) {

	ListOfDefectNames.add(key_resdefnames);

	return ListOfDefectNames ;
}

public static List<String> ListOfDefects_Names1(String key_resdefnames1) {
	ListOfDefectNames1.add(key_resdefnames1);
	return ListOfDefectNames1 ;
}

public static List<String> ListOfDefects_Status(String status_resdefnames) 
{
	ListOfDefectStatus.add(status_resdefnames);

	return ListOfDefectStatus ;
}

public static List<String> ListOfDefects_Summary(String summary_resdefsummary) {
	
	ListOfDefectSummary.add(summary_resdefsummary);
	return ListOfDefectSummary ;
}

 
public static void folderwise_defects(List<String> n, List<String> n1 , List<String> s , List<String> summary, int folder_rowNo) throws IOException
{
	try {
		list_names.clear();
		list_names1.clear();
		list_statuses.clear();
		list_summaries.clear();


	for(int l = 0 ; l <s.size() ; l++)
	{
		if(s.get(l).equals(defectStatus) || s.get(l).equals(defectStatus1))
		{

		}
		else
		{
			
			int t = l;

			list_names.add(n.get(t));
			list_names1.add(n1.get(t));
			list_statuses.add(s.get(t));
			
			list_summaries.add(summary.get(t));
		}
		
	}

	for(int i = 0 ; i<list_names.size() ; i++)
	{
		setAllDefectList(i ,list_names.get(i).toString() );
	}


	for(int i = 0 ; i<list_summaries.size() ; i++)
	{
		setAllSummaryList(i ,list_summaries.get(i).toString() );
	}

	list_summaries.clear();
	for(int i = 0 ; i<list_statuses.size() ; i++)
	{
		setAllStatusList(i ,list_statuses.get(i).toString() );

	}

	//for(int w = 1 ; w<=folderCount ; w++) {
	int def_Count = list_names.size();
	int def_Count1 = list_names1.size();
	String count = Integer.toString(def_Count);
	String count1 = Integer.toString(def_Count1);
	WriteExcel we=new WriteExcel();
	finalList.add(list_names.toString().trim().replace('[', ' ').replace(']', ' '));
	finalList1.add(list_names1.toString().trim().replace('[', ' ').replace(']', ' '));
	
	we.writeExcel(folder_rowNo,finalList.toString().trim().replace('[', ' ').replace(']', ' '), 13);
	
	finalList.clear();
	finalList1.clear();
	int fname_sizee = list_names.size();
	int fsum_size = list_summaries.size();
	int fstat_size = list_statuses.size();
	if(fname_sizee==0) 
	{
		we.writeExcel(folder_rowNo , "0" , 13);
		we.writeExcel(folder_rowNo,Integer.toString(fname_sizee), 12);
		
	}
	else	
	{
	 we.writeExcel(folder_rowNo,Integer.toString(fname_sizee), 12);
	}


	list_names.clear();
	list_names1.clear();
	} catch (Exception e) {
		e.printStackTrace();
	}
}


public static void setAllDefectList(int i ,String value )
{
	if(masterDefectList.contains(value))
	{
		
	}
	else {
	masterDefectList.add(i, value);
	}

}
public static void getvalueAllDefectList() throws IOException
{
	defects_rowsize = masterDefectList.size();
	for(int i = 0 ;i <masterDefectList.size() ;i++)
	{
		
		WriteExcel.writeExcel2(i , 0 , masterDefectList.get(i).toString() );
	}
	for(int i = 0 ;i <masterDefectList.size() ;i++)
	{
		String jiraID=masterDefectList.get(i).toString();	
		String sev1= getAllSeverity(jiraID);
		{
			WriteExcel.writeExcel2(i , 3 , sev1); 
		}		
	}
	
}

	public static void setAllSummaryList(int i ,String value )
	{
		
		if(masterSummaryList.contains(value))
		{
			
		}
		else {
		masterSummaryList.add(i, value);
		}
	
	}
	public static void getvalueAllSummaryList()
	{
		for(int i = 0 ;i <masterSummaryList.size() ;i++)
		{
			WriteExcel.writeExcel2(i , 1, masterSummaryList.get(i).toString() );
		}
		
	}
	

	public static void setAllStatusList(int i ,String value )
	{
		masterStatusList.add(i, value);
	
	}
	public static void getvalueAllStatusList()
	{
		for(int i = 0 ;i <masterStatusList.size() ;i++)
		{
			WriteExcel.writeExcel2(i , 2, masterStatusList.get(i).toString() );
			
		}
//		masterStatusList.clear();
	}
	
}