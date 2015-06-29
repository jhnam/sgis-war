<%@page import="java.io.File"%>
<%@page import="java.io.FileInputStream"%>
<%@page import="org.apache.poi.hssf.usermodel.HSSFRow"%>
<%@page import="org.apache.poi.hssf.usermodel.HSSFSheet"%>
<%@page import="org.apache.poi.hssf.usermodel.HSSFWorkbook"%>
<%@page import="org.apache.poi.xssf.usermodel.XSSFRow"%>
<%@page import="java.util.UUID"%>
<%@page import="org.apache.poi.xssf.usermodel.XSSFSheet"%>
<%@page import="java.io.FileOutputStream"%>
<%@page import="com.google.gson.internal.LinkedTreeMap"%>
<%@page import="org.apache.poi.ss.usermodel.Workbook"%>
<%@page import="org.apache.poi.xssf.usermodel.XSSFWorkbook"%>
<%@ page language="java" contentType="text/html; charset=utf-8" pageEncoding="utf-8"%>
<%@page import="java.util.HashMap"%>
<%@page import="java.util.List"%>
<%@page import="java.util.ArrayList"%>
<%@page import="com.google.gson.Gson"%>



<%
	response.setHeader("Access-Control-Allow-Origin","*");
	response.setHeader("Access-Control-Allow-Headers", "origin, x-requested-with, content-type, accept");
	try{
		Gson gson = new Gson();
		String fileName = request.getParameter("fileName");
		
		if(fileName==null){
			ArrayList headList = gson.fromJson(request.getParameter("header"), ArrayList.class);
			ArrayList dataList = gson.fromJson(request.getParameter("datas"), ArrayList.class);
			
			HSSFWorkbook workbook = new HSSFWorkbook();
			HSSFSheet sheet = workbook.createSheet("검색결과");
			
			for(int i=0; i<dataList.size(); i++){
				LinkedTreeMap map = (LinkedTreeMap)dataList.get(i);
				HSSFRow row = sheet.createRow(i);
				for(int k=0; k<headList.size(); k++){
					String head = (String)headList.get(k);
					if(map.get(head)==null){
						row.createCell(k).setCellValue("");
					}else{
						row.createCell(k).setCellValue(map.get(head)+"");
					}
				}
			}
			String randomId =  "excel_" + UUID.randomUUID().toString() + ".xls";
			FileOutputStream outFile;
			outFile = new FileOutputStream("C:\\arcgisserver\\directories\\arcgisoutput\\customPrintTask\\" + randomId);
			workbook.write(outFile);
			outFile.close();
			
			HashMap hashMap = new HashMap();
			hashMap.put("url", "http://" + request.getServerName()+ ":" + request.getServerPort() + request.getContextPath() + request.getServletPath() + "?fileName=" + randomId);
	 		out.println(gson.toJson(hashMap));
		}else{
			
			
			File file = new File("C:\\arcgisserver\\directories\\arcgisoutput\\customPrintTask\\" + fileName);
			FileInputStream fin = new FileInputStream(file);
			int ifilesize = (int)file.length();
			byte b[] = new byte[ifilesize];
			response.setContentLength(ifilesize);
			response.setContentType("application/octet-stream");
			response.setHeader("Content-Disposition","attachment; filename="+fileName+";");
			ServletOutputStream oout = response.getOutputStream();
			fin.read(b);
			oout.write(b,0,ifilesize);
			oout.flush();
			oout.close();
			fin.close();
			fin = null;
			oout = null;
			Runtime.getRuntime().gc();		
		}
	}catch(Exception e){
		e.printStackTrace();
	}
%>