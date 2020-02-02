package com.deloitte.exceltojson.processor;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Properties;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Export{

	public LinkedList<Map<String,String>> getDataAsList(Map<String, Object> data)
	{
		LinkedList<Map<String,String>> ll = new LinkedList<Map<String,String>>();
		Map <String,String> singleRowData = new HashMap<String, String>();
		for (Map.Entry<String,Object> dataEntry : data.entrySet())
		{
			System.out.println("dataEntry.getKey() : " + dataEntry.getKey());
			if (dataEntry.getKey().equalsIgnoreCase("details"))
			{
				Map <String,Object> singleRowDetails = (Map<String, Object>) dataEntry.getValue();
				System.out.println("inside details singleRowDetails " + singleRowDetails);
				for (Map.Entry<String,Object> detailsEntry : singleRowDetails.entrySet())
				{
						singleRowData.put(detailsEntry.getKey(), detailsEntry.getValue().toString());
				}
			}
			else if (dataEntry.getKey().equalsIgnoreCase("children"))
			{
				ArrayList<Map<String, Object>> children =  (ArrayList<Map<String, Object>>) dataEntry.getValue();
				System.out.println("size of ll b4 for : " + ll.size());
				for (int i=0; i<children.size(); i++)
				{
					ll.addAll(getDataAsList(children.get(i)));
				}
				System.out.println("size of ll a4 for : " + ll.size());
			}
			else
			{
				singleRowData.put(dataEntry.getKey(), dataEntry.getValue().toString());
			}
		}
		ll.add(singleRowData);
		System.out.println("ll :::::::: "+ll);
		System.out.println("final size is :::::::::: " + ll.size());
		for (int x=0; x<ll.size(); x++)
		{
			System.out.println("x value " + x);
			System.out.println("ll value " + ll.get(x));
		}
		return ll;
	}
	
	
	private void writeBook(Map<String,String> rowData, Row row,Properties fieldProperties) {
	    Cell cell = row.createCell(0);
	    cell.setCellValue(rowData.get(fieldProperties.getProperty("text.data.serialNoIndex")));
	    
	    cell = row.createCell(1);cell.setCellValue(rowData.get(fieldProperties.getProperty("text.data.linkIndex")));
	    cell = row.createCell(2);cell.setCellValue(rowData.get(fieldProperties.getProperty("text.data.name")));
	    cell = row.createCell(3);cell.setCellValue(rowData.get(fieldProperties.getProperty("text.data.description")));
	    cell = row.createCell(4);cell.setCellValue(rowData.get(fieldProperties.getProperty("text.data.etaStatus")));
	    cell = row.createCell(5);cell.setCellValue(rowData.get(fieldProperties.getProperty("text.data.category")));
	    cell = row.createCell(6);cell.setCellValue(rowData.get(fieldProperties.getProperty("text.data.connectionType")));
	    cell = row.createCell(7);cell.setCellValue(rowData.get(fieldProperties.getProperty("text.data.details.process")));
	    cell = row.createCell(8);cell.setCellValue(rowData.get(fieldProperties.getProperty("text.data.details.owner")));
	    cell = row.createCell(9);cell.setCellValue(rowData.get(fieldProperties.getProperty("text.data.details.hostedName")));
	    cell = row.createCell(10);cell.setCellValue(rowData.get(fieldProperties.getProperty("text.data.details.provider")));
	    cell = row.createCell(11);cell.setCellValue(rowData.get(fieldProperties.getProperty("text.data.details.consumer")));
	    cell = row.createCell(12);cell.setCellValue(rowData.get(fieldProperties.getProperty("text.data.details.integrationPattern")));
	    cell = row.createCell(13);cell.setCellValue(rowData.get(fieldProperties.getProperty("text.data.details.confidenceFactor")));
	    cell = row.createCell(14);cell.setCellValue(rowData.get(fieldProperties.getProperty("text.data.details.countOfInvocations")));
	    cell = row.createCell(15);cell.setCellValue(rowData.get(fieldProperties.getProperty("text.data.details.dbCalls")));
	    cell = row.createCell(16);cell.setCellValue(rowData.get(fieldProperties.getProperty("text.data.details.logic")));
	    cell = row.createCell(17);cell.setCellValue(rowData.get(fieldProperties.getProperty("text.data.details.rules")));
	    cell = row.createCell(18);cell.setCellValue(rowData.get(fieldProperties.getProperty("text.data.details.latency")));
	    cell = row.createCell(19);cell.setCellValue(rowData.get(fieldProperties.getProperty("text.data.details.ProcessNFR")));
	    cell = row.createCell(20);cell.setCellValue(rowData.get(fieldProperties.getProperty("text.data.details.timeOfFuncation")));
	    cell = row.createCell(21);cell.setCellValue(rowData.get(fieldProperties.getProperty("text.data.details.totalProcessTime")));
	    cell = row.createCell(22);cell.setCellValue(rowData.get(fieldProperties.getProperty("text.data.details.totalProcessPathLength")));
	    cell = row.createCell(23);cell.setCellValue(rowData.get(fieldProperties.getProperty("text.data.details.totalProcessNFR")));
	    cell = row.createCell(24);cell.setCellValue(rowData.get(fieldProperties.getProperty("text.data.details.totalProcessPathLengthNFR")));
	    
	}
	public void writeExcel(LinkedList<Map<String,String>> rows, Properties fieldProperties, String filePath) throws IOException {
	    Workbook workbook = new XSSFWorkbook();
	    Sheet sheet = workbook.createSheet();
	 
	    int rowCount = 0;
	 
	    for (Map<String,String> rowData : rows) {
	        Row row = sheet.createRow(++rowCount);
	        writeBook(rowData, row,fieldProperties);
	    }
	    System.out.println("writing into file path " + filePath);
	    try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
	        workbook.write(outputStream);
	        workbook.close();
	    }
	}
	
}
