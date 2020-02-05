package com.deloitte.exceltojson.processor;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.Map;
import java.util.Properties;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.deloitte.exceltojson.pojo.MetaDatCode;

public class Export{

	public Map<String, LinkedList<MetaDatCode>> getProcessedMetaData(Map<String, Object> metaData)
	{
		Map<String,LinkedList<MetaDatCode>> mataDataDetailsInList = new HashMap<String,LinkedList<MetaDatCode>>();
		mataDataDetailsInList.put("etaStatus", getMetaDataAsList(metaData, "etaStatus"));
		mataDataDetailsInList.put("connectionType", getMetaDataAsList(metaData, "connectionType"));
		mataDataDetailsInList.put("category", getMetaDataAsList(metaData, "category"));
		return mataDataDetailsInList;
	}
	
	public LinkedList<MetaDatCode> getMetaDataAsList (Map<String,Object> metaData, String metaDataKey)
	{
		Map<String,Object> metaDataKeyValue = (Map<String, Object>) metaData.get(metaDataKey);
		Map<String,Object> items = (Map<String, Object>) metaDataKeyValue.get("items");
		LinkedList<MetaDatCode> codeList =  new LinkedList<MetaDatCode>();
		for (Map.Entry<String,Object> itemsEntry : items.entrySet())
		{
			String itemKey = itemsEntry.getKey();
			Map<String,String> itemValue = (Map<String, String>) itemsEntry.getValue();
			codeList.add(new MetaDatCode(itemKey,itemValue.get("color"),itemValue.get("description")));
		}
		return codeList;
	}

	public LinkedList<Map<String,String>> getDataAsList(Map<String, Object> data)
	{
		LinkedList<Map<String,String>> ll = new LinkedList<Map<String,String>>();
		Map <String,String> singleRowData = new HashMap<String, String>();
		for (Map.Entry<String,Object> dataEntry : data.entrySet())
		{
			if (dataEntry.getKey().equalsIgnoreCase("details"))
			{
				Map <String,Object> singleRowDetails = (Map<String, Object>) dataEntry.getValue();
				for (Map.Entry<String,Object> detailsEntry : singleRowDetails.entrySet())
						singleRowData.put(detailsEntry.getKey(), detailsEntry.getValue().toString());
			}
			else if (dataEntry.getKey().equalsIgnoreCase("children"))
			{
				ArrayList<Map<String, Object>> children =  (ArrayList<Map<String, Object>>) dataEntry.getValue();
				for (int i=0; i<children.size(); i++)
					ll.addAll(getDataAsList(children.get(i)));
			}
			else
				singleRowData.put(dataEntry.getKey(), dataEntry.getValue().toString());
		}
		ll.add(singleRowData);
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
	public byte[] writeExcel(LinkedList<Map<String,String>> rows, Map<String,LinkedList<MetaDatCode>> processedMetaData,
			Properties fieldProperties) throws IOException {
	    Workbook workbook = new XSSFWorkbook();
	    Sheet dataSheet = workbook.createSheet("Data");
	    Sheet metaDataSheet = workbook.createSheet("meta Data");
	    //sheet.addMergedRegion(new CellRangeAddress(1,1,1,4));
	 
	    
	    createDataHeader(workbook, dataSheet, fieldProperties);
	    createMetaDataHeader(workbook, metaDataSheet, fieldProperties,processedMetaData);
	    int rowCount = 0;
	    for (Map<String,String> rowData : rows) {
	        Row row = dataSheet.createRow(++rowCount);
	        writeBook(rowData, row,fieldProperties);
	    }
	    ByteArrayOutputStream bos = new ByteArrayOutputStream();
	    try {
	        workbook.write(bos);
	    } finally {
	        bos.close();
	    }
	    byte[] bytes = bos.toByteArray();
	    return bytes;
	}
	
	public void createDataHeader(Workbook workbook, Sheet sheet,Properties fieldProperties)
	{
        // Create a Font for styling header cells
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.BLACK1.getIndex());
        // Create a CellStyle with the font
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);
        headerCellStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // Create a Row
        Row headerRow = sheet.createRow(0);

        // Create cells
        String[] headerTitle = {fieldProperties.getProperty("text.data.serialNoIndex"), fieldProperties.getProperty("text.data.linkIndex"), fieldProperties.getProperty("text.data.name"),fieldProperties.getProperty("text.data.description"),fieldProperties.getProperty("text.data.etaStatus"),fieldProperties.getProperty("text.data.category"),fieldProperties.getProperty("text.data.connectionType"), fieldProperties.getProperty("text.data.details.process"), fieldProperties.getProperty("text.data.details.owner"), fieldProperties.getProperty("text.data.details.hostedName"), fieldProperties.getProperty("text.data.details.provider"), fieldProperties.getProperty("text.data.details.consumer"), fieldProperties.getProperty("text.data.details.integrationPattern"), fieldProperties.getProperty("text.data.details.confidenceFactor"), fieldProperties.getProperty("text.data.details.countOfInvocations"), fieldProperties.getProperty("text.data.details.dbCalls"), fieldProperties.getProperty("text.data.details.logic"), fieldProperties.getProperty("text.data.details.rules"), fieldProperties.getProperty("text.data.details.latency"), fieldProperties.getProperty("text.data.details.ProcessNFR"), fieldProperties.getProperty("text.data.details.timeOfFuncation"), fieldProperties.getProperty("text.data.details.totalProcessTime"), fieldProperties.getProperty("text.data.details.totalProcessPathLength"), fieldProperties.getProperty("text.data.details.totalProcessNFR"), fieldProperties.getProperty("text.data.details.totalProcessPathLengthNFR")};
        for(int i = 0; i < headerTitle.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headerTitle[i]);
            cell.setCellStyle(headerCellStyle);
        }
	}
	
	
	public void createMetaDataHeader(Workbook workbook, Sheet metaDataSheet,Properties fieldProperties,
			Map<String,LinkedList<MetaDatCode>> processedMetaData)
	{
        // Create a Font for styling header cells
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.BLACK1.getIndex());
        // Create a CellStyle with the font
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);
        headerCellStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerCellStyle.setAlignment(HorizontalAlignment.CENTER);
        
        
        Font metaDataLine2headerFont = workbook.createFont();
        metaDataLine2headerFont.setBold(true);
        metaDataLine2headerFont.setFontHeightInPoints((short) 12);
        metaDataLine2headerFont.setColor(IndexedColors.BLACK1.getIndex());
        CellStyle metaDataLine2headerCellStyle = workbook.createCellStyle();
        metaDataLine2headerCellStyle.setFont(metaDataLine2headerFont);
        metaDataLine2headerCellStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        metaDataLine2headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        metaDataLine2headerCellStyle.setAlignment(HorizontalAlignment.CENTER);

        Row headerRow = metaDataSheet.createRow(0);
        metaDataSheet.addMergedRegion(new CellRangeAddress(0,0,0,2));
        metaDataSheet.addMergedRegion(new CellRangeAddress(0,0,4,6));
        metaDataSheet.addMergedRegion(new CellRangeAddress(0,0,8,10));
        
        Cell cell = headerRow.createCell(0);
        cell.setCellStyle(headerCellStyle);
        cell.setCellValue("eta status");
        
        Cell cellConnectionHDR = headerRow.createCell(4);
        cellConnectionHDR.setCellStyle(headerCellStyle);
        cellConnectionHDR.setCellValue("Connection Type");
        
        Cell cellCategoryHDR = headerRow.createCell(8);
        cellCategoryHDR.setCellStyle(headerCellStyle);
        cellCategoryHDR.setCellValue("Category");
        
        LinkedList<MetaDatCode> etaList = processedMetaData.get("etaStatus");
	    LinkedList<MetaDatCode> categoryList = processedMetaData.get("category");
	    LinkedList<MetaDatCode> connectionTypeList = processedMetaData.get("connectionType");
	    int noOfMaxObjInList = getMaxLength(etaList.size(), categoryList.size(), connectionTypeList.size());
	    
	    int totalRowCount = 1;
	    // add metedata headers 
	    Row metaDataHeader = metaDataSheet.createRow(1);
	    writeMetaDataHeaderLine2(metaDataHeader,metaDataLine2headerCellStyle);
	    
	    for (int x =0; x<=noOfMaxObjInList; x++)
	    {
	    	MetaDatCode etaData = null;
		    MetaDatCode connectionTypeData = null;
		    MetaDatCode categoryListData = null;
	    	Row row = metaDataSheet.createRow(++totalRowCount);
	    	if (etaList.size()-1>=x)
	    		etaData = etaList.get(x);
	    	if (connectionTypeList.size()-1>=x)
	    		connectionTypeData = connectionTypeList.get(x);
	    	if (categoryList.size()-1>=x)
	    		categoryListData = categoryList.get(x);
	    	
	    	writeMetaDataInRow(etaData,connectionTypeData, categoryListData, row );
	    }
	}
	
	
	private void writeMetaDataHeaderLine2 ( Row metaDataHeader, CellStyle headerCellStyle) {
		
		int etaSecondColumn=1, etaThirdColumn=2, 
				connenctionFirstColumn=4, connectionSecondColumn=5, connectionThirdColumn=6,
				categoryFirstColumn=8, categorySecondColumn=9, categoryThirdColumnn=10;
		
	    Cell cell = metaDataHeader.createCell(0); 
	    cell.setCellStyle(headerCellStyle); cell.setCellValue("ETA Status");
	    cell = metaDataHeader.createCell(etaSecondColumn);cell.setCellStyle(headerCellStyle);cell.setCellValue("color");
	    cell = metaDataHeader.createCell(etaThirdColumn);cell.setCellStyle(headerCellStyle);cell.setCellValue("description");
	    
	    cell = metaDataHeader.createCell(connenctionFirstColumn);cell.setCellStyle(headerCellStyle);cell.setCellValue("Connection Type");
	    cell = metaDataHeader.createCell(connectionSecondColumn);cell.setCellStyle(headerCellStyle);cell.setCellValue("description");
	    cell = metaDataHeader.createCell(connectionThirdColumn);cell.setCellStyle(headerCellStyle);cell.setCellValue("color");
	    cell = metaDataHeader.createCell(categoryFirstColumn);cell.setCellStyle(headerCellStyle);cell.setCellValue("category");
	    cell = metaDataHeader.createCell(categorySecondColumn);cell.setCellStyle(headerCellStyle);cell.setCellValue("color");
	    cell = metaDataHeader.createCell(categoryThirdColumnn);cell.setCellStyle(headerCellStyle);cell.setCellValue("Description");
	}
	
	private void writeMetaDataInRow(MetaDatCode eta, MetaDatCode connection, MetaDatCode category,Row row) {
		
		int etaFirstColumn = 0, etaSecondColumn=1, etaThirdColumn=2, 
				connenctionFirstColumn=4, connectionSecondColumn=5, connectionThirdColumn=6,
				categoryFirstColumn=8, categorySecondColumn=9, categoryThirdColumnn=10;
		Cell cell = row.createCell(etaFirstColumn);
		if (eta!= null){
			cell.setCellValue(eta.getCode());
			cell = row.createCell(etaSecondColumn);cell.setCellValue(eta.getColor());
		    cell = row.createCell(etaThirdColumn);cell.setCellValue(eta.getDescription());
		}
		if (connection!= null){
			cell = row.createCell(connenctionFirstColumn);cell.setCellValue(connection.getCode());
			cell = row.createCell(connectionSecondColumn);cell.setCellValue(connection.getColor());
		    cell = row.createCell(connectionThirdColumn);cell.setCellValue(connection.getDescription());
		}
		if (category!= null){
			cell = row.createCell(categoryFirstColumn);cell.setCellValue(category.getCode());
			cell = row.createCell(categorySecondColumn);cell.setCellValue(category.getColor());
		    cell = row.createCell(categoryThirdColumnn);cell.setCellValue(category.getDescription());
		}
	}

	private static int getMaxLength (int etaLen, int categoryLen, int connectionLen){

	      if( etaLen >= categoryLen && etaLen >= connectionLen)
	          return etaLen;
	      else if (categoryLen >= etaLen && categoryLen >= connectionLen)
	          return categoryLen;
	      else
	          return connectionLen;
	  }
}
