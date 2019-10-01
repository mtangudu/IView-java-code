package com.deloitte.exceltojson.util;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.Map;

@SuppressWarnings("unchecked")
public class  Calculate {
	
	public static int getTimeOfFunction (LinkedHashMap<String, Object> nodeData)
	{
		LinkedHashMap<String, Object> nodeDetails = (LinkedHashMap<String, Object>) nodeData.get("details");
		int dbCall = convertStringToInt(nodeDetails.get("DB Calls (ms)"));
		int logic = convertStringToInt(nodeDetails.get("Logic (ms)"));
		int rules =convertStringToInt(nodeDetails.get("Rules (ms)"));
		int latency =convertStringToInt(nodeDetails.get("Latency (ms)"));
		return dbCall + logic+ rules+ latency;
	}
	public static String getTotalProcessPathLenghtNFRForParent (int totalProcessNFR, ArrayList<LinkedHashMap<String, Object>> children)
	{
		int nfr=0;
		for (Map<String, Object> c: children){
			Map<String, Object> nd = (Map<String, Object>) c.get("details");
			if (nd.get("Total Process Path Length NFR") != null)
				nfr = nfr + Integer.valueOf(nd.get("Total Process Path Length NFR").toString());
		}
		return String.valueOf(nfr + totalProcessNFR);
	}
	
	public static String getTotalProcessPathLenghtForParent (int totalProcessPathLength, ArrayList<LinkedHashMap<String, Object>> children)
	{
		int nfr=0;
		for (Map<String, Object> c: children){
			Map<String, Object> nd = (Map<String, Object>) c.get("details");
			if (nd.get("Total Process Path Length") != null)
				nfr = nfr + Integer.valueOf(nd.get("Total Process Path Length").toString());
		}
		return String.valueOf(nfr + totalProcessPathLength);
	}
	public static int convertStringToInt(Object valueToConvert)
	{
		if (valueToConvert == null)
			return 0;
		else
		return Integer.valueOf(valueToConvert.toString());
	}
}
