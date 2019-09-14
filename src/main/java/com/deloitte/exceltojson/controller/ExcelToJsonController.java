package com.deloitte.exceltojson.controller;

import java.io.ByteArrayOutputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

import javax.servlet.http.HttpServletResponse;
import javax.websocket.server.PathParam;

import org.apache.log4j.BasicConfigurator;
import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.bson.Document;
import org.json.JSONObject;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.deloitte.exceltojson.pojo.ExcelReader;
import com.deloitte.exceltojson.pojo.NodeData;
import com.deloitte.exceltojson.processor.Constants;
import com.deloitte.exceltojson.processor.MessageProcessor;
import com.deloitte.exceltojson.repo.NodeDataRepository;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.mongodb.BasicDBObject;
import com.mongodb.DBCursor;
import com.mongodb.MongoClient;
import com.mongodb.client.FindIterable;
import com.mongodb.client.MongoCollection;
import com.mongodb.client.MongoDatabase;
import com.mongodb.client.model.Filters;

@RestController
public class ExcelToJsonController {

	final static Logger log = Logger.getLogger(ExcelToJsonController.class);
	
	@CrossOrigin(origins = "*")
	@GetMapping("/getSL")
	public ResponseEntity<String> getSLDetails() throws JsonProcessingException {
		ObjectMapper mapper = new ObjectMapper();
		String jsonData = new String();
		InputStream propertiesInput = ExcelToJsonController.class.getClassLoader()
				.getResourceAsStream("application.properties");
		Properties appProperties = new Properties();

		try {
			appProperties.load(propertiesInput);
			MongoClient mongoClient = new MongoClient(appProperties.getProperty("mongodb.host"),
					new Integer(appProperties.getProperty("mongodb.port")));
			MongoDatabase db = mongoClient.getDatabase(appProperties.getProperty("mongodb.db"));
			MongoCollection<Document> coll = db.getCollection("service_line");

			BasicDBObject whereQuery = new BasicDBObject();
			//whereQuery.put("_id",appProperties.getProperty("mongodb.docid"));
			FindIterable<Document> cursor = coll.find(whereQuery);
			ArrayList responseAL = new ArrayList();
			for (Document d : cursor) {
				responseAL.add(d);
			}

			jsonData = mapper.writerWithDefaultPrettyPrinter().writeValueAsString(responseAL);

			mongoClient.close();

		} catch (IOException e) {
			log.error("Sorry, unable to find application.properties");
		}

		HttpHeaders responseHeaders = new HttpHeaders();
		responseHeaders.set("Content-Type", "application/json");

		return ResponseEntity.ok().headers(responseHeaders).body(jsonData);

	}

	
	/*@CrossOrigin(origins = "http://localhost:4200")
	@PostMapping("/update/{serviceLine}")
	public ResponseEntity<String> updateData( @RequestBody NodeData updatedData, @PathVariable  String  serviceLine) throws IOException {
		
		log.info("Initiating the process ");
		InputStream propertiesInput = ExcelToJsonController.class.getClassLoader()
				.getResourceAsStream("application.properties");
		Properties appProperties = getFieldPasroperties();
		ObjectMapper mapper = new ObjectMapper();
		JsonNode rootNode = mapper.createObjectNode();
		try {
			appProperties.load(propertiesInput);

			MongoClient mongoClient = new MongoClient(appProperties.getProperty("mongodb.host"),
					new Integer(appProperties.getProperty("mongodb.port")));
			MongoDatabase db = mongoClient.getDatabase(appProperties.getProperty("mongodb.db"));
			MongoCollection<Document> coll = db.getCollection(appProperties.getProperty("mongodb.collection"));

			BasicDBObject whereQuery = new BasicDBObject();
			whereQuery.put("_id",serviceLine);
			FindIterable<Document> cursor = coll.find(whereQuery);
			//NodeData currentData = new NodeData();
			Map<String,Object> metaObject = new HashMap<String, Object>();
			Map<String,Object> dataObject = new HashMap<String, Object>();
			for (Document d : cursor) {
				//currentData = (NodeData) d.get("data");
				((ObjectNode) rootNode).putPOJO("data", d.get("data"));
				dataObject = (Map<String, Object>) d.get("data");
				metaObject = (Map<String, Object>) d.get("meta");
				//((ObjectNode) rootNode).putPOJO(Constants.META_STR, d.get("meta"));
			}
			
			JSONObject json = new JSONObject(dataObject);
			
			NodeData currentData = mapper.readValue(json.toString(), NodeData.class);
			Properties fieldAppProperties = getFieldPasroperties();
			//NodeData processedData = MessageProcessor.updateData(currentData, fieldAppProperties, updatedData);
			NodeData processedData = MessageProcessor.getUpdatedData(currentData, fieldAppProperties, updatedData);
			try {
				appProperties.load(propertiesInput);

				coll.deleteOne(Filters.eq("_id", serviceLine));

				Document doc = new Document();
				doc.append("_id", serviceLine);
				doc.append("meta", mapper.convertValue(metaObject, Map.class));
				doc.append("data", mapper.convertValue(processedData, Map.class));

				coll.insertOne(doc);

				mongoClient.close();

			} catch (IOException e) {
				e.printStackTrace();
				log.error("Sorry, unable to find application.properties");
			}
			
			mongoClient.close();

		} catch (IOException e) {
			e.printStackTrace();
			log.error("Sorry, unable to find application.properties");
		}
		
		//return "Successfully uploaded file : " ;
		return ResponseEntity.status(HttpStatus.CREATED)
			       .contentType(MediaType.TEXT_PLAIN)
			       .body("Data Updated!");
	}*/
	
	@CrossOrigin(origins = "*")
	@PostMapping("/update/{serviceLine}")
	public String updateData( @RequestBody Map<String, Object> updatedData, @PathVariable  String  serviceLine) throws IOException {
		
		log.info("Initiating the process ");
		InputStream propertiesInput = ExcelToJsonController.class.getClassLoader()
				.getResourceAsStream("application.properties");
		Properties appProperties = new Properties();
		ObjectMapper mapper = new ObjectMapper();
		JsonNode rootNode = mapper.createObjectNode();
		try {
			appProperties.load(propertiesInput);

			MongoClient mongoClient = new MongoClient(appProperties.getProperty("mongodb.host"),
					new Integer(appProperties.getProperty("mongodb.port")));
			MongoDatabase db = mongoClient.getDatabase(appProperties.getProperty("mongodb.db"));
			MongoCollection<Document> coll = db.getCollection(appProperties.getProperty("mongodb.collection"));
			try {
				appProperties.load(propertiesInput);

				coll.deleteOne(Filters.eq("_id", serviceLine));

				Document doc = new Document();
				doc.append("_id", serviceLine);
				doc.append("meta", mapper.convertValue(updatedData.get("meta"), Map.class));
				doc.append("data", mapper.convertValue(updatedData.get("data"), Map.class));

				coll.insertOne(doc);

				mongoClient.close();

			} catch (IOException e) {
				e.printStackTrace();
				log.error("Sorry, unable to find application.properties");
			}
			
			mongoClient.close();

		} catch (IOException e) {
			e.printStackTrace();
			log.error("Sorry, unable to find application.properties");
		}
		
		//return "Successfully uploaded file : " ;

		HttpHeaders responseHeaders = new HttpHeaders();
		responseHeaders.set("Content-Type", "application/json");

		return JSONObject.quote("Data Updated!");
		/*return ResponseEntity.status(HttpStatus.OK)
			       .contentType(MediaType.TEXT_PLAIN)
			       .body("Data Updated!");*/
	}
	
	@CrossOrigin(origins = "*")
	@PostMapping("/upload")
	public String uploadData(@RequestParam("file") MultipartFile file, @RequestParam("description") String description) throws IOException {
		log.info("Initiating the process " + description);
		log.info("File passed : " + file.getOriginalFilename());	
		InputStream is = file.getInputStream();
		String fileName = file.getOriginalFilename();
		String serviceLine = fileName.substring(0, fileName.lastIndexOf("."));
		Properties fieldDetails = getFieldPasroperties();

		boolean isValidInput = isValidInput(file.getOriginalFilename());

		if (!fieldDetails.isEmpty() && isValidInput) {
			ExcelReader excelReader = new ExcelReader();
			ArrayList<Sheet> sheets;
			try {
				sheets = excelReader.readExcel(is);
				if (sheets == null || sheets.size() == 0)
					log.error(Constants.ERR_SHEET_NOT_FOUND);
				else {
					MessageProcessor processor = new MessageProcessor();
					processor.processSheetAndGetData(fieldDetails, sheets, serviceLine, description);
				}
			} catch (InvalidFormatException e) {
				e.printStackTrace();
				return "Failed to upload file";
			}
		}
		return JSONObject.quote("Successfully uploaded file : " + file.getOriginalFilename());
	}

	@CrossOrigin(origins = "http://localhost:4200")
	@GetMapping("/delete")
	public String deleteData() {

		return "No Logic implemented!!";

	}

	@CrossOrigin(origins = "*")
	@GetMapping("/fetch/{serviceLine}")
	public ResponseEntity<String> fetchData(@PathVariable  String  serviceLine) throws JsonProcessingException {
		ObjectMapper mapper = new ObjectMapper();
		JsonNode rootNode = mapper.createObjectNode();
		String jsonData = new String();
		InputStream propertiesInput = ExcelToJsonController.class.getClassLoader()
				.getResourceAsStream("application.properties");
		Properties appProperties = new Properties();

		try {
			appProperties.load(propertiesInput);

			MongoClient mongoClient = new MongoClient(appProperties.getProperty("mongodb.host"),
					new Integer(appProperties.getProperty("mongodb.port")));
			MongoDatabase db = mongoClient.getDatabase(appProperties.getProperty("mongodb.db"));
			MongoCollection<Document> coll = db.getCollection(appProperties.getProperty("mongodb.collection"));

			BasicDBObject whereQuery = new BasicDBObject();
			whereQuery.put("_id",serviceLine);
			FindIterable<Document> cursor = coll.find(whereQuery);

			for (Document d : cursor) {
				((ObjectNode) rootNode).putPOJO("data", d.get("data"));
				((ObjectNode) rootNode).putPOJO(Constants.META_STR, d.get("meta"));
				((ObjectNode) rootNode).putPOJO(Constants.ID, d.get("_id"));
			}

			jsonData = mapper.writerWithDefaultPrettyPrinter().writeValueAsString(rootNode);

			mongoClient.close();

		} catch (IOException e) {
			log.error("Sorry, unable to find application.properties");
		}

		HttpHeaders responseHeaders = new HttpHeaders();
		responseHeaders.set("Content-Type", "application/json");

		return ResponseEntity.ok().headers(responseHeaders).body(jsonData);

	}
	
	@CrossOrigin(origins = "*")
	@GetMapping("/fetch/{serviceLine}/{serialNo}")
	public ResponseEntity<String> fetchDataById(@PathVariable  String  serviceLine, @PathVariable  String  serialNo) throws JsonProcessingException {
		InputStream propertiesInput = ExcelToJsonController.class.getClassLoader()
				.getResourceAsStream("application.properties");
		Properties appProperties = getFieldPasroperties();
		ObjectMapper mapper = new ObjectMapper();
		JsonNode rootNode = mapper.createObjectNode();
		Map a= new HashMap();
		try {
			appProperties.load(propertiesInput);

			MongoClient mongoClient = new MongoClient(appProperties.getProperty("mongodb.host"),
					new Integer(appProperties.getProperty("mongodb.port")));
			MongoDatabase db = mongoClient.getDatabase(appProperties.getProperty("mongodb.db"));
			MongoCollection<Document> coll = db.getCollection(appProperties.getProperty("mongodb.collection"));

			BasicDBObject whereQuery = new BasicDBObject();
			whereQuery.put("_id",serviceLine);
			FindIterable<Document> cursor = coll.find(whereQuery);
			//NodeData currentData = new NodeData();
			Map<String,Object> metaObject = new HashMap<String, Object>();
			Map<String,Object> dataObject = new HashMap<String, Object>();
			for (Document d : cursor) {
				//currentData = (NodeData) d.get("data");
				((ObjectNode) rootNode).putPOJO("data", d.get("data"));
				dataObject = (Map<String, Object>) d.get("data");
				metaObject = (Map<String, Object>) d.get("meta");
				//((ObjectNode) rootNode).putPOJO(Constants.META_STR, d.get("meta"));
			}
			mongoClient.close();
			JSONObject json = new JSONObject(dataObject);
			
			NodeData currentData = mapper.readValue(json.toString(), NodeData.class);
			Properties fieldAppProperties = getFieldPasroperties();
			//NodeData processedData = MessageProcessor.updateData(currentData, fieldAppProperties, updatedData);
			NodeData processedData = MessageProcessor.getDataBySerialNo(currentData, fieldAppProperties, serialNo);
			a.put("_id", serviceLine);
			a.put("meta", mapper.convertValue(metaObject, Map.class));
			a.put("data", mapper.convertValue(processedData, Map.class));
		} catch (IOException e) {
			e.printStackTrace();
			log.error("Sorry, unable to find application.properties");
		}
		String jsonData = mapper.writerWithDefaultPrettyPrinter().writeValueAsString(a);
		HttpHeaders responseHeaders = new HttpHeaders();
		responseHeaders.set("Content-Type", "application/json");

		return ResponseEntity.ok().headers(responseHeaders).body(jsonData);

	}


	public static Properties getFieldPasroperties() {
		Properties properties = new Properties();
		try (InputStream input = ExcelToJsonController.class.getClassLoader()
				.getResourceAsStream("field-details.properties")) {

			if (input == null) {
				log.error("Sorry, unable to find field-details.properties");
				return null;
			}

			properties.load(input);

		} catch (IOException ex) {
			ex.printStackTrace();
		}

		return properties;
	}

	public static boolean isValidInput(String args) {
		if (args.length() == 0 || args == null || args.equals(Constants.EMPTY_STRING)) {
			log.error(Constants.ERR_INVALID_ARG);
			return false;
		}
		if (args.contains(".xlsm") || args.contains(".xlsx")) {
			return true;
		} else {
			log.error(Constants.ERR_FILE_FORMAT_NOT_FOUND);
			return false;
		}
	}

	public static boolean createJSONFile(String processedData) {
		try {

			FileOutputStream fos = new FileOutputStream(
					Constants.CURRENT_FOLDER_PATH + Constants.PATH_SLASH + Constants.GENERATING_FILE_NAME);
			fos.write(processedData.toString().getBytes());
			fos.flush();
			fos.close();

			log.info("JSON File generated in location " + Constants.CURRENT_FOLDER_PATH + Constants.PATH_SLASH
					+ Constants.GENERATING_FILE_NAME);
			return true;
		} catch (IOException e) {
			log.error("Error in creating the file ");
			return false;
		}
	}

}
