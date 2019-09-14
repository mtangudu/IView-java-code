package com.deloitte.exceltojson.repo;

import java.util.List;

import org.springframework.data.mongodb.repository.MongoRepository;
import org.springframework.data.mongodb.repository.Query;
import org.springframework.stereotype.Component;
import org.springframework.stereotype.Repository;

import com.deloitte.exceltojson.pojo.SheetData;

@Component
public interface NodeDataRepository extends MongoRepository<SheetData, String> {
	
	

}
