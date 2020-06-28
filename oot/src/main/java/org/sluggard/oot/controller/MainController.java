package org.sluggard.oot.controller;

import org.sluggard.oot.dao.SimpleDao;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
public class MainController {
	
	@Autowired
	private SimpleDao simpleDao;
	
	@RequestMapping("main")
	public String main() {
		simpleDao.runSimpleSql();
		return "a";
	}
	
	@RequestMapping("test")
	public String test(@RequestBody String data) {
		System.out.println(data);
		return data;
	}

}
