package code.selenium.controllers;

import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;

import code.selenium.entities.Location;
import code.selenium.services.LocationServiceImpl;

@RestController
public class LocationController {

	@Autowired
	private LocationServiceImpl locationService;
	
	@GetMapping("/locations")
	public List<Location> getLocations() throws InterruptedException {
		return this.locationService.getAllLocations();
	}
	
	@GetMapping("/store")
	public String storeData() {
		return this.locationService.writeDataInExcel();
	}
}
