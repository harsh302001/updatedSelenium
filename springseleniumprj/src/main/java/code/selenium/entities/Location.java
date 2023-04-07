package code.selenium.entities;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class Location {

	String city;
	String address;
	String phNo;
	String LocType;
	String LocHrs;
}

