package com.report;
import java.io.FileInputStream;
import java.util.Properties;


public class Auth {public static Properties prop;
public String username;
public String password;

public Auth() 
{
	try {
	prop = new Properties();
	String userDirectory = System.getProperty("user.dir");
	FileInputStream fis = new FileInputStream(userDirectory + "/Properties");
	prop.load(fis);
	} catch (Exception e) {
		e.printStackTrace();
	}
}


public String getusername() {
	username = prop.getProperty("username");
	return username;
}

public void setusername(String username) {
	this.username = username;
}
public String getpassword()  {
	password = prop.getProperty("password");
	return password;
}

public void setpassword(String password) {
	this.password = password;
}

}
