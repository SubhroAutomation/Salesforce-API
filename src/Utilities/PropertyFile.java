package Utilities;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;

public class PropertyFile {
	SubhroAutomation commit-2
	Subhro80 - commit - 5
	Dubhro80 - Commit6
	private static FileInputStream fStream;
	private static Properties propertyFile = new Properties();
	
	public static boolean LoadPropertyFile (String strPropertyFileName) throws IOException
	{
		boolean bStatus = false;
		try{
			fStream = new FileInputStream(strPropertyFileName);
			propertyFile.load(fStream);
			bStatus = true;
		}
		catch (Exception e){
			bStatus = false;
			e.printStackTrace();
			System.out.println("Property File not present in folder " +strPropertyFileName);
		}
		return bStatus;
	}
	
	public static String getKeyValue (String strLocatorName) throws Exception{
		String strKey = "";
		String strValue = null;
		try{
			strValue = propertyFile.getProperty(strLocatorName);
		}
		catch(Exception e){
			e.printStackTrace();
			strValue = null;
			System.out.println("Error in retrieving value from Property File");
		}
		return strValue;
	}
}
