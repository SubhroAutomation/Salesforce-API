package Utilities;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

public class FileOperations {
	SubhroAutomation commit-1
	SubhroAutomation commit-3
	SubhroAutomation commit-7
	
	public static File file;
	
	public static void createFile(String strMetadata) throws IOException{
    	DateFormat dateFormat = new SimpleDateFormat("MM/dd_HH:mm:ss");
	    Date date = new Date();
	    
	    System.getProperty("user.dir");
	    
    	String strFileName = strMetadata + "_Log_" + dateFormat.format(date) + ".txt";
    	strFileName = strFileName.replace(" ", "");
    	strFileName = strFileName.replace("/", "-");
    	strFileName = strFileName.replace(":", "-");
    	
    	strFileName = System.getProperty("user.dir") + "\\" + strFileName; 
    	System.out.println("Log File: " +strFileName);
    	file = new File(strFileName);
        
        // creates the file
        file.createNewFile();
        System.out.println("Log File: " +strFileName);
    }
    
    public static void writeToLog(String strContentToFile, boolean... bNewLine) throws IOException{
        // creates a FileWriter Object
        FileWriter writer = new FileWriter(file, true); //  (file, "true"); 
        
        // Writes the content to the file
        writer.append(strContentToFile);
        
        //If optional parameter not included, then by default new line is added.
        //If optional parameter included, then finding out if newline optional parameter argument is true or false
        boolean bAddNewLine = true;
        if (bNewLine.length > 0){
        	bAddNewLine = bNewLine[0];
        }
        
        if(bAddNewLine == true){
        	writer.append(System.lineSeparator());
        }
        else{
        	//Don't Add a new line
        }
        
        writer.flush();
        writer.close();
    }
}
