package GUI;

import static Utilities.ExcelPOI.strTestDataFilePath;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.sforce.soap.enterprise.EnterpriseConnection;
import com.sforce.soap.enterprise.LoginResult;
import com.sforce.soap.metadata.AsyncResult;
import com.sforce.soap.metadata.DescribeMetadataObject;
import com.sforce.soap.metadata.DescribeMetadataResult;
import com.sforce.soap.metadata.Metadata;
import com.sforce.soap.metadata.MetadataConnection;
import com.sforce.soap.metadata.Profile;
import com.sforce.soap.metadata.ProfileFieldLevelSecurity;
import com.sforce.soap.metadata.SaveResult;
import com.sforce.soap.metadata.UpdateMetadata_element;
import com.sforce.ws.ConnectionException;
import com.sforce.ws.ConnectorConfig;
import com.sun.xml.internal.ws.util.MetadataUtil;

import Components.MetaDataAPI;
import Utilities.ExcelPOI;
import Utilities.FileOperations;
import Utilities.PropertyFile;

public class PostDepTool {
	public static Map<String, ArrayList<String>> hmapProfileFieldObjAccess;
	public static Map<String, ArrayList<String>> hmapProfileVFPageAccess;
	public static Map<String, ArrayList<String>> hmapProfileApexClassAccess;
	public static Map<String, ArrayList<String>> hmapProfileObjectAccessRecTypAssDefault;
	public static Map<String, ArrayList<String>> hmapProfileObjectRecTypePageLayout;
	public static Map<String, ArrayList<String>> hmapPermissionSetFieldObjAccess;
	
	public static ArrayList<String> alFLSProfilesfrmRTP = new ArrayList <String> ();
	public static ArrayList<String> alVFPageProfilesfrmRTP = new ArrayList <String> ();
	public static ArrayList<String> alApexClassProfilesfrmRTP = new ArrayList <String> ();
	public static ArrayList<String> alObjectProfilesfrmRTP = new ArrayList <String> ();
	public static ArrayList<String> alPageLayoutProfilesfrmRTP = new ArrayList <String> ();
	public static ArrayList<String> alPermissionSetsfrmRTP = new ArrayList <String> ();
	
	public static ArrayList<String> alALLProfilesinSFDC = new ArrayList <String> ();
	public static ArrayList<String> alALLPermissionSetsinSFDC = new ArrayList <String> ();
	
	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		
		String strPropertyFileName = System.getProperty("user.dir") + "\\Config.txt"; 
		String strRTPFilePath = null;
		if ( false != PropertyFile.LoadPropertyFile(strPropertyFileName)){
			strRTPFilePath = PropertyFile.getKeyValue("RTPFilePath");
			if ( strRTPFilePath != null ){
				System.out.println("RTP File path: = " +strRTPFilePath);
				ExcelPOI.strTestDataFilePath = strRTPFilePath;				
			}
		}
		if (strRTPFilePath == null)
			ExcelPOI.browseForExcelFile(); 
				
		//ExcelPOI.strTestDataFilePath = "C:\\Users\\267567\\Desktop\\POST DEP\\Salesforce Release 02_16jan.xlsx";  //"C:\\Users\\267567\\Desktop\\POST DEP\\Salesforce Release 46_14thNov.xlsx";
		
		MetaDataAPI.login();
		MetaDataAPI.loginUsingPartnerConnection();
		
		MetaDataAPI.getListofAllProfilesinSFDC ();
		MetaDataAPI.getListofAllPermissionSetsinSFDC ();
		
		//FLS.setFLS();
		//VFPageAccess.setVFPageAccess();
		//ApexClassAccess.setApexClassAccess();
		ObjectAccess.setObjectAccess();
		//PageLayoutAssignment.assignPageLayout();	
		//PermissionSet.setPermissionSet();
	}
    
}
