package Test;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import javax.swing.JOptionPane;

import com.sforce.soap.metadata.Metadata;
import com.sforce.soap.metadata.Profile;
import com.sforce.soap.metadata.ProfileObjectPermissions;
import com.sforce.soap.metadata.ProfileRecordTypeVisibility;
import com.sforce.soap.metadata.SaveResult;
import com.sforce.soap.partner.QueryResult;
import com.sforce.ws.ConnectionException;

import Utilities.FileOperations;

public class Test {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		
		String str = "Blank";
		//String str2 = "Blank";
		
		if (!str.equals("Default") && !str.equals("Blank"))
			System.out.println("1");
		else
			System.out.println("2");
		
		/*String[]ar = new String[5];
		ar[0] = "hello ";
		ar[1] = "hello1  ";
		ar[2] = "hello2 ";
		ar[3] = " hello3";
		ar[4] = "hello";
		
		//ArrayList<String> arlistRecTypeAss = (ArrayList<String>) Arrays.asList(ar);
		//arlistRecTypeAss.replaceAll(String::trim);*/
		
		/*String strURL = "https://sunpower--uat.cs61.my.salesforce.com/0014C000009ASdxQAG";
		System.out.println("Acc URL for Comm: " +strURL);
		String strToReplace = strURL.substring(strURL.lastIndexOf('/') + 1);
		System.out.println("Acc URL for Comm: " +strToReplace);
		strURL = strURL.replace(strToReplace, "AccountNewDetail?id=" + "0014C000009ASdxQAG");
		System.out.println("Acc URL for Comm: " +strURL);
		strURL = strURL.replace("SPCommunityAccountDetails?id=","");
		System.out.println("Acc URL for Comm: " +strURL);*/
		
		
		/*String a = "123#456#-#789#-#";
		//a = a.trim();
		System.out.println(a.trim());
		
		String[] al = new String[100];
		al = a.split("#");
		System.out.println(al.length);
		for(String str : al){
			//System.out.println(str);
		}
		
		List <String> alRecTyp = Arrays.asList(al);
		System.out.println(alRecTyp.get(1));*/
		
		/*Map<String, ArrayList<String>> hmapProfileFieldObjAccess = new LinkedHashMap<> ();
		
		ArrayList <String> temp = new ArrayList();
		temp.add("Hello");
		
		ArrayList <String> alTemp = new ArrayList <String> ();
		alTemp.add("Test1");
		alTemp.add("Test2");
		
		hmapProfileFieldObjAccess.put("One", alTemp);
		hmapProfileFieldObjAccess.put("Two", null);
		temp.add("Fellow");
		hmapProfileFieldObjAccess.put("One", temp);
		
		Set set = hmapProfileFieldObjAccess.entrySet();
		Iterator iterator = set.iterator();
		  while(iterator.hasNext()) {
		     Map.Entry mentry = (Map.Entry)iterator.next();
		     System.out.print("key is: "+ mentry.getKey() + " & Value is: ");
		     System.out.println(mentry.getValue());
		  }*/
	}

}
