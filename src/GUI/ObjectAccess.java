package GUI;

import static GUI.PostDepTool.alObjectProfilesfrmRTP;
import static GUI.PostDepTool.hmapProfileObjectAccessRecTypAssDefault;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

import Components.MetaDataAPI;
import Components.ReadObjectfromRTPExcel;
import Utilities.ExcelPOI;
import Utilities.FileOperations;

public class ObjectAccess {
	public static void setObjectAccess() throws Exception{
		FileOperations.createFile("Object");	
		
		int iStartRow = -1;
		iStartRow = ReadObjectfromRTPExcel.getStartRowIndexAfterAllProfTaginRTP_Object();
		
		alObjectProfilesfrmRTP = ReadObjectfromRTPExcel.getListofObjectProfilesfrmRTP (iStartRow);
		ReadObjectfromRTPExcel.getCombinationofAllObjectProfinRTP(iStartRow);
		
		int iStartRowAllProf = -1;
		int iStartRowofObjectAccess = ExcelPOI.GetRowIndexofValueinCol("VFPage,Class,Obj,PageLayout", "Object Access:", 0, 0);
		iStartRowAllProf = ExcelPOI.GetRowIndexofValueinCol("VFPage,Class,Obj,PageLayout", "All Profiles-Start", 0, iStartRowofObjectAccess) + 1;
		ReadObjectfromRTPExcel.getCombinationofObjectfrmALLPROFILES (iStartRowAllProf);
		
		//RTPExcel.createProfileFieldObjAccessSheet ();
		
		MetaDataAPI.createProfileObjPermAndProfileRecVisibilityArray();
	}
}
