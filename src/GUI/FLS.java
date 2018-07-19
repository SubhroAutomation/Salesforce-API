package GUI;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

import Components.MetaDataAPI;
import Components.ReadFLSfromRTPExcel;
import Utilities.ExcelPOI;
import Utilities.FileOperations;

import static GUI.PostDepTool.hmapProfileFieldObjAccess;
import static GUI.PostDepTool.alFLSProfilesfrmRTP;

public class FLS {
	public static void setFLS() throws Exception{
		FileOperations.createFile("FLS");	
		
		int iStartRow = -1;
		iStartRow = ReadFLSfromRTPExcel.getStartRowIndexAfterAllProfTaginRTP();
		
		alFLSProfilesfrmRTP = ReadFLSfromRTPExcel.getListofFLSProfilesfrmRTP (iStartRow);
		ReadFLSfromRTPExcel.getCombinationofAllFieldObjAccessProfinRTP (iStartRow);
		
		int iStartRowAllProf = -1;
		iStartRowAllProf = ExcelPOI.GetRowIndexofValueinCol("FLS", "ALL Profiles-Start", 0, 0) + 1;
		ReadFLSfromRTPExcel.getCombinationofFieldObjAccfrmALLPROFILES (iStartRowAllProf);
		
		//RTPExcel.createProfileFieldObjAccessSheet ();
		
		MetaDataAPI.createProfileFieldLvlSecurityArray();		
	}
}
