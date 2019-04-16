/*
 * Created by Ranorex
 * User: jbhosash
 * Date: 8/30/2018
 * Time: 11:51 AM
 * 
 * To change this template use Tools > Options > Coding > Edit standard headers.
 */
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Threading;
using WinForms = System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace TestProject.Libraries
{
    /// <summary>
    /// Ranorex user code collection. A collection is used to publish user code methods to the user code library.
    /// </summary>
    [UserCodeCollection]
    public class Excel_Utilities
    {
    	
    	public static Excel.Workbook ExcelWB{get; set;}
    	public static Excel.Sheets ExcelSheets{get;set;}
    	public static Excel.Worksheet ExcelWorksheet{get;set;}
    	public static Excel.Range ExcelRange{get;set;}
    	
    	public static Excel.Application ExcelAppl
    	{
    		get
    		{
    			if(_application==null)
    			{
    				_application = new Excel.ApplicationClass();
    			}
    			
    			return _application;
    		}
    	}
    	public static Excel.Application _application;
			
       /********************************************************************
		 * Function Name: OpenExcelFile
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :
		 ********************************************************************/
       [UserCodeMethod]
        public static void OpenExcelFile(string sFileName, string sSheetName)
      	{
        	string actualDirPath = Common_Functions.GetDirPath();
			string filePath = actualDirPath +"Test Data"+"\\";
		    string DataFile =filePath+sFileName;
		    
		    ExcelWB = ExcelAppl.Workbooks.Open(DataFile);
   			ExcelSheets = ExcelWB.Worksheets;
   			ExcelWorksheet = (Excel.Worksheet)ExcelSheets.get_Item(sSheetName);
   			ExcelRange=ExcelWorksheet.UsedRange;		
   		}
        
        /********************************************************************
		 * Function Name: CloseExcel
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :
		 ********************************************************************/
         public static void CloseExcel()
        {
        	if(ExcelWB!=null)
			{
				ExcelWB.Close(false, null, null);
			}
        }
         
//        [UserCodeMethod]
//        public static void OpenSheet(string sSheetName)
//        {
//        	
//   			
//   			
//        }
//        
//        public static bool IsFileOpen(string filePath)
//		{
// 		 bool rtnvalue = false;
//  			try
//  			{
//     			System.IO.FileStream fs = System.IO.File.OpenWrite(filePath);
//    			fs.Close();
//  			}
// 	  		catch(System.IO.IOException ex)
//  			{
//    			rtnvalue = true;
//  			}	
//    	return rtnvalue;
//		}
//        
}
}