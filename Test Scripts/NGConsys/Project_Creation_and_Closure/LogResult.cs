/*
 * Created by Ranorex
 * User: jbhosash
 * Date: 12/21/2017
 * Time: 11:09 AM
 * 
 * To change this template use Tools > Options > Coding > Edit standard headers.
 */
using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.IO.MemoryMappedFiles;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Threading;
using WinForms = System.Windows.Forms;
using System.Runtime.InteropServices; 
using Excel = Microsoft.Office.Interop.Excel;

using System.Xml;

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;
using TestProject.Libraries;


namespace TestProject
{
    /// <summary>
    /// Description of LogResult.
    /// </summary>
    [TestModule("CEB1BBA3-253D-4729-A317-9E2EF51370BF", ModuleType.UserCode, 1)]
    public class LogResult : ITestModule
    {
        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public LogResult()
        {
            // Do not delete - a parameterless constructor is required!
        }

    	int _rowNumber = 3;
		[TestVariable("e174da3c-6113-4c84-97af-7523f0d60fe9")]
		public int rowNumber
		{
			get { return _rowNumber; }
			set { _rowNumber = value; }
		}

       string _ReportFile = "";
       [TestVariable("2c8624bb-58d9-4427-8013-f758fb4b4217")]
       public string ReportFile
       {
	       get { return _ReportFile; }
	       set { _ReportFile = value; }
       }

        /// <summary>
        /// Performs the playback of actions in this module.
        /// </summary>
        /// <remarks>You should not call this method directly, instead pass the module
        /// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
        /// that will in turn invoke this method.</remarks>
        void ITestModule.Run()
        {
            Mouse.DefaultMoveTime = 300;
            Keyboard.DefaultKeyPressTime = 100;
            Delay.SpeedFactor = 1.0;
            
            //Get current test case name    	  			 
        	string tcName = TestSuite.CurrentTestContainer.Name; 
        	        	     	
        	//Get execution status of test case
        	string status = TestSuite.CurrentTestContainer.Status.ToString();
        	
        	//Get test suite report link         
      		string reportLink = Ranorex.Core.Reporting.TestReport.ReportEnvironment.ReportViewFilePath;
      		         	        	    	
            /* Code related to notepad
        	 
  			String header = "TC Name" + "                          " + "Status";
        	String result = tcName + "                          " +status;
        	
			String path ="C:\\Users\\jbhosash\\Desktop\\test.txt";
	     	System.IO.File.WriteAllText(path,header);
		  	System.IO.File.AppendAllText(path,Environment.NewLine + result);
        	*/
		
        	// Create objects of excel 
		    Excel.Workbook excelWB;
			Excel.Sheets excelSheets;
			Excel.Worksheet excelWorksheet;
		    Excel.Range excelRange;
			
			Excel.Application excelFile = new Excel.ApplicationClass();
			
			//Create file name if empty
			if(ReportFile=="")
			{
									
				string actualDirPath = Common_Functions.GetDirPath();
				string filePath = actualDirPath +"Test Result"+"\\";
		    	string fileName = "Result_"+ System.DateTime.Now.ToString("MM_dd_yyyy_hh_mm_ss") +".xlsx";
		    	ReportFile =filePath+fileName;
			}
			
			//Check if Result file exist already
			if (!System.IO.File.Exists(ReportFile))
			{
				//Create new file
   			 	excelWB =excelFile.Workbooks.Add(1);
   			 	excelSheets = excelWB.Worksheets;
            	excelWorksheet = (Excel.Worksheet)excelSheets.get_Item("Sheet1");            	          	
            	
            	// Generate Headers
            	excelFile.Cells[1,1] = "Report Link";
            	excelFile.Cells[1,2] = reportLink;
            	excelFile.Cells[2,1] = "Test Case Name";  
				excelFile.Cells[2,2] = "Status"; 
				 
			}
			
			else
			{
			    //Open existing file
				excelWB = excelFile.Workbooks.Open(ReportFile);
				excelSheets = excelWB.Worksheets;
            	excelWorksheet = (Excel.Worksheet)excelSheets.get_Item("Sheet1");
			}		

				//Cell formatting
            	excelRange = excelWorksheet.get_Range("A1");
    			excelRange.Font.Bold = true;
    			excelRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSkyBlue);
    			
    			excelRange = excelWorksheet.get_Range("A2");
    			excelRange.Font.Bold=true;	
    			excelRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightPink);
				 
				excelRange = excelWorksheet.get_Range("B2");
    			excelRange.Font.Bold=true;
    			excelRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightPink);
    			
				excelRange = excelWorksheet.get_Range("B"+rowNumber);
								
				if(status=="Success")
				{
					status="Pass";
					excelRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
				}
				else
				{
					status="Fail";
					excelRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
				}
				
				excelRange=excelWorksheet.get_Range("B1");
				excelWorksheet.Hyperlinks.Add(excelRange,reportLink,Type.Missing,"Ranorex Report","Ranorex Report");
				
				
				//Add test case name and status to excel file
				excelFile.Cells[rowNumber,1] = tcName;
				excelFile.Cells[rowNumber,2] = status;
			
				rowNumber = rowNumber+1;
				
				//Add border to all used cells
				excelRange = excelWorksheet.UsedRange;
				excelRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
				excelRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

				// save excel file 
				excelFile.DisplayAlerts = false;
				excelWB.SaveAs(ReportFile);
			     
				// close excel file  
				excelWB.Close();  
			
			 }
             
         

       
    }
}
