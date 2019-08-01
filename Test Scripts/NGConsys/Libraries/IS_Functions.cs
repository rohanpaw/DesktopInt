/*
 * Created by Ranorex
 * User: jdhakaa
 * Date: 12/4/2018
 * Time: 12:03 PM
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
using Microsoft.Office.Interop.Excel;


using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace TestProject.Libraries
{
	
	[UserCodeCollection]
	public class IS_Functions
	{
		//Create instance of repository to access repository items
		static NGConsysRepository repo = NGConsysRepository.Instance;
		
		static string ModelNumber
		{
			
			get { return repo.ModelNumber; }
			set { repo.ModelNumber = value; }
		}
		
		static string sISUnits
		{
			get { return repo.sISUnits; }
			set { repo.sISUnits = value; }
		}
		
		static string sRow
		{
			get { return repo.sRow; }
			set { repo.sRow = value; }
		}

		static string sLabelName
		{
			
			get { return repo.sLabelName; }
			set { repo.sLabelName = value; }
		}
		
		static string sBase
		{
			get { return repo.sBase; }
			set { repo.sBase = value; }
		}
		
		static string sRowIndex
		{
			get { return repo.sRowIndex; }
			set { repo.sRowIndex = value; }
		}

		static string sExtraISUnitsReq
		{
			get { return repo.sExtraISUnitsReq; }
			set { repo.sExtraISUnitsReq = value; }
		}
		
		static string sMaxExtraISUnits
		{
			get { return repo.sMaxExtraISUnits; }
			set { repo.sMaxExtraISUnits = value; }
		}
		
		static string sCell
		{
			get { return repo.sCell; }
			set { repo.sCell = value; }
		}
		
		static string sDeviceOrderName
		{
			get { return repo.sDeviceOrderName; }
			set { repo.sDeviceOrderName = value; }
		}
		
		static string sDeviceOrderRow
		{
			get { return repo.sDeviceOrderRow; }
			set { repo.sDeviceOrderRow = value; }
		}
		
		static string sPhysicalLayoutDeviceIndex
		{
			get { return repo.sPhysicalLayoutDeviceIndex; }
			set { repo.sPhysicalLayoutDeviceIndex = value; }
		}
		
		/********************************************************************
		 * Function Name: VerifyISCalculation
		 * Function Details: To verify IS calculation of IS devices
		 * Parameter/Arguments:  fileName, sheetName
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 13/12/2018
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyISCalculation(string sFileName,string sAddDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared string type
			string sType, sLabelName,ExpectedMaxISUnits,ChangedValue,PanelType;
			int noOfDevices = 0;
			PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[4,7]).Value.ToString();
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				noOfDevices++;
				ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				ExpectedMaxISUnits = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				sISUnits = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				ChangedValue =  ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				
				// Verify panel type and then accordingly assign sRow value
				if(PanelType.Equals("FIM"))
				{
					sRow = (i+1).ToString();
				}
				else
				{
					sRow = (i+2).ToString();
				}
				
				// Add devices from the gallery as per test data from the excel sheet
				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
				Report.Log(ReportLevel.Info, "Device "+ModelNumber+" added to Panel");
				
				//Click on Physical layout tab
				repo.ProfileConsys1.tab_PhysicalLayout.Click();
				
				//To call verifyMaxISUnits method and verify max IS units value
				Report.Log(ReportLevel.Info,"Verification of Maximum IS Units on addition of devices");
				verifyMaxISUnitsMultipleDevices(ExpectedMaxISUnits, noOfDevices, PanelType);
				
				Report.Log(ReportLevel.Info, "Verified Maximum IS Units.");
				
				//Click on Points tab
				repo.ProfileConsys1.tab_Points.Click();
				
				// Double click on cell capacitance cell
				repo.ProfileConsys1.cell_CableCapacitance.DoubleClick();
				
				// Change the cable capacitance value
				repo.ProfileConsys1.txt_CableCapacitance.PressKeys((ChangedValue) +"{ENTER}");

			}
			//Close opened excel sheet
			Excel_Utilities.CloseExcel();
			
		}


		/****************************************************************************************************************
		 * Function Name: verifyISUnits
		 * Function Details: Set 2nd Argument sRow 9 for FIM and 10 for PFI (When called directly in recording and only EXI is present,
		 * 			         but if any IS devices added then this row number will change accordingly)
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 17/12/2018  25/01/2019 Alpesh Dhakad  - Updated Xpath ISUnits
		 * 				 04/02/2019 Alpesh Dhakad - Added Argument Row and updated wherever the function used. Also, added sRow=Row line
		 *
		 ****************************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyISUnits(string expectedISunit, string Row)
		{
			// Set sRow value
			sRow=Row;
			
			//Click on Physical layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			// Store IS unit value in ActualMaxISUnit variable
			string ActualISUnit = repo.FormMe.ISUnits.TextValue;
			
			// Compare ActualISUnit and expectedISUnit values and then displaying result
			if(ActualISUnit.Equals(expectedISunit))
			{
				Report.Log(ReportLevel.Success,"IS Unit " + ActualISUnit + " is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"IS Unit is not displayed correctly, IS Units displayed as: " +ActualISUnit + " instead of : "+expectedISunit);
			}
		}
		
		
		/*********************************************************************************************************************************
		 * Function Name: verifyMaxISUnits
		 * Function Details: To verify maximum IS units value and Set 2nd Argument sRow 9 for FIM and 10 for PFI (When called directly in recording and only EXI is present,
		 * 			         but if any IS devices added then this row number will change accordingly)
		 * Parameter/Arguments: expectedMax IS unit value
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 13/12/2018  04/02/2019 Alpesh Dhakad - Added Argument Row and updated wherever the function used. Also, added sRow=Row line
		 **********************************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyMaxISUnits(string expectedMaxISUnit,string Row)
		{
			// Set sRow value
			sRow=Row;
			
			string actualMaxISUnits = repo.FormMe.MaxISUnits.TextValue;
			
			// Compare ActualMaxISUnit and expectedMaxISUnit values and then displaying result
			if(actualMaxISUnits.Equals(expectedMaxISUnit))
			{
				Report.Log(ReportLevel.Success,"Max IS Unit " + actualMaxISUnits + " is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Max IS Unit is not displayed correctly, Max IS Units displayed as: " +actualMaxISUnits + " instead of : "+expectedMaxISUnit);
			}
		}
		
		/********************************************************************
		 * Function Name: verifyMaxISUnitsMultipleDevices
		 * Function Details: To verify maximum IS units value for multiple devices.
		 * Parameter/Arguments: expectedMax IS unit value, number of devices
		 * Output:
		 * Function Owner: Devendra Kulkarni
		 * Last Update : 13/12/2018
		 ********************************************************************/
		[UserCodeMethod]
		public static void verifyMaxISUnitsMultipleDevices(string expectedMaxISUnit, int noOfDevices, string panelType)
		{
			int rowNo;

			// Determine the row number of first panel
			if(panelType.Equals("FIM"))
			{
				rowNo = noOfDevices + 8;
			}
			else
			{
				rowNo = noOfDevices + 9;
			}
			
			Report.Log(ReportLevel.Info, "Verifying Max IS units for "+noOfDevices+ " devices.");
			for(int j=rowNo; j<(rowNo+noOfDevices);j++)
			{
				sRow = j.ToString();
				
				verifyMaxISUnits(expectedMaxISUnit,sRow);
			}

		}
		
		/********************************************************************
		 * Function Name: VerifyISCalculationForDifferentCableCapacitance
		 * Function Details: To verify IS calculation on updating Cable capacitance
		 * Parameter/Arguments:  fileName, sheetName
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 06/12/2018
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyISCalculationForDifferentCableCapacitance(string sFileName,string sAddDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared string type
			string sType, sLabelName,ExpectedMaxISUnits,ChangedValue,PanelType;
			
			ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[8,1]).Value.ToString();
			sType = ((Range)Excel_Utilities.ExcelRange.Cells[8,2]).Value.ToString();
			
			// Add devices from the gallery as per test data from the excel sheet
			Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
			Report.Log(ReportLevel.Info, "Device "+ModelNumber+" added to Panel");
			
			
			PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[4,7]).Value.ToString();
			
			// Setting the row count j to 8 to assign the sRow value as per panel type
			int j=8;
			
			// Verify panel type and then accordingly assign sRow value
			if(PanelType.Equals("FIM"))
			{
				sRow = (j+1).ToString();
			}
			else
			{
				sRow = (j+2).ToString();
			}
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				ExpectedMaxISUnits = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				sISUnits= ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				ChangedValue =  ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				
				//Click on Physical layout tab
				repo.ProfileConsys1.tab_PhysicalLayout.Click();
				
				//To call verifyISUnits method and verify IS units value
				Report.Log(ReportLevel.Info,"Verification of IS Units on addition of devices");
				verifyISUnits(sISUnits,sRow);
				
				//Click on Points tab
				repo.ProfileConsys1.tab_Points.Click();
				
				// Double click on cell capacitance cell
				repo.ProfileConsys1.cell_CableCapacitance.DoubleClick();
				
				// Change the cable capacitance value
				repo.ProfileConsys1.txt_CableCapacitance.PressKeys((ChangedValue) +"{ENTER}");
				Report.Log(ReportLevel.Info, "Verified IS Units.");
			}
			//Close opened excel sheet
			Excel_Utilities.CloseExcel();
		}
		
		/********************************************************************
		 * Function Name: VerifyISCalculationOnAddDeleteISDevice
		 * Function Details: Verify the IS unit calculation when we add/delete IS devices from EXI800
		 * Parameter/Arguments: sFileName, sAddDevicesSheet, sAddIsDevicesSheet
		 * Output:
		 * Function Owner: Poonam Kadam
		 * Last Update : 7/12/2018
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyISCalculationOnAddDeleteISDevice(string sFileName,string sAddEXIDevicesSheet, string sAddIsDevicesSheet)
		{
			//Add devies in loop A,
			Excel_Utilities.OpenExcelFile(sFileName,sAddEXIDevicesSheet);
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			string sType,sExpectedMaxISUnits,PanelType;
			for(int i=8; i<=rows; i++)
			{
				
				ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sExpectedMaxISUnits = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				sISUnits = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[4,7]).Value.ToString();
				
				if(PanelType.Equals("FIM"))
				{
					sRow = (i+1).ToString();
				}
				else
				{
					sRow = (i+2).ToString();
				}
				
				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
				
				repo.ProfileConsys1.tab_PhysicalLayout.Click();
				//Verify IS units in physical layout
				verifyISUnits(sISUnits,sRow);
				
				//Verify Max IS units in physical layout
				verifyMaxISUnits(sExpectedMaxISUnits,sRow);
				
				repo.ProfileConsys1.tab_Points.Click();
				
			}
			//Close excel
			Excel_Utilities.CloseExcel();
			
			//Open excel
			Excel_Utilities.OpenExcelFile(sFileName,sAddIsDevicesSheet);
			rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			//Add IS devices to EXI800
			for(int i=8; i<=rows; i++)
			{
				ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				
				
				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
				Report.Log(ReportLevel.Info,"Device"+ModelNumber+" Added to loop");
				sLabelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,10]).Value.ToString();
				
				Report.Log(ReportLevel.Success,"Label is "+ sLabelName);
				
				repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1.Click();
			}
			
			//Go to physical layout
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			// Setting the row count k to 8 to set sRow value according to Panel type and k is used in expected MaxIS unit
			int k=8;
			
			PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[4,7]).Value.ToString();
			
			if(PanelType.Equals("FIM"))
			{
				sRow = (k+2).ToString();
			}
			else
			{
				sRow = (k+3).ToString();
			}
			
			
			sISUnits = ((Range)Excel_Utilities.ExcelRange.Cells[9,8]).Value.ToString();
			//Verify IS Units
			verifyISUnits(sISUnits,sRow);
			sExpectedMaxISUnits = ((Range)Excel_Utilities.ExcelRange.Cells[k,7]).Value.ToString();
			
			//Verify MAX IS units
			verifyMaxISUnits(sExpectedMaxISUnits,sRow);
			
			// Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			//Setting the row count j to 8 to set sRow value and fetch values from excel
			int j=8;
			sExpectedMaxISUnits = ((Range)Excel_Utilities.ExcelRange.Cells[j,7]).Value.ToString();
			sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[j,9]).Value.ToString();
			
			//Verify IS Unit after deleting IS device
			sISUnits = ((Range)Excel_Utilities.ExcelRange.Cells[j+1,8]).Value.ToString();
			
			//Delete IS device
			Devices_Functions.DeleteDeviceUsingLabel(sLabelName);
			Report.Log(ReportLevel.Info,"Device "+sLabelName+" deleted");
			
			// Click on Physical Layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			//Verify IS Units
			verifyISUnits(sISUnits,sRow);
			
			//Verify MaxIS Units
			verifyMaxISUnits(sExpectedMaxISUnits,sRow);
		}
		
		
		/********************************************************************
		 * Function Name: VerifyExtraISUnitsForPFI_FIMLoops
		 * Function Details: Verify Extra IS Units Required in case of split loops/FIM loops
		 * Parameter/Arguments: sFileName, sAddDevicesSheet, sAddIsDevicesSheet
		 * Output:
		 * Function Owner: Poonam Kadam
		 * Last Update : 31/12/2018 Poonam Kadam - Updated Xpath for Extra IS Unit and added sCell
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyExtraISUnitsForPFI_FIMLoops(string sFileName,string sAddEXIDevicesSheet, string sAddIsDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddEXIDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared string type
			string sType, sLabelName, ExpectedMaxISUnits, ChangedValue, PanelType, sRowNumber, sISUnits, sExpectedColorCode, sActualColorCode, CellNumber;
			
			for(int i=8; i<=rows; i++)
			{
				
				ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				ChangedValue =  ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				ExpectedMaxISUnits = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				sISUnits = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				sCell=(i-8).ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[4,7]).Value.ToString();
				Report.Log(ReportLevel.Info,"All data imported from excel sheet");
				
				if(PanelType.Equals("FIM"))
				{
					sRow = (i+5).ToString();
				}
				else
				{
					sRow = (i+2).ToString();
				}
				
				//Add EXI devices from gallery
				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
				Report.Log(ReportLevel.Info,"Device "+ModelNumber+" Added");
				repo.ProfileConsys1.cell_CableCapacitance.DoubleClick();
				
				// Change the cable capacitance value
				repo.ProfileConsys1.txt_CableCapacitance.PressKeys((ChangedValue) +"{ENTER}");
				Report.Log(ReportLevel.Info,"Cable capacitance changed to: "+ChangedValue);
			}

			for(int i=8; i<=rows; i++)
			{
				sMaxExtraISUnits= ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				sExtraISUnitsReq= ((Range)Excel_Utilities.ExcelRange.Cells[i,10]).Value.ToString();
				sExpectedColorCode= ((Range)Excel_Utilities.ExcelRange.Cells[i,11]).Value.ToString();
				CellNumber=(i-8).ToString();
				sCell="["+CellNumber+"]";
				sRow=(i+5).ToString();
				//Verify Extra Is units required
				verifyExtraISUnitsReq(sExtraISUnitsReq);
				Report.Log(ReportLevel.Info,"Extra IS Units required for EXI row " + (i-7) + " verified as: "+sExtraISUnitsReq);

				//Verify Max Extra IS Units
				verifyMaxExtraISUnits(sMaxExtraISUnits);
				Report.Log(ReportLevel.Info,"Max Extra IS Units verified as: "+sMaxExtraISUnits);

				//Verify Progress bar color code for extra IS units
				sActualColorCode= getProgressBarColorForExtraISUnits();
				Devices_Functions.VerifyPercentage(sExpectedColorCode, sActualColorCode);
				Report.Log(ReportLevel.Info,"Progress bar color code for EXI row " + (i-7) + " verified as: "+sExpectedColorCode);
				sRow= (i+6).ToString();
			}
			
			//Click points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			//Import data from excel sheet
			int j=8;
			ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[j,1]).Value.ToString();
			sType = ((Range)Excel_Utilities.ExcelRange.Cells[j,2]).Value.ToString();
			PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[4,7]).Value.ToString();
			ChangedValue =  ((Range)Excel_Utilities.ExcelRange.Cells[j,5]).Value.ToString();
			Report.Log(ReportLevel.Info,"All data imported from excel sheet");
			
			if(PanelType.Equals("FIM"))
			{
				sRow = (j+5).ToString();
			}
			else
			{
				sRow = (j+2).ToString();
			}
			
			//Add one EXI800 devices from gallery
			Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
			repo.ProfileConsys1.cell_CableCapacitance.DoubleClick();
			
			// Change the cable capacitance value
			repo.ProfileConsys1.txt_CableCapacitance.PressKeys((ChangedValue) +"{ENTER}");
			Report.Log(ReportLevel.Info,"Cable capacitance changed to: "+ChangedValue);
			
			//Close excel
			Excel_Utilities.CloseExcel();
			
			//Open excel
			Excel_Utilities.OpenExcelFile(sFileName,sAddIsDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int ISrows= Excel_Utilities.ExcelRange.Rows.Count;
			
			//Click points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			//Add IS devices to EXI800
			for(int i=8; i<=ISrows; i++)
			{
				ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sRowNumber= (i-7).ToString();
				Report.Log(ReportLevel.Info,"Row number is: "+sRowNumber);
				Devices_Functions.SelectPointsGridRow(sRowNumber);
				Report.Log(ReportLevel.Info,"EXI800 on Row number: "+sRowNumber+"is Selected");
				
				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
				Report.Log(ReportLevel.Info,"Device"+ModelNumber+" Added to loop");
				sLabelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,10]).Value.ToString();
				
				Report.Log(ReportLevel.Success,"Label is "+ sLabelName);
			}
			
			//Verify IS units for each EXI
			for(int i=8; i<=ISrows; i++)
			{
				sISUnits = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				Report.Log(ReportLevel.Info,"Verify IS units as: "+sISUnits);
				sRow= (i+12).ToString();
				Report.Log(ReportLevel.Info,"EXI800 row set to: "+sRow);
				verifyISUnits(sISUnits,sRow);
				Report.Log(ReportLevel.Info,"IS units is verified for " +i+ " EXI800");
			}
			
			//Click points grid
			repo.ProfileConsys1.tab_Points.Click();
			
			//Add one IS device to EXI800
			int k= 8;
			ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[k,1]).Value.ToString();
			sType = ((Range)Excel_Utilities.ExcelRange.Cells[k,2]).Value.ToString();
			sRowNumber= (k-7).ToString();
			Report.Log(ReportLevel.Info,"Row number is: "+sRowNumber);
			Devices_Functions.SelectPointsGridRow(sRowNumber);
			Report.Log(ReportLevel.Info,"EXI800 on Row number: "+sRowNumber+"is Selected");
			
			Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
			Report.Log(ReportLevel.Info,"Device"+ModelNumber+" Added to loop");
			sLabelName =  ((Range)Excel_Utilities.ExcelRange.Cells[k,10]).Value.ToString();
			Report.Log(ReportLevel.Success,"Label is "+ sLabelName);
			
			//Click physical layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			//Verify Extra IS units for each EXI800
			for(int i=8; i<=rows; i++)
			{
				sMaxExtraISUnits= ((Range)Excel_Utilities.ExcelRange.Cells[i,11]).Value.ToString();
				sExtraISUnitsReq= ((Range)Excel_Utilities.ExcelRange.Cells[i,12]).Value.ToString();
				sExpectedColorCode= ((Range)Excel_Utilities.ExcelRange.Cells[i,13]).Value.ToString();
				sRow= (i+12).ToString();
				//Verify Extra Is units required
				verifyExtraISUnitsReq(sExtraISUnitsReq);
				Report.Log(ReportLevel.Info,"Extra IS Units required for EXI row " + (i-7) + " verified as: "+sExtraISUnitsReq);

				//Verify Max Extra IS Units
				verifyMaxExtraISUnits(sMaxExtraISUnits);
				Report.Log(ReportLevel.Info,"Max Extra IS Units verified as: "+sMaxExtraISUnits);

				//Verify Progress bar color code for extra IS units
				sActualColorCode= getProgressBarColorForExtraISUnits();
				Devices_Functions.VerifyPercentage(sExpectedColorCode, sActualColorCode);
				Report.Log(ReportLevel.Info,"Progress bar color code for EXI row " + (i-7) + " verified as: "+sExpectedColorCode);
				//sRow= (i+6).ToString();
			}
			
			//Click points grid tab
			repo.ProfileConsys1.tab_Points.Click();
			
			//Setting the row count j to 8 to set sRow value and fetch values from excel
			int l=8;
			sMaxExtraISUnits= ((Range)Excel_Utilities.ExcelRange.Cells[l,11]).Value.ToString();
			sExtraISUnitsReq= ((Range)Excel_Utilities.ExcelRange.Cells[l,12]).Value.ToString();
			sExpectedColorCode= ((Range)Excel_Utilities.ExcelRange.Cells[l,13]).Value.ToString();
			sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[l,9]).Value.ToString();
			
			//Delete IS device
			Devices_Functions.DeleteDeviceUsingLabel(sLabelName);
			Report.Log(ReportLevel.Info,"Device "+sLabelName+" deleted");
			
			// Click on Physical Layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			sRow= (l+12).ToString();
			Report.Log(ReportLevel.Info,"Extra IS Units required for EXI row " + sRow);
			
			//Verify Extra Is units required
			verifyExtraISUnitsReq(sExtraISUnitsReq);
			Report.Log(ReportLevel.Info,"Extra IS Units required for EXI row " + (l-2) + " verified as: "+sExtraISUnitsReq);

			//Verify Max Extra IS Units
			//verifyMaxExtraISUnits(sMaxExtraISUnits);
			Report.Log(ReportLevel.Info,"Max Extra IS Units verified as: "+sMaxExtraISUnits);

			//Verify Progress bar color code for extra IS units
			sActualColorCode= getProgressBarColorForExtraISUnits();
			Devices_Functions.VerifyPercentage(sExpectedColorCode, sActualColorCode);
			Report.Log(ReportLevel.Info,"Progress bar color code for EXI row " + (l-2) + " verified as: "+sExpectedColorCode);
			
			//Click points grid tab
			repo.ProfileConsys1.tab_Points.Click();
			
			//Setting the row count l to 8 to set sRow value and fetch values from excel
			int m=8;
			sMaxExtraISUnits= ((Range)Excel_Utilities.ExcelRange.Cells[m,11]).Value.ToString();
			sExtraISUnitsReq= ((Range)Excel_Utilities.ExcelRange.Cells[m,12]).Value.ToString();
			sExpectedColorCode= ((Range)Excel_Utilities.ExcelRange.Cells[m,13]).Value.ToString();
			sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[m,10]).Value.ToString();
			
			//Delete IS device
			Devices_Functions.DeleteDeviceUsingLabel(sLabelName);
			Report.Log(ReportLevel.Info,"Device "+sLabelName+" deleted");
			sRow= (m+17).ToString();
			
			// Click on Physical Layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			for(int n=8;n<=rows-1;n++)
			{
				sMaxExtraISUnits= ((Range)Excel_Utilities.ExcelRange.Cells[n,11]).Value.ToString();
				sExtraISUnitsReq= ((Range)Excel_Utilities.ExcelRange.Cells[n,14]).Value.ToString();
				sExpectedColorCode= ((Range)Excel_Utilities.ExcelRange.Cells[n,13]).Value.ToString();
				sRow= (n+10).ToString();
				//Verify Extra Is units required
				verifyExtraISUnitsReq(sExtraISUnitsReq);
				Report.Log(ReportLevel.Info,"Extra IS Units required for EXI row " + (n-7) + " verified as: "+sExtraISUnitsReq);

				//Verify Max Extra IS Units
				verifyMaxExtraISUnits(sMaxExtraISUnits);
				Report.Log(ReportLevel.Info,"Max Extra IS Units verified as: "+sMaxExtraISUnits);

				//Verify Progress bar color code for extra IS units
				sActualColorCode= getProgressBarColorForExtraISUnits();
				Devices_Functions.VerifyPercentage(sExpectedColorCode, sActualColorCode);
				Report.Log(ReportLevel.Info,"Progress bar color code for EXI row " + (n-7) + " verified as: "+sExpectedColorCode);
				//sRow= (n+6).ToString();
			}
		}

		
		/********************************************************************
		 * Function Name: VerifyProgressBarIndicatorForISUnits
		 * Function Details: Verify progress bar indicator for Intrinsically-safe Units
		 * Parameter/Arguments: sFileName, sAddDevicesSheet, sAddIsDevicesSheet
		 * Output:
		 * Function Owner: Poonam Kadam
		 * Last Update : 31/12/2018
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyProgressBarIndicatorForISUnits(string sFileName,string sAddEXIDevicesSheet)
		{
			//Open excel sheet and read its values
			Excel_Utilities.OpenExcelFile(sFileName,sAddEXIDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared string type
			string sType, sLabelName, PanelType, sExpectedColorCode, sActualColorCode;
			
			for(int i=8; i<=rows; i++)
			{
				
				ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sExpectedColorCode= ((Range)Excel_Utilities.ExcelRange.Cells[i,11]).Value.ToString();
				
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[4,7]).Value.ToString();
				Report.Log(ReportLevel.Info,"All data imported from excel sheet");
				
				if(PanelType.Equals("FIM"))
				{
					sRow = (i+5).ToString();
				}
				else
				{
					sRow = (i+2).ToString();
				}
				
				//Add EXI devices from gallery
				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
				Report.Log(ReportLevel.Info,"Device "+ModelNumber+" Added");
				
				//Click physical layout tab
				repo.ProfileConsys1.tab_PhysicalLayout.Click();
				sRowIndex=(i+1).ToString();
				
				//Verify Progress bar color code for extra IS units
				sActualColorCode= getProgressBarColorForISUnits();
				Devices_Functions.VerifyPercentage(sExpectedColorCode, sActualColorCode);
				Report.Log(ReportLevel.Info,"Progress bar color code for EXI row " + (i-7) + " verified as: "+sExpectedColorCode);
				
				//Click physical layout tab
				repo.ProfileConsys1.tab_Points.Click();
			}
			
			//Setting the row count l to 8 to set sRow value and fetch values from excel
			int m=8;
			sExpectedColorCode= ((Range)Excel_Utilities.ExcelRange.Cells[m,11]).Value.ToString();
			sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[m,1]).Value.ToString();
			
			//Delete IS device
			Devices_Functions.DeleteDeviceUsingLabel(sLabelName);
			Report.Log(ReportLevel.Info,"Device "+sLabelName+" deleted");
			
			//Click physical layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			sRowIndex=(m+1).ToString();
			
			//Verify Progress bar color code for extra IS units
			sActualColorCode= getProgressBarColorForISUnits();
			Devices_Functions.VerifyPercentage(sExpectedColorCode, sActualColorCode);
			Report.Log(ReportLevel.Info,"Progress bar color code for EXI row " + (m-7) + " verified as: "+sExpectedColorCode);
			
		}
		

		/********************************************************************
		 * Function Name: getProgressBarColorForISUnits
		 * Function Details: Method to verify progress bar color for IS units
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Poonam Kadam
		 * Last Update : 31/12/2018
		 ********************************************************************/
		[UserCodeMethod]
		public static string getProgressBarColorForISUnits()
		{
			string actualColour;
			return actualColour = repo.FormMe.ISUnitProgressBar.GetAttributeValue<string>("foreground");
			
		}

		/********************************************************************
		 * Function Name: verifyExtraISUnitsReq
		 * Function Details: verify extra IS units required
		 * Parameter/Arguments: expectedExtraISUnits
		 * Output:
		 * Function Owner: Poonam Kadam
		 * Last Update : 26/12/2018
		 ********************************************************************/
		public static void verifyExtraISUnitsReq(string expectedExtraISUnitsReq)
		{
			//Click on Physical layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			sExtraISUnitsReq=expectedExtraISUnitsReq;
			// Store IS unit value in ActualMaxISUnit variable
			string ActualExtraISUnitsReq = repo.FormMe.ExtraISUnitsReq.TextValue;
			// Compare ActualISUnit and expectedISUnit values and then displaying result
			if(ActualExtraISUnitsReq.Equals(expectedExtraISUnitsReq))
			{
				Report.Log(ReportLevel.Success,"Extra IS Units " + ActualExtraISUnitsReq + " is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Extra IS Unit is not displayed correctly, Extra IS Units displayed as: " +ActualExtraISUnitsReq + " instead of : "+expectedExtraISUnitsReq);
			}
		}

		/********************************************************************
		 * Function Name: verifyMaxExtraISUnits
		 * Function Details: To verify maximum Extra IS units value
		 * Parameter/Arguments: expectedMax Extra IS unit value
		 * Output:
		 * Function Owner: Poonam Kadam
		 * Last Update : 31/12/2018
		 ********************************************************************/
		public static void verifyMaxExtraISUnits(string expectedMaxExtraISUnits)
		{
			string actualMaxExtraISUnits = repo.FormMe.MaxExtraISUnits.TextValue;
			
			// Compare ActualMaxExtraISUnit and expectedMaxExtraISUnit values and then displaying result
			if(actualMaxExtraISUnits.Equals(expectedMaxExtraISUnits))
			{
				Report.Log(ReportLevel.Success,"Max Extra IS Unit " + actualMaxExtraISUnits + " is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Max Extra IS Unit is not displayed correctly, Max IS Units displayed as: " +actualMaxExtraISUnits + " instead of : "+expectedMaxExtraISUnits);
			}
		}

		/********************************************************************
		 * Function Name: getProgressBarColor
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Poonam Kadam
		 * Last Update : 31/12/2018
		 ********************************************************************/
		[UserCodeMethod]
		public static string getProgressBarColorForExtraISUnits()
		{
			string actualColour;
			return actualColour = repo.FormMe.ExtraISUnitProgressBar.GetAttributeValue<string>("foreground");
			
		}


		/********************************************************************
		 * Function Name: VerifyISDevicesOnAddingEXI
		 * Function Details: Verify the IS device state when we add  EXI800
		 * Parameter/Arguments: sFileName, sAddEXIDevicesSheet, sVerifyISDevicesStateSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 29/01/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyISDevicesOnAddingEXI(string sFileName,string sAddEXIDevicesSheet, string sVerifyISDevicesStateSheet)
		{
			//Add devices in loop A,
			Excel_Utilities.OpenExcelFile(sFileName,sAddEXIDevicesSheet);
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			string sType,sRowNumber;
			sRowNumber= 1.ToString();
			for(int i=8; i<=rows; i++)
			{
				ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();

				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
			}
			
			//Close excel
			Excel_Utilities.CloseExcel();
			
			// Open Another excel to verify IS devices state
			Excel_Utilities.OpenExcelFile(sFileName,sVerifyISDevicesStateSheet);
			
			// Count the number of rows in excel
			rows = Excel_Utilities.ExcelRange.Rows.Count;
			
			for(int i=8; i<=rows; i++)
			{
				
				string sDeviceName = ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				string state =  ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				
				// Verify gallery item state
				Devices_Functions.VerifyGalleryItem(sType,sDeviceName,state);
				
				// Click on Points tab
				repo.ProfileConsys1.tab_Points.Click();
			}
			
			//Close excel
			Excel_Utilities.CloseExcel();
		}
		
		/********************************************************************
		 * Function Name: VerifyISDevicesOnAddingNonEXI
		 * Function Details: Verify the IS device state when we add non-EXI device
		 * Parameter/Arguments: sFileName, sAddNonEXIDevicesSheet, sVerifyISDevicesStateSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 29/01/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyISDevicesOnAddingNonEXI(string sFileName,string sAddNonEXIDevicesSheet, string sVerifyISDevicesStateSheet)
		{
			//Add devices in loop A,
			Excel_Utilities.OpenExcelFile(sFileName,sAddNonEXIDevicesSheet);
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			string sType;
			for(int i=8; i<=rows; i++)
			{
				ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();

				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
			}
			
			//Close excel
			Excel_Utilities.CloseExcel();
			
			// Open Another excel to verify IS devices state
			Excel_Utilities.OpenExcelFile(sFileName,sVerifyISDevicesStateSheet);
			
			// Count the number of rows in excel
			rows = Excel_Utilities.ExcelRange.Rows.Count;
			
			for(int i=8; i<=rows; i++)
			{
				string sDeviceName = ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				string state =  ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				
				// Verify gallery item state
				Devices_Functions.VerifyGalleryItem(sType,sDeviceName,state);
				
				// Click on Points tab
				repo.ProfileConsys1.tab_Points.Click();
			}
			
			//Close excel
			Excel_Utilities.CloseExcel();
		}
		
		/***********************************************************************************************************************************
		 * Function Name: VerifyAdditionISDevicesOnEXI
		 * Function Details: To verify addition of IS devices and observe in Physical layout
		 * Parameter/Arguments: sFileName, sAddEXIDeviceSheet, sISDevicesSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 01/02/2019 Updated on 07/02/2019 - Alpesh Dhakad - Updated device order verification steps
		 ***********************************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyAdditionISDevicesOnEXI(string sFileName,string sAddEXIDeviceSheet, string sISDevicesSheet)
		{
			//Add devices in loop A,
			Excel_Utilities.OpenExcelFile(sFileName,sAddEXIDeviceSheet);
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			string sType,sRowNumber;
			for(int i=8; i<=rows; i++)
			{
				ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();

				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
			}
			//Close excel
			Excel_Utilities.CloseExcel();
			
			// Select Point grid and right click on it
			Mouse.Click(repo.FormMe.HeadersPanel1, System.Windows.Forms.MouseButtons.Right);
			
			// Click Show column chooser to select column
			repo.ContextMenu.ShowColumnChooser.Click();
			
			// Click on Device order checkbox to add the column in points grid
			repo.ProfileConsys.chkBox_DeviceOrder1.Click();
			Report.Log(ReportLevel.Info," Device order column added successfully ");
			
			// Close column choose window
			repo.ProfileConsys.btn_CloseColumnChooser.Click();
			
			// Open Another excel to verify IS devices state
			Excel_Utilities.OpenExcelFile(sFileName,sISDevicesSheet);
			
			// Count the number of rows in excel
			rows = Excel_Utilities.ExcelRange.Rows.Count;
			
			for(int i=8; i<=rows; i++)
			{
				sDeviceOrderRow= (i-6).ToString();
				string sDeviceName = ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				string state =  ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sDeviceOrderName = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				sRowNumber= (1).ToString();
				
				// Verify gallery item state
				Devices_Functions.AddDevicesfromGallery(sDeviceName,sType);
				Report.Log(ReportLevel.Info,"Device " +sDeviceName+ " Added to loop");
				
				// Click on Device Name
				repo.FormMe.txt_DeviceOrderLabel.Click();
				
				// Click on Device Name
				repo.FormMe.txt_DeviceOrderLabel.Click();
				
				string actualDeviceOrderValue = repo.FormMe.txt_DeviceOrderLabel.TextValue;
				
				// Compare actualDeviceOrderValue and sDeviceOrderName values and then displaying result
				if(actualDeviceOrderValue.Equals(sDeviceOrderName))
				{
					Report.Log(ReportLevel.Success, "Device " +sDeviceName+ " with " +sDeviceOrderName+ " added successfully and displaying correct device order");
				}
				else
				{
					Report.Log(ReportLevel.Failure, "Device " +sDeviceName+ " with " +sDeviceOrderName+ " not added or not displaying correct device order");
				}
				
				// Click on first added EXI800
				Devices_Functions.SelectPointsGridRow(sRowNumber);
				
			}
			
			//Close excel
			Excel_Utilities.CloseExcel();
			
			// Click on Physical Layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			// Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			// Click on Physical Layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			// Open Another excel to verify IS devices state
			Excel_Utilities.OpenExcelFile(sFileName,sISDevicesSheet);
			
			// Count the number of rows in excel
			rows = Excel_Utilities.ExcelRange.Rows.Count;
			
			
			for(int i=8; i<=rows; i++)
			{
				string sDeviceName = ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sDeviceOrderName = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				sPhysicalLayoutDeviceIndex = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				//string actualIndex = repo.FormMe.lst_PhysicalLayoutDevices.Index.ToString();
				
				string actualDeviceName = repo.FormMe.lst_PhysicalLayoutDevice.TextValue;

				// Compare actualIndex and sPhysicalLayoutDeviceIndex values and then displaying result
				if(actualDeviceName.Equals(sDeviceName))
				{
					Report.Log(ReportLevel.Success, "Device " +sDeviceName+ " added successfully and displaying correctly in Physical Layout");
				}
				else
				{
					Report.Log(ReportLevel.Failure, "Device " +sDeviceName+ " not added or not displaying correctly in Physical Layout");
				}
				
				
			}
			//Close excel
			Excel_Utilities.CloseExcel();
			
		}
		
		/********************************************************************
		 * Function Name: VerifyDeletionOfISDevicesOnEXI
		 * Function Details: To verify deletion of IS devices and observe in Physical layout
		 * Parameter/Arguments: sFileName, sDeleteEXIDeviceSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 04/02/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyDeletionOfISDevicesOnEXI(string sFileName,string sDeleteEXIDeviceSheet, string verifyISDevices)
		{
			//Delete devices from loop A,
			Excel_Utilities.OpenExcelFile(sFileName,sDeleteEXIDeviceSheet);
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			for(int i=8; i<=rows; i++)
			{
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				
				// Click on Points tab
				repo.ProfileConsys1.tab_Points.Click();
				
				
				Devices_Functions.DeleteDevices(sFileName,sDeleteEXIDeviceSheet);
			}
			//Close excel
			Excel_Utilities.CloseExcel();
			
			// Click on Physical Layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			// Open Another excel to verify IS devices state
			Excel_Utilities.OpenExcelFile(sFileName,verifyISDevices);
			
			// Count the number of rows in excel
			rows = Excel_Utilities.ExcelRange.Rows.Count;
			
			for(int i=8; i<=rows; i++)
			{
				string sDeviceName = ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sDeviceOrderName = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				sPhysicalLayoutDeviceIndex = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				
				string actualDeviceName = repo.FormMe.lst_PhysicalLayoutDevice.TextValue;

				// Compare actualIndex and sPhysicalLayoutDeviceIndex values and then displaying result
				if(actualDeviceName.Equals(sDeviceName))
				{
					Report.Log(ReportLevel.Failure, "Device " +sDeviceName+ " not removed successfully from Physical Layout");
				}
				else
				{
					Report.Log(ReportLevel.Success, "Device " +sDeviceName+ " removed successfully from Physical Layout");
					
				}
			}
			//Close excel
			Excel_Utilities.CloseExcel();
		}
		
		/********************************************************************
		 * Function Name: VerifyAdditionOfMultipleISDevicesOnEXI
		 * Function Details: To verify addition of multiple IS devices and observe in Physical layout
		 * Parameter/Arguments: sFileName, sAddEXIDeviceSheet, sISDevicesSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 06/02/2019   Alpesh Dhakad - 01/08/2019 - Updated test scripts as per new build and xpaths
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyAdditionOfMultipleISDevicesOnEXI(string sFileName,string sAddEXIDeviceSheet, string sISDevicesSheet)
		{
			//Add devices in loop A,
			Excel_Utilities.OpenExcelFile(sFileName,sAddEXIDeviceSheet);
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			string sType,sRowNumber;
			sRowNumber= (1).ToString();
			
			for(int i=8; i<=rows; i++)
			{
				ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();

				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
			}
			//Close excel
			Excel_Utilities.CloseExcel();
			
			// Select Point grid and right click on it
			Mouse.Click(repo.FormMe.HeadersPanel1, System.Windows.Forms.MouseButtons.Right);
			
			// Click Show column chooser to select column
			repo.ContextMenu.ShowColumnChooser.Click();
			
			// Click on Device order checkbox to add the column in points grid
			repo.ProfileConsys.chkBox_DeviceOrder1.Click();
			Report.Log(ReportLevel.Info," Device order column added successfully ");
			
			// Close column choose window
			repo.ProfileConsys.btn_CloseColumnChooser.Click();
			
			// Open Another excel to verify IS devices state
			Excel_Utilities.OpenExcelFile(sFileName,sISDevicesSheet);
			
			// Count the number of rows in excel
			rows = Excel_Utilities.ExcelRange.Rows.Count;
			
			for(int i=8; i<=rows; i++)
			{
				sDeviceOrderRow= (i-6).ToString();
				string sDeviceName = ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				string state =  ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sDeviceOrderName = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				sPhysicalLayoutDeviceIndex = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				string NumberOfDevicesCount = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();;
				
				int noOfDevices;
				
				int.TryParse(NumberOfDevicesCount, out noOfDevices);
				for (int j = 1; j <= noOfDevices; j++) {
					Devices_Functions.AddDevicesfromGallery(sDeviceName,sType);
					Report.Log(ReportLevel.Info,"Device " +sDeviceName+ "  Added to loop");
					
					// Click on first added EXI800
					Devices_Functions.SelectPointsGridRow(sRowNumber);
					
				}
				
				// Verify gallery item state
				Devices_Functions.VerifyGalleryItem(sType,sDeviceName,state);
				
				repo.FormMe.Loop_A1.Click();
				Delay.Milliseconds(200);
				
				// Click on Physical Layout tab
				repo.ProfileConsys1.tab_PhysicalLayout.Click();
				Delay.Milliseconds(500);
				
				repo.FormMe.Loop_A1.Click();
				Delay.Milliseconds(200);
				
				// Click on Points tab
				repo.ProfileConsys1.tab_Points.Click();
				Delay.Milliseconds(500);
				
				// Click on Physical Layout tab
				repo.ProfileConsys1.tab_PhysicalLayout.Click();
				Delay.Milliseconds(500);
				
				// Click on Points tab
				repo.ProfileConsys1.tab_Points.Click();
				Delay.Milliseconds(500);
				
				// Click on Physical Layout tab
				repo.ProfileConsys1.tab_PhysicalLayout.Click();
				Delay.Milliseconds(500);
				
				// Split Device name and then add devices as per the device name and number of devices from main gallery
				string[] PhysicalLayoutIndex  = sPhysicalLayoutDeviceIndex.Split(',');
				int splitCount  = sPhysicalLayoutDeviceIndex.Split(',').Length;
				
				for(int k=0; k<splitCount; k++){
					sPhysicalLayoutDeviceIndex = PhysicalLayoutIndex[k];
					
					string actualDeviceName = repo.FormMe.lst_PhysicalLayoutDevice.TextValue;

					// Compare actualIndex and sPhysicalLayoutDeviceIndex values and then displaying result
					if(actualDeviceName.Equals(sDeviceName))
					{
						Report.Log(ReportLevel.Success, "Device " +sDeviceName+ " added successfully with " +sPhysicalLayoutDeviceIndex+ " and displaying correctly in Physical Layout");
					}
					else
					{
						Report.Log(ReportLevel.Failure, "Device " +sDeviceName+ " not added or not displaying correctly in Physical Layout");
					}
					
				}
				// Click on Points tab
				repo.ProfileConsys1.tab_Points.Click();
				Delay.Milliseconds(500);
				
				// Click on first added EXI800
				Devices_Functions.SelectPointsGridRow(sRowNumber);
				
				// Delete Exi800 devices
				repo.ProfileConsys1.btn_Delete.Click();
				
				ModelNumber = ((Range)Excel_Utilities.ExcelRange.Cells[4,5]).Value.ToString();
				string sType1 = ((Range)Excel_Utilities.ExcelRange.Cells[4,6]).Value.ToString();
				
				// Add EXI800 device
				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType1);
			}
			//Close excel
			Excel_Utilities.CloseExcel();
			
		}
		
		/********************************************************************
		 * Function Name: VerifyEnableDisableOfMultipleISDevicesOnEXI
		 * Function Details: To verify addition of multiple IS devices and observe in Physical layout
		 * Parameter/Arguments: sFileName, sAddEXIDeviceSheet, sISDevicesSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 06/02/2019  Alpesh Dhakad - 01/08/2019 - Updated test scripts as per new build and xpaths
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyEnableDisableOfMultipleISDevicesOnEXI(string sFileName,string sAddEXIDeviceSheet, string sISDevicesSheet)
		{
			//Add devices in loop A,
			Excel_Utilities.OpenExcelFile(sFileName,sAddEXIDeviceSheet);
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			string sType,sRowNumber;
			sRowNumber= (1).ToString();
			
			for(int i=8; i<=rows; i++)
			{
				ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();

				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
			}
			//Close excel
			Excel_Utilities.CloseExcel();
			
			// Select Point grid and right click on it
			Mouse.Click(repo.FormMe.HeadersPanel1, System.Windows.Forms.MouseButtons.Right);
			
			// Click Show column chooser to select column
			repo.ContextMenu.ShowColumnChooser.Click();
			
			// Click on Device order checkbox to add the column in points grid
			repo.ProfileConsys.chkBox_DeviceOrder1.Click();
			Report.Log(ReportLevel.Info," Device order column added successfully ");
			
			// Close column choose window
			repo.ProfileConsys.btn_CloseColumnChooser.Click();
			
			// Open Another excel to verify IS devices state
			Excel_Utilities.OpenExcelFile(sFileName,sISDevicesSheet);
			
			// Count the number of rows in excel
			rows = Excel_Utilities.ExcelRange.Rows.Count;
			
			for(int i=8; i<=rows; i++)
			{
				sDeviceOrderRow= (i-6).ToString();
				string sDeviceName = ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				string state =  ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sDeviceOrderName = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				sPhysicalLayoutDeviceIndex = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				string NumberOfDevicesCount = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();;
				string changedState =  ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				
				int noOfDevices;
				
				int.TryParse(NumberOfDevicesCount, out noOfDevices);
				for (int j = 1; j <= noOfDevices; j++) {
					Devices_Functions.AddDevicesfromGallery(sDeviceName,sType);
					Report.Log(ReportLevel.Info,"Device " +sDeviceName+ "  Added to loop");
					
					// Click on first added EXI800
					Devices_Functions.SelectPointsGridRow(sRowNumber);
					
				}
				
				// Verify gallery item state
				Devices_Functions.VerifyGalleryItem(sType,sDeviceName,state);
				
				repo.FormMe.Loop_A1.Click();
				Delay.Milliseconds(200);
				
				// Click on Physical Layout tab
				repo.ProfileConsys1.tab_PhysicalLayout.Click();
				Delay.Milliseconds(500);
				
				repo.FormMe.Loop_A1.Click();
				Delay.Milliseconds(200);
				
				// Click on Points tab
				repo.ProfileConsys1.tab_Points.Click();
				Delay.Milliseconds(500);
				
				// Click on Physical Layout tab
				repo.ProfileConsys1.tab_PhysicalLayout.Click();
				Delay.Milliseconds(500);
				
				// Click on Points tab
				repo.ProfileConsys1.tab_Points.Click();
				Delay.Milliseconds(500);
				
				// Click on Physical Layout tab
				repo.ProfileConsys1.tab_PhysicalLayout.Click();
				Delay.Milliseconds(500);
				
				// Split Device name and then add devices as per the device name and number of devices from main gallery
				string[] PhysicalLayoutIndex  = sPhysicalLayoutDeviceIndex.Split(',');
				int splitCount  = sPhysicalLayoutDeviceIndex.Split(',').Length;
				
				for(int k=0; k<splitCount; k++){
					sPhysicalLayoutDeviceIndex = PhysicalLayoutIndex[k];
					
					string actualDeviceName = repo.FormMe.lst_PhysicalLayoutDevice.TextValue;

					// Compare actualIndex and sPhysicalLayoutDeviceIndex values and then displaying result
					if(actualDeviceName.Equals(sDeviceName))
					{
						Report.Log(ReportLevel.Success, "Device " +sDeviceName+ " added successfully with " +sPhysicalLayoutDeviceIndex+ " and displaying correctly in Physical Layout");
					}
					else
					{
						Report.Log(ReportLevel.Failure, "Device " +sDeviceName+ " not added or not displaying correctly in Physical Layout");
					}
					
				}
				// Click on Points tab
				repo.ProfileConsys1.tab_Points.Click();
				Delay.Milliseconds(500);
				
				// Click on Device order label
				repo.FormMe.txt_DeviceOrderLabel.Click();
				
				// Click on Delete button to delete one device
				repo.ProfileConsys1.btn_Delete.Click();
				
				// Click on first added EXI800
				Devices_Functions.SelectPointsGridRow(sRowNumber);
				
				// Verify gallery item state
				Devices_Functions.VerifyGalleryItem(sType,sDeviceName,changedState);
				
				// Delete Exi800 devices
				repo.ProfileConsys1.btn_Delete.Click();
				
				ModelNumber = ((Range)Excel_Utilities.ExcelRange.Cells[4,5]).Value.ToString();
				string sType1 = ((Range)Excel_Utilities.ExcelRange.Cells[4,6]).Value.ToString();
				
				// Add EXI800 device
				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType1);
				
			}
			//Close excel
			Excel_Utilities.CloseExcel();
			
		}
		
		
		/********************************************************************
		 * Function Name:VerifyConnectedISDevicesOnEXI
		 * Function Details: To verify connected IS devices on EXI
		 * Parameter/Arguments: sFileName, sAddEXIDeviceSheet, sISDevicesSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 07/02/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyConnectedISDevicesOnEXI(string sFileName,string sAddEXIDeviceSheet, string sISDevicesSheet)
		{
			//Add devices in loop A,
			Excel_Utilities.OpenExcelFile(sFileName,sAddEXIDeviceSheet);
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			string sType,sRowNumber;
			for(int i=8; i<=rows; i++)
			{
				ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();

				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
			}
			//Close excel
			Excel_Utilities.CloseExcel();
			
			// Select Point grid and right click on it
			Mouse.Click(repo.FormMe.HeadersPanel1, System.Windows.Forms.MouseButtons.Right);
			
			// Click Show column chooser to select column
			repo.ContextMenu.ShowColumnChooser.Click();
			
			// Click on Device order checkbox to add the column in points grid
			repo.ProfileConsys.chkBox_DeviceOrder1.Click();
			Report.Log(ReportLevel.Info," Device order column added successfully ");
			
			// Close column choose window
			repo.ProfileConsys.btn_CloseColumnChooser.Click();
			
			// Open Another excel to verify IS devices state
			Excel_Utilities.OpenExcelFile(sFileName,sISDevicesSheet);
			
			// Count the number of rows in excel
			rows = Excel_Utilities.ExcelRange.Rows.Count;

			for(int i=8; i<=rows; i++)
			{
				sDeviceOrderRow= ((i+1)-6).ToString();
				string sDeviceName = ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				string state =  ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sDeviceOrderName = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				sRowNumber= (1).ToString();
				
				// Click on first added EXI800
				Devices_Functions.SelectPointsGridRow(sRowNumber);
				
				// Verify gallery item state
				Devices_Functions.AddDevicesfromGallery(sDeviceName,sType);
				Report.Log(ReportLevel.Info,"Device " +sDeviceName+ " Added to loop");
				
				// Click on Device Name
				repo.FormMe.txt_DeviceOrderLabel.Click();
				
				string actualDeviceOrderValue = repo.FormMe.txt_DeviceOrderLabel.TextValue;
				
				// Compare actualDeviceOrderValue and sDeviceOrderName values and then displaying result
				if(actualDeviceOrderValue.Equals(sDeviceOrderName))
				{
					Report.Log(ReportLevel.Success, "Device " +sDeviceName+ " with " +sDeviceOrderName+ " added successfully and displaying correct device order");
				}
				else
				{
					Report.Log(ReportLevel.Failure, "Device " +sDeviceName+ " with " +sDeviceOrderName+ " not added or not displaying correct device order");
				}
				
				
				// Click on first added EXI800
				Devices_Functions.SelectPointsGridRow(sRowNumber);
			}
			
			// Add devices for Second EXI
			for(int j=8; j<=rows; j++)
			{
				sDeviceOrderRow= ((j+4)-6).ToString();
				string sDeviceName = ((Range)Excel_Utilities.ExcelRange.Cells[j,7]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[j,8]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[j,9]).Value.ToString();
				string state =  ((Range)Excel_Utilities.ExcelRange.Cells[j,10]).Value.ToString();
				sDeviceOrderName = ((Range)Excel_Utilities.ExcelRange.Cells[j,11]).Value.ToString();
				sRowNumber= (2).ToString();
				
				// Click on first added EXI800
				Devices_Functions.SelectPointsGridRow(sRowNumber);
				
				// Verify gallery item state
				Devices_Functions.AddDevicesfromGallery(sDeviceName,sType);
				Report.Log(ReportLevel.Info,"Device " +sDeviceName+ " Added to loop");
				
				if(repo.FormMe.txt_DeviceOrderLabelInfo.Exists())
				{
					// Click on Device Name
					repo.FormMe.txt_DeviceOrderLabel.Click();
					
					string actualDeviceOrderValue = repo.FormMe.txt_DeviceOrderLabel.TextValue;
					
					// Compare actualDeviceOrderValue and sDeviceOrderName values and then displaying result
					if(actualDeviceOrderValue.Equals(sDeviceOrderName))
					{
						Report.Log(ReportLevel.Success, "Device " +sDeviceName+ " with " +sDeviceOrderName+ " added successfully and displaying correct device order");
					}
					else
					{
						Report.Log(ReportLevel.Failure, "Device " +sDeviceName+ " with " +sDeviceOrderName+ " not added or not displaying correct device order");
					}
				}
				
				// Click on first added EXI800
				Devices_Functions.SelectPointsGridRow(sRowNumber);
			}

			//Close excel
			Excel_Utilities.CloseExcel();
			
			// Set sRowNumber to 1 to select first EXI
			sRowNumber= (1).ToString();
			
			// Click on first added EXI800
			Devices_Functions.SelectPointsGridRow(sRowNumber);
			
			// Click on Delete button to delete first EXI
			repo.ProfileConsys1.btn_Delete.Click();
			
			// Open Another excel to verify IS devices
			Excel_Utilities.OpenExcelFile(sFileName,sISDevicesSheet);
			
			// Count the number of rows in excel
			rows = Excel_Utilities.ExcelRange.Rows.Count;

			for(int i=8; i<=rows; i++)
			{
				sDeviceOrderRow= ((i+4)-6).ToString();
				string sDeviceName = ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				string state =  ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sDeviceOrderName = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				sRowNumber= (1).ToString();
				
				
				// Click on Device Name
				repo.FormMe.txt_DeviceOrderLabel.Click();
				
				string actualDeviceOrderValue = repo.FormMe.txt_DeviceOrderLabel.TextValue;
				
				// Compare actualDeviceOrderValue and sDeviceOrderName values and then displaying result
				if(actualDeviceOrderValue.Equals(sDeviceOrderName))
				{
					Report.Log(ReportLevel.Failure, "Associated Device not removed successfully");
				}
				else
				{
					Report.Log(ReportLevel.Success, "Associated Device removed successfully after deleting EXI");
				}
				
			}
			//Close excel
			Excel_Utilities.CloseExcel();
		}
		
		/********************************************************************
		 * Function Name: VerifyDeviceOrderOfEXIandISDevices
		 * Function Details:  To verify device order for EXI and IS devices
		 * Parameter/Arguments: sFileName, sAddEXIDeviceSheet, sISDevicesSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 08/02/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyDeviceOrderOfEXIandISDevices(string sFileName,string sAddDeviceSheet, string sISDevicesSheet)
		{
			//Add devices in loop A,
			Excel_Utilities.OpenExcelFile(sFileName,sAddDeviceSheet);
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			string sType,sRowNumber;
			for(int i=8; i<=rows; i++)
			{
				ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();

				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
			}
			//Close excel
			Excel_Utilities.CloseExcel();
			
			// Select Point grid and right click on it
			Mouse.Click(repo.FormMe.HeadersPanel1, System.Windows.Forms.MouseButtons.Right);
			
			// Click Show column chooser to select column
			repo.ContextMenu.ShowColumnChooser.Click();
			
			// Click on Device order checkbox to add the column in points grid
			repo.ProfileConsys.chkBox_DeviceOrder1.Click();
			Report.Log(ReportLevel.Info," Device order column added successfully ");
			
			// Close column choose window
			repo.ProfileConsys.btn_CloseColumnChooser.Click();
			
			// Open Another excel to verify IS devices state
			Excel_Utilities.OpenExcelFile(sFileName,sISDevicesSheet);
			
			// Count the number of rows in excel
			rows = Excel_Utilities.ExcelRange.Rows.Count;

			for(int i=8; i<=rows; i++)
			{
				sDeviceOrderRow= ((i+1)-6).ToString();
				string sDeviceName = ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				string state =  ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sDeviceOrderName = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				sRowNumber= (2).ToString();
				
				// Click on first added EXI800
				Devices_Functions.SelectPointsGridRow(sRowNumber);
				
				// Verify gallery item state
				Devices_Functions.AddDevicesfromGallery(sDeviceName,sType);
				Report.Log(ReportLevel.Info,"Device " +sDeviceName+ " Added to loop");
				
				// Click on Device Name
				repo.FormMe.txt_DeviceOrderLabel.Click();
				
				// To retrieve the Device order text value
				string actualDeviceOrderValue = repo.FormMe.txt_DeviceOrderLabel.TextValue;
				
				// Compare actualDeviceOrderValue and sDeviceOrderName values and then displaying result
				if(actualDeviceOrderValue.Equals(sDeviceOrderName))
				{
					Report.Log(ReportLevel.Success, "Device " +sDeviceName+ " with " +sDeviceOrderName+ " added successfully and displaying correct device order");
				}
				else
				{
					Report.Log(ReportLevel.Failure, "Device " +sDeviceName+ " with " +sDeviceOrderName+ " not added or not displaying correct device order");
				}
				
				// Click on first added EXI800
				Devices_Functions.SelectPointsGridRow(sRowNumber);
			}
			
			// Set the Device order row value and device order name to delete the required device
			sDeviceOrderRow  = ((Range)Excel_Utilities.ExcelRange.Cells[5,5]).Value.ToString();
			sDeviceOrderName = ((Range)Excel_Utilities.ExcelRange.Cells[5,6]).Value.ToString();
			
			// Click on device order
			repo.FormMe.txt_DeviceOrderLabel.Click();
			
			// Click on delete button
			repo.ProfileConsys1.btn_Delete.Click();
			
			//Close excel
			Excel_Utilities.CloseExcel();
			
			
			// Open Another excel to verify IS devices state
			Excel_Utilities.OpenExcelFile(sFileName,sISDevicesSheet);
			
			// Count the number of rows in excel
			rows = Excel_Utilities.ExcelRange.Rows.Count;
			
			for(int k=9; k<=rows; k++)
			{
				sDeviceOrderRow= (k-5).ToString();
				sDeviceOrderName = ((Range)Excel_Utilities.ExcelRange.Cells[k,10]).Value.ToString();
				
				// Click on Device Name
				repo.FormMe.txt_DeviceOrderLabel.Click();
				
				// To retrieve the Device order text value
				string changedDeviceOrderValue = repo.FormMe.txt_DeviceOrderLabel.TextValue;
				
				// Compare actualDeviceOrderValue and sDeviceOrderName values and then displaying result
				if(changedDeviceOrderValue.Equals(sDeviceOrderName))
				{
					Report.Log(ReportLevel.Success, " Device order "+sDeviceOrderName+ " changed successfully");
				}
				else
				{
					Report.Log(ReportLevel.Failure, " Device order "+sDeviceOrderName+ " not changed successfully");
				}
			}
			//Close excel
			Excel_Utilities.CloseExcel();
		}
		
		/************************************************************************************************************************
		 * Function Name: VerifyEnableDisableISDevicesOnChangingCableCapacitance
		 * Function Details: Verify if we can add more IS devices to EXI800 when we decrease/increase the Cable Capacitance value
		 * Parameter/Arguments: sFileName, sAddEXIDeviceSheet, sISDevicesSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 12/02/2019
		 ************************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyEnableDisableISDevicesOnChangingCableCapacitance(string sFileName,string sAddEXIDeviceSheet, string sISDevicesSheet)
		{
			//Add devices in loop A,
			Excel_Utilities.OpenExcelFile(sFileName,sAddEXIDeviceSheet);
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			string sType,sRowNumber,defaultCableCapacitance,changedCableCapacitance;
			sRowNumber= (1).ToString();
			
			for(int i=8; i<=rows; i++)
			{
				ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();

				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
				
			}
			//Close excel
			Excel_Utilities.CloseExcel();
			
			// Select Point grid and right click on it
			Mouse.Click(repo.FormMe.HeadersPanel1, System.Windows.Forms.MouseButtons.Right);
			
			// Click Show column chooser to select column
			repo.ContextMenu.ShowColumnChooser.Click();
			
			// Click on Device order checkbox to add the column in points grid
			repo.ProfileConsys.chkBox_DeviceOrder1.Click();
			Report.Log(ReportLevel.Info," Device order column added successfully ");
			
			// Close column choose window
			repo.ProfileConsys.btn_CloseColumnChooser.Click();
			
			// Open Another excel to verify IS devices state
			Excel_Utilities.OpenExcelFile(sFileName,sISDevicesSheet);
			
			// Count the number of rows in excel
			rows = Excel_Utilities.ExcelRange.Rows.Count;
			
			for(int i=8; i<=rows; i++)
			{
				sDeviceOrderRow= (i-6).ToString();
				string sDeviceName = ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				string state =  ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sDeviceOrderName = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				defaultCableCapacitance = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				string NumberOfDevicesCount = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();;
				string changedState =  ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				changedCableCapacitance = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				
				int noOfDevices;
				
				int.TryParse(NumberOfDevicesCount, out noOfDevices);
				for (int j = 1; j <= noOfDevices; j++) {
					Devices_Functions.AddDevicesfromGallery(sDeviceName,sType);
					Report.Log(ReportLevel.Info,"Device " +sDeviceName+ "  Added to loop");
					
					// Click on first added EXI800
					Devices_Functions.SelectPointsGridRow(sRowNumber);
					
				}
				
				// Click on Device Name
				repo.FormMe.txt_DeviceOrderLabel.Click();
				
				// To retrieve the Device order text value
				string DeviceOrderValue = repo.FormMe.txt_DeviceOrderLabel.TextValue;
				
				// Compare actualDeviceOrderValue and sDeviceOrderName values and then displaying result
				if(DeviceOrderValue.Equals(sDeviceOrderName))
				{
					Report.Log(ReportLevel.Success, " Device order "+sDeviceOrderName+ " displayed successfully");
				}
				else
				{
					Report.Log(ReportLevel.Failure, " Device order "+sDeviceOrderName+ " not displayed correctly");
				}
				
				// Click on first added EXI800
				Devices_Functions.SelectPointsGridRow(sRowNumber);
				
				// Verify gallery item state
				Devices_Functions.VerifyGalleryItem(sType,sDeviceName,state);

				// Click on Search Properties text field
				repo.ProfileConsys1.txt_SearchProperties.Click();
				
				// Click on cable capacitance cell and then fetch its value
				repo.ProfileConsys1.cell_CableCapacitance.Click();
				string actualCableCapacitance=repo.ProfileConsys1.txt_CableCapacitance.TextValue;
				
				// Compare actual and default cable capacitance value
				if(actualCableCapacitance.Equals(defaultCableCapacitance))
				{
					Report.Log(ReportLevel.Success,"Default cable capacitance displayed as "+actualCableCapacitance);
				}
				
				else
				{
					Report.Log(ReportLevel.Failure,"Cable capacitance displayed as "+actualCableCapacitance+" instead of "+defaultCableCapacitance);
				}
				
				
				// Double click on cable capacitance cell and then enter the cable capacitance value which needs to be changed
				repo.ProfileConsys1.cell_CableCapacitance.DoubleClick();
				repo.ProfileConsys1.txt_CableCapacitance.PressKeys((changedCableCapacitance) +"{ENTER}");

				string actualChangedCableCapacitance=repo.ProfileConsys1.txt_CableCapacitance.TextValue;
				
				// Compare actual and changed cable capacitance value
				if(actualChangedCableCapacitance.Equals(changedCableCapacitance))
				{
					Report.Log(ReportLevel.Success,"Cable capacitance displayed as "+actualChangedCableCapacitance);
				}
				
				else
				{
					Report.Log(ReportLevel.Failure,"Cable capacitance displayed as "+actualChangedCableCapacitance+" instead of "+changedCableCapacitance);
					
				}
				
				
				// Click on first added EXI800
				Devices_Functions.SelectPointsGridRow(sRowNumber);
				
				// Verify gallery item state
				Devices_Functions.VerifyGalleryItem(sType,sDeviceName,changedState);
				
				// Click on Search Properties text field
				repo.ProfileConsys1.txt_SearchProperties.Click();
				
				// Click on cable capacitance cell
				repo.ProfileConsys1.cell_CableCapacitance.Click();
				
				// Enter the cable capacitance value which needs to be changed
				repo.ProfileConsys1.txt_CableCapacitance.PressKeys((defaultCableCapacitance) +"{ENTER}");

				// Click on cable capacitance cell and then fetch its value
				repo.ProfileConsys1.cell_CableCapacitance.Click();
				string CableCapacitance=repo.ProfileConsys1.txt_CableCapacitance.TextValue;
				
				// Compare cable capacitance and default cable capacitance value
				if(CableCapacitance.Equals(defaultCableCapacitance))
				{
					Report.Log(ReportLevel.Success,"Cable capacitance displayed as "+CableCapacitance);
				}
				
				else
				{
					Report.Log(ReportLevel.Failure,"Cable capacitance displayed as "+CableCapacitance+" instead of "+defaultCableCapacitance);
				}
				
				// Click on Points tab
				repo.ProfileConsys1.tab_Points.Click();
				
				// Click on first added EXI800
				Devices_Functions.SelectPointsGridRow(sRowNumber);
				
				// Verify gallery item state
				Devices_Functions.VerifyGalleryItem(sType,sDeviceName,state);
				
				
				// Delete Exi800 devices
				repo.ProfileConsys1.btn_Delete.Click();
				
				ModelNumber = ((Range)Excel_Utilities.ExcelRange.Cells[4,5]).Value.ToString();
				string sType1 = ((Range)Excel_Utilities.ExcelRange.Cells[4,6]).Value.ToString();
				
				// Add EXI800 device
				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType1);
				
			}
			//Close excel
			Excel_Utilities.CloseExcel();
			
		}
		
		/********************************************************************************************************************************
		 * Function Name: VerifyAdditionOfExiDevices
		 * Function Details: To verify addition of EXI devices to its max limit on updating cable capacitance and verify its IS units
		 * Parameter/Arguments:  fileName, sheetName
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 13/02/2019
		 ********************************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyAdditionOfExiDevices(string sFileName,string sAddEXIDeviceSheet, string sAddMaxEXIDeviceSheet)
		{
			//Add devices in loop A,
			Excel_Utilities.OpenExcelFile(sFileName,sAddEXIDeviceSheet);
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			string sType,defaultCableCapacitance,changedState,sRowNumber,changedCableCapacitance;
			
			for(int i=8; i<=rows; i++)
			{
				ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				string state =  ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				
				// Add devices from the gallery
				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
				
				// Verify gallery item state
				Devices_Functions.VerifyGalleryItem(sType,ModelNumber,state);
				
				// Click on Points tab
				repo.ProfileConsys1.tab_Points.Click();
			}
			
			for(int i=8; i<=rows; i++)
			{
				sRowNumber= (i-7).ToString();
				defaultCableCapacitance = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				changedCableCapacitance =  ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				changedState =  ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				
				//Click on Points tab
				repo.ProfileConsys1.tab_Points.Click();
				
				// Click on added EXI800
				Devices_Functions.SelectPointsGridRow(sRowNumber);
				
				// Click on Cable capacitance cell
				repo.ProfileConsys1.cell_CableCapacitance.Click();
				string actualCableCapacitance=repo.ProfileConsys1.txt_CableCapacitance.TextValue;
				
				// Compare actual and default cable capacitance value
				if(actualCableCapacitance.Equals(defaultCableCapacitance))
				{
					Report.Log(ReportLevel.Success,"Default cable capacitance displayed as "+actualCableCapacitance);
				}
				else
				{
					Report.Log(ReportLevel.Failure,"Cable capacitance displayed as "+actualCableCapacitance+" instead of "+defaultCableCapacitance);
				}
				
				// Double click on cable capacitance cell and then enter the value
				repo.ProfileConsys1.cell_CableCapacitance.DoubleClick();
				repo.ProfileConsys1.txt_CableCapacitance.PressKeys((changedCableCapacitance) +"{ENTER}");

				string actualChangedCableCapacitance=repo.ProfileConsys1.txt_CableCapacitance.TextValue;
				
				// Compare actual and changed cable capacitance value
				if(actualChangedCableCapacitance.Equals(changedCableCapacitance))
				{
					Report.Log(ReportLevel.Success,"Cable capacitance displayed as "+actualChangedCableCapacitance);
				}
				else
				{
					Report.Log(ReportLevel.Failure,"Cable capacitance displayed as "+actualChangedCableCapacitance+" instead of "+changedCableCapacitance);
				}
				
				// Click on Points tab
				repo.ProfileConsys1.tab_Points.Click();
				
				// Verify gallery item state
				Devices_Functions.VerifyGalleryItem(sType,ModelNumber,changedState);
			}
			
			//Close excel
			Excel_Utilities.CloseExcel();
			
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddMaxEXIDeviceSheet);
			
			// Count number of rows in excel and store it in rows variable
			rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared string type
			string ExpectedMaxISUnits,ChangedValue,PanelType;
			int noOfDevices = 2;
			PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[4,7]).Value.ToString();
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				noOfDevices++;
				ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				defaultCableCapacitance = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				ChangedValue =  ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				changedState =  ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				ExpectedMaxISUnits = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				
				// Verify panel type and then accordingly assign sRow value
				if(PanelType.Equals("FIM"))
				{
					sRow = (i+1).ToString();
				}
				else
				{
					sRow = (i+2).ToString();
				}
				
				// Click on Points tab
				repo.ProfileConsys1.tab_Points.Click();
				
				// Add devices from the gallery as per test data from the excel sheet
				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
				Report.Log(ReportLevel.Info, "Device "+ModelNumber+" added to Panel");
				
				
				//Click on Physical layout tab
				repo.ProfileConsys1.tab_PhysicalLayout.Click();
				
				//To call verifyMaxISUnits method and verify max IS units value
				Report.Log(ReportLevel.Info,"Verification of Maximum IS Units on addition of devices");
				verifyMaxISUnitsMultipleDevices(ExpectedMaxISUnits, noOfDevices, PanelType);
				
				Report.Log(ReportLevel.Info, "Verified Maximum IS Units.");
				
				//Click on Points tab
				repo.ProfileConsys1.tab_Points.Click();
				
				// Click on Cable capacitance cell
				repo.ProfileConsys1.cell_CableCapacitance.Click();
				string actualCableCapacitance=repo.ProfileConsys1.txt_CableCapacitance.TextValue;
				
				// Compare actual and default cable capacitance value
				if(actualCableCapacitance.Equals(defaultCableCapacitance))
				{
					Report.Log(ReportLevel.Success,"Default cable capacitance displayed as "+actualCableCapacitance);
				}
				
				else
				{
					Report.Log(ReportLevel.Failure,"Cable capacitance displayed as "+actualCableCapacitance+" instead of "+defaultCableCapacitance);
				}
				
				
				// Double click on cell capacitance cell
				repo.ProfileConsys1.cell_CableCapacitance.DoubleClick();
				
				// Change the cable capacitance value
				repo.ProfileConsys1.txt_CableCapacitance.PressKeys((ChangedValue) +"{ENTER}");
				
				//Click on Points tab
				repo.ProfileConsys1.tab_Points.Click();
				
				// Verify gallery item state
				Devices_Functions.VerifyGalleryItem(sType,ModelNumber,changedState);

			}
			//Close opened excel sheet
			Excel_Utilities.CloseExcel();
			
		}

		
		/********************************************************************************************************************************
		 * Function Name: VerifyDragDropOfDevices
		 * Function Details: To verify drag and drop functionality in Physical Layout
		 * Parameter/Arguments:  fileName, sheetName
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 20/02/2019  Alpesh Dhakad - 21/02/2019 - Updated comments and script
		 ********************************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyDragDropOfDevices(string sFileName,string sAddEXIDeviceSheet, string sISDevicesSheet)
		{
			//Add devices in loop A,
			Excel_Utilities.OpenExcelFile(sFileName,sAddEXIDeviceSheet);
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			string sType,sRowNumber;
			for(int i=8; i<=rows; i++)
			{
				ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();

				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
			}
			//Close excel
			Excel_Utilities.CloseExcel();
			
			// Select Point grid and right click on it
			Mouse.Click(repo.FormMe.HeadersPanel1, System.Windows.Forms.MouseButtons.Right);
			
			// Click Show column chooser to select column
			repo.ContextMenu.ShowColumnChooser.Click();
			
			// Click on Device order checkbox to add the column in points grid
			repo.ProfileConsys.chkBox_DeviceOrder1.Click();
			Report.Log(ReportLevel.Info," Device order column added successfully ");
			
			// Close column choose window
			repo.ProfileConsys.btn_CloseColumnChooser.Click();
			
			// Open Another excel to verify IS devices state
			Excel_Utilities.OpenExcelFile(sFileName,sISDevicesSheet);
			
			// Count the number of rows in excel
			rows = Excel_Utilities.ExcelRange.Rows.Count;

			// Add IS devices under first EXI
			for(int i=8; i<=rows; i++)
			{
				sDeviceOrderRow= ((i+1)-6).ToString();
				string sDeviceName = ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				sDeviceOrderName = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sPhysicalLayoutDeviceIndex =  ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				
				sRowNumber= (1).ToString();
				
				// Click on first added EXI800
				Devices_Functions.SelectPointsGridRow(sRowNumber);
				
				// Verify gallery item state
				Devices_Functions.AddDevicesfromGallery(sDeviceName,sType);
				Report.Log(ReportLevel.Info,"Device " +sDeviceName+ " Added to loop");
				
				// Click on Device Name
				repo.FormMe.txt_DeviceOrderLabel.Click();
				
				// Retrieve Device order value
				string actualDeviceOrderValue = repo.FormMe.txt_DeviceOrderLabel.TextValue;
				
				// Compare actualDeviceOrderValue and sDeviceOrderName values and then displaying result
				if(actualDeviceOrderValue.Equals(sDeviceOrderName))
				{
					Report.Log(ReportLevel.Success, "Device " +sDeviceName+ " with " +sDeviceOrderName+ " added successfully and displaying correct device order");
				}
				else
				{
					Report.Log(ReportLevel.Failure, "Device " +sDeviceName+ " with " +sDeviceOrderName+ " not added or not displaying correct device order");
				}
				
				
				// Click on first added EXI800
				Devices_Functions.SelectPointsGridRow(sRowNumber);
			}
			
			// Add IS devices under Second EXI
			for(int j=8; j<=rows; j++)
			{
				sDeviceOrderRow= ((j+4)-6).ToString();
				string sDeviceName = ((Range)Excel_Utilities.ExcelRange.Cells[j,7]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[j,8]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[j,9]).Value.ToString();
				sDeviceOrderName = ((Range)Excel_Utilities.ExcelRange.Cells[j,10]).Value.ToString();
				sRowNumber= (2).ToString();
				
				// Click on first added EXI800
				Devices_Functions.SelectPointsGridRow(sRowNumber);
				
				// Verify gallery item state
				Devices_Functions.AddDevicesfromGallery(sDeviceName,sType);
				Report.Log(ReportLevel.Info,"Device " +sDeviceName+ " Added to loop");
				
				if(repo.FormMe.txt_DeviceOrderLabelInfo.Exists())
				{
					// Click on Device Name
					repo.FormMe.txt_DeviceOrderLabel.Click();
					
					string actualDeviceOrderValue = repo.FormMe.txt_DeviceOrderLabel.TextValue;
					
					// Compare actualDeviceOrderValue and sDeviceOrderName values and then displaying result
					if(actualDeviceOrderValue.Equals(sDeviceOrderName))
					{
						Report.Log(ReportLevel.Success, "Device " +sDeviceName+ " with " +sDeviceOrderName+ " added successfully and displaying correct device order");
					}
					else
					{
						Report.Log(ReportLevel.Failure, "Device " +sDeviceName+ " with " +sDeviceOrderName+ " not added or not displaying correct device order");
					}
				}
				
				// Click on first added EXI800
				Devices_Functions.SelectPointsGridRow(sRowNumber);
			}

			// Click on Physical Layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			// Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			// Click on Physical Layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			// Read Index value of first and second EXI
			string firstEXIIndex = ((Range)Excel_Utilities.ExcelRange.Cells[4,6]).Value.ToString();
			string secondEXIIndex =  ((Range)Excel_Utilities.ExcelRange.Cells[5,6]).Value.ToString();
			
			// Read Physical layout Index value of drop IS device
			sPhysicalLayoutDeviceIndex =  ((Range)Excel_Utilities.ExcelRange.Cells[9,11]).Value.ToString();
			
			// Create a adapter and stored in source adapter element
			Adapter sourceElement = repo.FormMe.lst_PhysicalLayoutDevice;
			
			// Assigning first EXI index value to Physical Layout index
			sPhysicalLayoutDeviceIndex =  firstEXIIndex;
			
			// Create a adapter and stored in targer adapter element
			Adapter targetElement = repo.FormMe.PhysicalLayoutDeviceIndex;

			// Drag and drop IS device from Second EXI to First EXI
			Ranorex.AutomationHelpers.UserCodeCollections.DragNDropLibrary.DragAndDrop(sourceElement,targetElement);
			
			// Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			// Verify device order for IS Devices after Drag and drop action
			for(int k=8; k<=rows; k++)
			{
				sDeviceOrderRow= ((k+4)-6).ToString();
				string sDeviceName = ((Range)Excel_Utilities.ExcelRange.Cells[k,7]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[k,8]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[k,9]).Value.ToString();
				sDeviceOrderName = ((Range)Excel_Utilities.ExcelRange.Cells[k,12]).Value.ToString();
				
				// Click on Device Name
				repo.FormMe.txt_DeviceOrderLabel.Click();
				
				// To retrieve the Device order text value
				string DeviceOrderValue = repo.FormMe.txt_DeviceOrderLabel.TextValue;
				
				// Compare actualDeviceOrderValue and sDeviceOrderName values and then displaying result
				if(DeviceOrderValue.Equals(sDeviceOrderName))
				{
					Report.Log(ReportLevel.Success, " Device order "+sDeviceOrderName+ " displayed correctly");
				}
				else
				{
					Report.Log(ReportLevel.Failure, " Device order "+sDeviceOrderName+ " not displayed correctly");
				}
			}
			//Close excel
			Excel_Utilities.CloseExcel();
		}
		
		/********************************************************************************************************************************
		 * Function Name: VerifyIBUnits
		 * Function Details: To verify IB units value
		 * Parameter/Arguments:  fileName, sheetName
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 22/02/2019
		 ********************************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyIBUnits(string sFileName,string sAddDeviceSheet)
		{
			Excel_Utilities.OpenExcelFile(sFileName,sAddDeviceSheet);
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			string sType, sLabelName,sBaseofDevice,sBasePropertyRowIndex,PanelType,expectedIsolatorUnits;
			for(int i=8; i<=rows; i++)
			{
				ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				sBase = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sRowIndex= ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				sBaseofDevice= ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				sBasePropertyRowIndex= ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				expectedIsolatorUnits= ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				PanelType= ((Range)Excel_Utilities.ExcelRange.Cells[5,5]).Value.ToString();
				
				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
				
				// Verify
				if(sBase.Equals("NA"))
				{
					Report.Log(ReportLevel.Info," No need to change default base ");
				}
				else
				{
					Devices_Functions.AssignDeviceBase(sLabelName,sBase,sRowIndex);
					Report.Log(ReportLevel.Success, "Base " +sBase+ " added successfully");
				}
				
				// Verify
				if(sBaseofDevice.Equals("NA"))
				{
					Report.Log(ReportLevel.Info,"Additional base not required");
				}
				else
				{
					Devices_Functions.AssignAdditionalBase(sLabelName,sBaseofDevice,sBasePropertyRowIndex);
					Report.Log(ReportLevel.Success, "Additional Base " +sBaseofDevice+ " added successfully");
				}
				
				VerifyIsolatorUnits(expectedIsolatorUnits,PanelType);
				
			}
			
			//Close excel
			Excel_Utilities.CloseExcel();
		}
		
		/*****************************************************************************************************************
		 * Function Name: VerifyIsolatorUnits
		 * Function Details: To verify isolator units value
		 * Parameter/Arguments:   expected value, panel type (FIM or PFI)
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 25/02/2019
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyIsolatorUnits(string expectedIsolatorUnits, string PanelType)
		{
			// Verify panel type and then accordingly assign sRow value
			if(PanelType.Equals("FIM"))
			{
				sRow = (7).ToString();
			}
			else
			{
				sRow = (8).ToString();
			}
			

			// Click on Physical layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			// Fetch PSU5V value and store in Actual 5VPSU value
			string ActualIsolatorUnitsValue = repo.FormMe.IsolatorUnits.TextValue;
			
			// Compare Actual and Expected 5V PSU load value
			if(ActualIsolatorUnitsValue.Equals(expectedIsolatorUnits))
			{
				Report.Log(ReportLevel.Success,"Isolator Units value " + ActualIsolatorUnitsValue + " is displayed correctly " );
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Isolator Units value is not displayed correctly, it is displayed as: " + ActualIsolatorUnitsValue + " instead of : " +expectedIsolatorUnits);
			}
			
			// CLick on Points tab
			repo.ProfileConsys1.tab_Points.Click();
		}
		
		
		/*****************************************************************************************************************
		 * Function Name: VerifySelectedIsolatorUnits
		 * Function Details: To verify selected isolator units value
		 * Parameter/Arguments:   expected value, panel type (FIM or PFI)
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 27/02/2019
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifySelectedIsolatorUnits(string expectedSelectedIsolatorUnits, string PanelType)
		{
			// Verify panel type and then accordingly assign sRow value
			if(PanelType.Equals("FIM"))
			{
				sRow = (6).ToString();
			}
			else
			{
				sRow = (7).ToString();
			}
			

			// Click on Physical layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			// Fetch PSU5V value and store in Actual 5VPSU value
			string ActualSelectedIsolatorUnitsValue = repo.FormMe.SelectedIsolatorUnits.TextValue;
			
			// Compare Actual and Expected 5V PSU load value
			if(ActualSelectedIsolatorUnitsValue.Equals(expectedSelectedIsolatorUnits))
			{
				Report.Log(ReportLevel.Success,"Selected Isolator Units value " + ActualSelectedIsolatorUnitsValue + " is displayed correctly " );
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Selected Isolator Units value is not displayed correctly, it is displayed as: " + ActualSelectedIsolatorUnitsValue + " instead of : " +expectedSelectedIsolatorUnits);
			}
			
			// CLick on Points tab
			repo.ProfileConsys1.tab_Points.Click();
		}
		
		/********************************************************************************************************************************
		 * Function Name: VerifyIsolatorUnitsCalculationWithLoopHavingLIDevices
		 * Function Details: To verify isolator unit calculation with loop having line isolator devices
		 * Parameter/Arguments:  fileName, sheetName
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 28/02/2019
		 ********************************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyIsolatorUnitsCalculationWithLoopHavingLIDevices(string sFileName,string sAddDeviceSheet)
		{
			Excel_Utilities.OpenExcelFile(sFileName,sAddDeviceSheet);
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			string PanelType,expectedIsolatorUnits,sDeviceName,sType,sLabelName,IBUnitUntilLI,IBUnitBelowLI;
			int DeviceQty;
			for(int i=8; i<=rows; i++)
			{
				sDeviceName=  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				DeviceQty=  int.Parse(((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString());
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				expectedIsolatorUnits= ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				PanelType= ((Range)Excel_Utilities.ExcelRange.Cells[5,5]).Value.ToString();
				
				Devices_Functions.AddDevicesfromMultiplePointWizard(sDeviceName,DeviceQty);
				
				VerifyIsolatorUnits(expectedIsolatorUnits,PanelType);
			}
			
			Excel_Utilities.CloseExcel();
			
			// Click on Physical Layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			Excel_Utilities.OpenExcelFile(sFileName,sAddDeviceSheet);
			rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			PanelType= ((Range)Excel_Utilities.ExcelRange.Cells[5,5]).Value.ToString();
			
			
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			////////////////////////Below portion of code needs to be removed later after Physical Layout device order issue gets corrected////////////
			/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			
			
			// Read Index value of first device index and drop device index
			string sourceDeviceIndex = ((Range)Excel_Utilities.ExcelRange.Cells[3,11]).Value.ToString();
			string targetDeviceIndex =  ((Range)Excel_Utilities.ExcelRange.Cells[4,11]).Value.ToString();
			
			sPhysicalLayoutDeviceIndex =  sourceDeviceIndex;
			
			// Create a adapter and stored in source adapter element
			Adapter sourceEle = repo.FormMe.PhysicalLayoutDeviceIndex;
			
			// Assigning first EXI index value to Physical Layout index
			sPhysicalLayoutDeviceIndex =  targetDeviceIndex;
			
			// Create a adapter and stored in targer adapter element
			Adapter targetEle = repo.FormMe.PhysicalLayoutDeviceIndex;


			// Drag and drop EXI or LI device from First position to its defined position
			Ranorex.AutomationHelpers.UserCodeCollections.DragNDropLibrary.DragAndDrop(sourceEle,targetEle);
			
			
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			////////////////////////Above portion of code needs to be removed later after Physical Layout device order issue gets corrected////////////
			/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

			//Summation of IB units until  Line Isolator
			IBUnitUntilLI =  ((Range)Excel_Utilities.ExcelRange.Cells[4,7]).Value.ToString();
			sPhysicalLayoutDeviceIndex =  ((Range)Excel_Utilities.ExcelRange.Cells[4,8]).Value.ToString();

			repo.FormMe.PhysicalLayoutDeviceIndex.Click();
			
			VerifySelectedIsolatorUnits(IBUnitUntilLI,PanelType);
			
			//Summation of IB units below  Line Isolator
			IBUnitBelowLI =  ((Range)Excel_Utilities.ExcelRange.Cells[5,7]).Value.ToString();
			sPhysicalLayoutDeviceIndex =  ((Range)Excel_Utilities.ExcelRange.Cells[5,8]).Value.ToString();
			
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			repo.FormMe.PhysicalLayoutDeviceIndex.Click();
			
			VerifySelectedIsolatorUnits(IBUnitBelowLI,PanelType);
			
			// CLick on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			Excel_Utilities.CloseExcel();
			
		}
		
		/********************************************************************************************************************************
		 * Function Name: VerifyIsolatorUnitsCalculationForDevicesInsideLI
		 * Function Details: To verify isolator unit calculation for devices added inside line isolator devices
		 * Parameter/Arguments:  fileName, sheetName
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 04/03/2019
		 ********************************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyIsolatorUnitsCalculationForDevicesInsideLI(string sFileName,string sAddDeviceSheet,string sVerifyIBDeviceSheet)
		{
			Excel_Utilities.OpenExcelFile(sFileName,sAddDeviceSheet);
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			string PanelType,expectedIsolatorUnits,sDeviceName,sType,sLabelName,IBUnitUntilLI,IBUnitUntilBuiltInLI,IBUnitInsideLI,IBUnitBelowLI,sourceDeviceIndex,targetDeviceIndex;
			int DeviceQty;
			for(int i=8; i<=rows; i++)
			{
				sDeviceName=  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				DeviceQty=  int.Parse(((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString());
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				expectedIsolatorUnits= ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				PanelType= ((Range)Excel_Utilities.ExcelRange.Cells[5,5]).Value.ToString();
				
				Devices_Functions.AddDevicesfromMultiplePointWizard(sDeviceName,DeviceQty);
				
				//VerifyIsolatorUnits(expectedIsolatorUnits,PanelType);
			}
			
			Excel_Utilities.CloseExcel();

			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			////////////////////////Below portion of code needs to be removed later after Physical Layout device order issue gets corrected////////////
			/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			
			// Click on Physical Layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			// Click on Physical Layout tab
			repo.ProfileConsys1.tab_Points.Click();
			
			// Click on Physical Layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			Excel_Utilities.OpenExcelFile(sFileName,sAddDeviceSheet);
			rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			PanelType= ((Range)Excel_Utilities.ExcelRange.Cells[5,5]).Value.ToString();

			// Read Index value of first device index and drop device index
			sourceDeviceIndex = ((Range)Excel_Utilities.ExcelRange.Cells[3,8]).Value.ToString();
			targetDeviceIndex =  ((Range)Excel_Utilities.ExcelRange.Cells[4,8]).Value.ToString();
			
			sPhysicalLayoutDeviceIndex =  sourceDeviceIndex;
			
			// Create a adapter and stored in source adapter element
			Adapter sourceElement1 = repo.FormMe.lst_PhysicalLayoutDevice;
			
			// Assigning first EXI index value to Physical Layout index
			sPhysicalLayoutDeviceIndex =  targetDeviceIndex;
			
			// Create a adapter and stored in targer adapter element
			Adapter targetElement1 = repo.FormMe.PhysicalLayoutDevice;


			// Drag and drop EXI or LI device from First position to its defined position
			Ranorex.AutomationHelpers.UserCodeCollections.DragNDropLibrary.DragAndDrop(sourceElement1,targetElement1);
			
			// Read Index value of first device index and drop device index
			string sourceDeviceIndex1 = ((Range)Excel_Utilities.ExcelRange.Cells[3,10]).Value.ToString();
			string targetDeviceIndex1 =  ((Range)Excel_Utilities.ExcelRange.Cells[4,10]).Value.ToString();
			
			sPhysicalLayoutDeviceIndex =  sourceDeviceIndex1;
			
			// Create a adapter and stored in source adapter element
			Adapter sourceElement2 = repo.FormMe.lst_PhysicalLayoutDevice;
			
			// Assigning first EXI index value to Physical Layout index
			sPhysicalLayoutDeviceIndex =  targetDeviceIndex1;
			
			// Create a adapter and stored in targer adapter element
			Adapter targetElement2 = repo.FormMe.PhysicalLayoutDevice;

			// Drag and drop EXI or LI device from First position to its defined position
			Ranorex.AutomationHelpers.UserCodeCollections.DragNDropLibrary.DragAndDrop(sourceElement2,targetElement2);
			
			Excel_Utilities.CloseExcel();
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			////////////////////////Above portion of code needs to be removed later after Physical Layout device order issue gets corrected////////////
			/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

			// Click on Physical Layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			Excel_Utilities.OpenExcelFile(sFileName,sVerifyIBDeviceSheet);
			rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			PanelType= ((Range)Excel_Utilities.ExcelRange.Cells[5,5]).Value.ToString();
			
			
			for(int i=8; i<=rows; i++)
			{
				sDeviceName=  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType= ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sourceDeviceIndex= ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				targetDeviceIndex= ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				
				// Assigning first EXI index value to Physical Layout index
				sPhysicalLayoutDeviceIndex =  sourceDeviceIndex;
				
				// Create a adapter and stored in source adapter element
				Adapter sourceElement = repo.FormMe.lst_PhysicalLayoutDevice;
				
				// Assigning first EXI index value to Physical Layout index
				sPhysicalLayoutDeviceIndex =  targetDeviceIndex;
				
				// Create a adapter and stored in targer adapter element
				Adapter targetElement = repo.FormMe.PhysicalLayoutDevice;

				// Drag and drop EXI or LI device from First position to its defined position
				Ranorex.AutomationHelpers.UserCodeCollections.DragNDropLibrary.DragAndDrop(sourceElement,targetElement);
				
			}
			
			repo.ProfileConsys1.tab_Points.Click();
			
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			
			//Summation of IB units until  Line Isolator
			IBUnitUntilLI =  ((Range)Excel_Utilities.ExcelRange.Cells[2,7]).Value.ToString();
			sPhysicalLayoutDeviceIndex =  ((Range)Excel_Utilities.ExcelRange.Cells[2,8]).Value.ToString();

			repo.FormMe.PhysicalLayoutDevice.Click();
			
			VerifySelectedIsolatorUnits(IBUnitUntilLI,PanelType);
			
			//Summation of IB units until built in  Line Isolator
			IBUnitUntilBuiltInLI = ((Range)Excel_Utilities.ExcelRange.Cells[3,7]).Value.ToString();
			sPhysicalLayoutDeviceIndex =  ((Range)Excel_Utilities.ExcelRange.Cells[3,8]).Value.ToString();
			
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			repo.FormMe.PhysicalLayoutDevice.Click();
			
			VerifySelectedIsolatorUnits(IBUnitUntilBuiltInLI,PanelType);
			
			
			//Summation of IB units inside built in  Line Isolator
			IBUnitInsideLI = ((Range)Excel_Utilities.ExcelRange.Cells[4,7]).Value.ToString();
			sPhysicalLayoutDeviceIndex =  ((Range)Excel_Utilities.ExcelRange.Cells[4,8]).Value.ToString();
			
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			repo.FormMe.PhysicalLayoutDevice.Click();
			
			VerifySelectedIsolatorUnits(IBUnitInsideLI,PanelType);
			
			
			//Summation of IB units below built in  Line Isolator
			IBUnitBelowLI = ((Range)Excel_Utilities.ExcelRange.Cells[5,7]).Value.ToString();
			sPhysicalLayoutDeviceIndex =  ((Range)Excel_Utilities.ExcelRange.Cells[5,8]).Value.ToString();
			
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			repo.FormMe.PhysicalLayoutDevice.Click();
			
			VerifySelectedIsolatorUnits(IBUnitBelowLI,PanelType);

			// CLick on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			Excel_Utilities.CloseExcel();
			
		}
		
		/********************************************************************************************************************************
		 * Function Name: VerifyIsolatorUnitsCalculationForDevicesWithBuiltInIsolator
		 * Function Details: to Verify Isolator units calculation for devices having built in isolator
		 * Parameter/Arguments:  fileName, sheetNames
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 05/03/2019
		 ********************************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyIsolatorUnitsCalculationForDevicesWithBuiltInIsolator(string sFileName,string sAddDeviceSheet,string sVerifyIBDeviceSheet)
		{
			Excel_Utilities.OpenExcelFile(sFileName,sAddDeviceSheet);
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			string PanelType,expectedIsolatorUnits,sDeviceName,sType,sLabelName,IBUnitUntilLI,IBUnitUntilBuiltPtN,IBUnitUntilLIBelow,IBUnitBetPtoNLI,IBUnitUntilLevelLI,IBUnitBetPtoNBuiltLI,IBUnitBetPtoNI,sourceDeviceIndex,targetDeviceIndex;
			int DeviceQty;
			for(int i=8; i<=rows; i++)
			{
				sDeviceName=  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				DeviceQty=  int.Parse(((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString());
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				expectedIsolatorUnits= ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				PanelType= ((Range)Excel_Utilities.ExcelRange.Cells[5,5]).Value.ToString();
				
				Devices_Functions.AddDevicesfromGallery(sDeviceName,sType);
				
				//Devices_Functions.AddDevicesfromMultiplePointWizard(sDeviceName,DeviceQty);
				
				//VerifyIsolatorUnits(expectedIsolatorUnits,PanelType);
			}
			
			Excel_Utilities.CloseExcel();
			
			
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			////////////////////////Below portion of code needs to be removed later after Physical Layout device order issue gets corrected////////////
			/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			
			// Click on Physical Layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			// Click on Physical Layout tab
			repo.ProfileConsys1.tab_Points.Click();
			
			// Click on Physical Layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			Excel_Utilities.OpenExcelFile(sFileName,sAddDeviceSheet);
			rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			PanelType= ((Range)Excel_Utilities.ExcelRange.Cells[5,5]).Value.ToString();

			// Read Index value of first device index and drop device index
			sourceDeviceIndex = ((Range)Excel_Utilities.ExcelRange.Cells[3,8]).Value.ToString();
			targetDeviceIndex =  ((Range)Excel_Utilities.ExcelRange.Cells[4,8]).Value.ToString();
			
			sPhysicalLayoutDeviceIndex =  sourceDeviceIndex;
			
			// Create a adapter and stored in source adapter element
			Adapter sourceElement = repo.FormMe.lst_PhysicalLayoutDevice;
			
			// Assigning first EXI index value to Physical Layout index
			sPhysicalLayoutDeviceIndex =  targetDeviceIndex;
			
			// Create a adapter and stored in targer adapter element
			Adapter targetElement = repo.FormMe.PhysicalLayoutDevice;


			// Drag and drop EXI or LI device from First position to its defined position
			Ranorex.AutomationHelpers.UserCodeCollections.DragNDropLibrary.DragAndDrop(sourceElement,targetElement);
			
			
			Excel_Utilities.CloseExcel();
			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			////////////////////////Above portion of code needs to be removed later after Physical Layout device order issue gets corrected////////////
			/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


			// Click on Physical Layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			Excel_Utilities.OpenExcelFile(sFileName,sVerifyIBDeviceSheet);
			rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			PanelType= ((Range)Excel_Utilities.ExcelRange.Cells[5,5]).Value.ToString();
			
			
			for(int i=8; i<=rows; i++)
			{
				sDeviceName=  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType= ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sourceDeviceIndex= ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				targetDeviceIndex= ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				
				// Assigning first EXI index value to Physical Layout index
				sPhysicalLayoutDeviceIndex =  sourceDeviceIndex;
				
				// Create a adapter and stored in source adapter element
				Adapter sourceE = repo.FormMe.lst_PhysicalLayoutDevice;
				
				// Assigning first EXI index value to Physical Layout index
				sPhysicalLayoutDeviceIndex =  targetDeviceIndex;
				
				// Create a adapter and stored in targer adapter element
				Adapter targetE = repo.FormMe.PhysicalLayoutDevice;

				// Drag and drop EXI or LI device from First position to its defined position
				Ranorex.AutomationHelpers.UserCodeCollections.DragNDropLibrary.DragAndDrop(sourceE,targetE);
				
			}
			
			repo.ProfileConsys1.tab_Points.Click();
			
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			//Summation of IB units until  Line Isolator
			IBUnitUntilLI =  ((Range)Excel_Utilities.ExcelRange.Cells[2,7]).Value.ToString();
			sPhysicalLayoutDeviceIndex =  ((Range)Excel_Utilities.ExcelRange.Cells[2,8]).Value.ToString();

			repo.FormMe.PhysicalLayoutDevice.Click();
			
			VerifySelectedIsolatorUnits(IBUnitUntilLI,PanelType);
			
			//Summation of IB units from previous to next built in  Line Isolator
			IBUnitUntilBuiltPtN = ((Range)Excel_Utilities.ExcelRange.Cells[3,7]).Value.ToString();
			sPhysicalLayoutDeviceIndex =  ((Range)Excel_Utilities.ExcelRange.Cells[3,8]).Value.ToString();
			
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			repo.FormMe.PhysicalLayoutDevice.Click();
			
			VerifySelectedIsolatorUnits(IBUnitUntilBuiltPtN,PanelType);
			
			
			//Summation of IB units until we get LI/built in  Line Isolator
			IBUnitUntilLIBelow = ((Range)Excel_Utilities.ExcelRange.Cells[4,7]).Value.ToString();
			sPhysicalLayoutDeviceIndex =  ((Range)Excel_Utilities.ExcelRange.Cells[4,8]).Value.ToString();
			
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			repo.FormMe.PhysicalLayoutDevice.Click();
			
			VerifySelectedIsolatorUnits(IBUnitUntilLIBelow,PanelType);
			
			
			//Summation of IB units between previous to next built in  Line Isolator
			IBUnitBetPtoNLI = ((Range)Excel_Utilities.ExcelRange.Cells[5,7]).Value.ToString();
			sPhysicalLayoutDeviceIndex =  ((Range)Excel_Utilities.ExcelRange.Cells[5,8]).Value.ToString();
			
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			repo.FormMe.PhysicalLayoutDevice.Click();
			
			VerifySelectedIsolatorUnits(IBUnitBetPtoNLI,PanelType);
			
			//Summation of IB units present at the level Line isolator until we get the Line Isolator/Built in isolator device
			IBUnitUntilLevelLI = ((Range)Excel_Utilities.ExcelRange.Cells[6,7]).Value.ToString();
			sPhysicalLayoutDeviceIndex =  ((Range)Excel_Utilities.ExcelRange.Cells[6,8]).Value.ToString();
			
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			repo.FormMe.PhysicalLayoutDevice.Click();
			
			VerifySelectedIsolatorUnits(IBUnitUntilLevelLI,PanelType);
			
			//Summation of IB units present at the level Line isolator until we get the Line Isolator/Built in isolator device
			IBUnitBetPtoNBuiltLI = ((Range)Excel_Utilities.ExcelRange.Cells[7,7]).Value.ToString();
			sPhysicalLayoutDeviceIndex =  ((Range)Excel_Utilities.ExcelRange.Cells[7,8]).Value.ToString();
			
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			repo.FormMe.PhysicalLayoutDevice.Click();
			
			VerifySelectedIsolatorUnits(IBUnitBetPtoNBuiltLI,PanelType);
			
			//Summation of IB units present at the level Line isolator until we get the Line Isolator/Built in isolator device
			IBUnitBetPtoNI = ((Range)Excel_Utilities.ExcelRange.Cells[8,7]).Value.ToString();
			sPhysicalLayoutDeviceIndex =  ((Range)Excel_Utilities.ExcelRange.Cells[8,8]).Value.ToString();
			
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			repo.FormMe.PhysicalLayoutDevice.Click();
			
			VerifySelectedIsolatorUnits(IBUnitBetPtoNI,PanelType);
			
			// CLick on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			Excel_Utilities.CloseExcel();
			
		}
		

		/********************************************************************
		 * Function Name:VerifyCableCapacitanceOnReopen
		 * Function Details: To verify IS calculation of IS devices
		 * Parameter/Arguments:  fileName, sheetName
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyCableCapacitanceOnReopen(string sFileName,string sAddDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared string type
			string sType, sLabelName,ExpectedMaxISUnits,CableCapacitanceChangedValue,PanelType,sRowNumber;
			int noOfDevices = 0;
			
			PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[4,7]).Value.ToString();
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				noOfDevices++;
				ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				ExpectedMaxISUnits = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				sISUnits = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				CableCapacitanceChangedValue =  ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				
				//Click on Points tab
				repo.ProfileConsys1.tab_Points.Click();
				
				sRowNumber = (i-7).ToString();
				
				
				Devices_Functions.SelectPointsGridRow(sRowNumber);
				
				repo.FormMe.cell_CableCapacitanceAfterReopen.Click();
				
				string CableCapacitance=repo.FormMe.txt_CableCapcitanceAfterReopen.TextValue;
				
				if(CableCapacitance.Equals(CableCapacitanceChangedValue))
				{
					Report.Log(ReportLevel.Success,"Cable capacitance displayed as "+CableCapacitance+ " and data is peristed after reopen");
				}
				
				else
				{
					Report.Log(ReportLevel.Failure,"Cable capacitance displayed as "+CableCapacitance+" instead of "+CableCapacitanceChangedValue+ " and data not persisted");
				}
				
				
				
			}
			//Close opened excel sheet
			Excel_Utilities.CloseExcel();
			
		}
		
		/********************************************************************
		 * Function Name: VerifyDeviceOrder
		 * Function Details: To verify Device order
		 * Parameter/Arguments: deviceOrder
		 * Output:
		 * Function Owner: Poonam kadam
		 * Last Update : 09/4/19
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyDeviceOrder(string sRow, string sDeviceOrderName)
		{
			sDeviceOrderRow=sRow;
			// Select Point grid and right click on it
			Mouse.Click(repo.FormMe.HeadersPanel1, System.Windows.Forms.MouseButtons.Right);
			
			// Click Show column chooser to select column
			repo.ContextMenu.ShowColumnChooser.Click();
			
			// Click on Device order checkbox to add the column in points grid
			repo.ProfileConsys.chkBox_DeviceOrder1.Click();
			Report.Log(ReportLevel.Info," Device order column added successfully ");
			
			// Close column choose window
			repo.ProfileConsys.btn_CloseColumnChooser.Click();

			// Click on Device Name
			repo.FormMe.txt_DeviceOrderLabel.Click();
			
			string actualDeviceOrderValue = repo.FormMe.txt_DeviceOrderLabel.TextValue;
			Report.Log(ReportLevel.Info,"Actual Device Order "+actualDeviceOrderValue+ " Expected Device Order "+sDeviceOrderName);
			// Compare actualDeviceOrderValue and sDeviceOrderName values and then displaying result
			if(actualDeviceOrderValue.Equals(sDeviceOrderName))
			{
				Report.Log(ReportLevel.Success, "Device with " +sDeviceOrderName+ " added successfully and displaying correct device order");
			}
			else
			{
				Report.Log(ReportLevel.Failure, "Device with " +sDeviceOrderName+ " not added or not displaying correct device order");
			}
			
		}
		
		/********************************************************************
		 * Function Name: ChangeCableCapacitance
		 * Function Details: To Change cable capacitance of the
		 * Parameter/Arguments: cableCapacitanceValue
		 * Output:
		 * Function Owner: Poonam kadam
		 * Last Update : 10/4/19
		 ********************************************************************/
		[UserCodeMethod]
		public static void ChangeCableCapacitance(string cableCapacitanceValue)
		{
			// Double click on cell capacitance cell
			repo.ProfileConsys1.cell_CableCapacitance.DoubleClick();
			
			// Change the cable capacitance value
			repo.ProfileConsys1.txt_CableCapacitance.PressKeys((cableCapacitanceValue) +"{ENTER}");
			Report.Log(ReportLevel.Info, "Cable capcitance value changed to"+cableCapacitanceValue);
		}
		
		/********************************************************************
		 * Function Name: VerifyIsolatorUnitIndicator
		 * Function Details: Verify indicator for isolator units progress bar
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Poonam kadam
		 * Last Update : 12/4/19
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyIsolatorUnitsAndIndicator(string expectedIsolatorUnits, string sExpectedColourCode, string PanelType)
		{
			// Verify panel type and then accordingly assign sRow value
			if(PanelType.Equals("FIM"))
			{
				sRowIndex = (7).ToString();
			}
			else
			{
				sRowIndex = (8).ToString();
			}
//
//			// Fetch Isolator units
//			string ActualIsolatorUnitsValue = repo.FormMe.IsolatorUnits.TextValue;
//
//			// Compare Actual and Expected 5V PSU load value
//			if(ActualIsolatorUnitsValue.Equals(expectedIsolatorUnits))
//			{
//				Report.Log(ReportLevel.Success,"Isolator Units value " + ActualIsolatorUnitsValue + " is displayed correctly " );
//			}
//			else
//			{
//				Report.Log(ReportLevel.Failure,"Isolator Units value is not displayed correctly, it is displayed as: " + ActualIsolatorUnitsValue + " instead of : " +expectedIsolatorUnits);
//			}
			
			//Verify Progress bar color code for extra IS units
			string sActualColorCode=repo.FormMe.IsolatorUnitsProgressBar.GetAttributeValue<string>("foreground");
			Report.Log(ReportLevel.Info,"Actual color is"+sActualColorCode);
			Devices_Functions.VerifyPercentage(sExpectedColourCode, sActualColorCode);
			Report.Log(ReportLevel.Info,"Progress bar color code for Isolator units verified as: "+sExpectedColourCode);
			
		}
		
	}
}


