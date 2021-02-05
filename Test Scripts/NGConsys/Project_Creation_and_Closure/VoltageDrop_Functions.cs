/*
 * Created by Ranorex
 * User: jdhakaa
 * Date: 11/22/2018
 * Time: 9:25 PM
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
	public class VoltageDrop_Functions
	{
		//Create instance of repository to access repository items
		static NGConsysRepository repo = NGConsysRepository.Instance;
		
		static string ModelNumber
		{
			
			get { return repo.ModelNumber; }
			set { repo.ModelNumber = value; }
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
		
		static string sGalleryIndex
		{
			get { return repo.sGalleryIndex; }
			set { repo.sGalleryIndex = value; }
		}
		
		static string sACUnits
		{
			get { return repo.sACUnits; }
			set { repo.sACUnits = value; }
		}
		
		static string sMaxACUnits
		{
			get { return repo.sMaxACUnits; }
			set { repo.sMaxACUnits = value; }
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
		
		static string sDeviceSensitivity
		{
			get { return repo.sDeviceSensitivity; }
			set { repo.sDeviceSensitivity = value; }
		}
		
		static string sDeviceMode
		{
			get { return repo.sDeviceMode; }
			set { repo.sDeviceMode = value; }
		}
		
		static string sDayMode
		{
			get { return repo.sDayMode; }
			set { repo.sDayMode = value; }
		}
		
		static string sDaySensitivity
		{
			get { return repo.sDaySensitivity; }
			set { repo.sDaySensitivity = value; }
		}
		
		static string sColumn
		{
			get { return repo.sColumn; }
			set { repo.sColumn = value; }
		}
		
		/*****************************************************************************************************************
		 * Function Name: verifyVoltageDropOnAddingAndRemovingDevices
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner:
		 * Last Update : Alpesh Dhakad - 30/07/2019 and 21/08/2019 - Updated script as per new build and xpath
		 * Alpesh Dhakad - 04/12/2019 - Updated test scripts with new method for loading details
		 * Alpesh Dhakad - 21/01/2021 Updated script as per new UI implementation
		 *****************************************************************************************************************/
		// To verify voltage drop value on adding and removing devices
		[UserCodeMethod]
		public static void verifyVoltageDropOnAddingAndRemovingDevices(string sFileName,string sAddDevicesLoopA, string sDeleteDevicesLoopA)
		{
			// Open the excel file with mentioned name in argument and sheet
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesLoopA);
			
			// Count the number of rows in excel
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared various fields as String type
			string Multichannel,sLabelName,expectedDCUnits,expectedVoltDrop,expectedVoltDropWorstcase,expectedMaxVoltDrop,expectedMaxVoltDropWorstcase,sType;
			string LoadingDetailsNameForVoltDrop,LoadingDetailsNameForVoltDropWorstCase,LoadingDetailsNameForMaxVoltDrop,LoadingDetailsNameForMaxVoltDropWorstCase;
			string LoopA_Details,LoopB_Details,LoopC_Details,sColumnDCUnit,sColumnVoltDrop,sColumnVoltDropWorstCase;
			
			
			LoopA_Details=((Range)Excel_Utilities.ExcelRange.Cells[2,10]).Value.ToString();
			LoopB_Details=((Range)Excel_Utilities.ExcelRange.Cells[3,10]).Value.ToString();
			LoopC_Details=((Range)Excel_Utilities.ExcelRange.Cells[4,10]).Value.ToString();
			sColumnDCUnit=((Range)Excel_Utilities.ExcelRange.Cells[2,11]).Value.ToString();
			sColumnVoltDrop=((Range)Excel_Utilities.ExcelRange.Cells[3,11]).Value.ToString();
			sColumnVoltDropWorstCase=((Range)Excel_Utilities.ExcelRange.Cells[4,11]).Value.ToString();
	
			
			// For loop to fetch values from the excel sheet and then add devices
			for(int i=6; i<=rows; i++)
			{
				ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				Multichannel = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				expectedDCUnits= ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				expectedVoltDrop = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				expectedVoltDropWorstcase = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				LoadingDetailsNameForVoltDrop = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				LoadingDetailsNameForVoltDropWorstCase = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				
				
				// Add devices from the gallery as per test data from the excel sheet
				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
				
				//Go to Physical layout
				Common_Functions.clickOnPhysicalLayoutTab();
				
				
				Common_Functions.clickOnPanelCalculationsTab();
				
				sColumn = sColumnDCUnit;
				// Call verifyDCUnitValue method and Verify DC units value
				//DC_Functions.verifyDCUnitsValue(expectedDCUnits);
				//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,"Current (DC Units)");
				Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopA_Details,sColumn);
			
				sColumn = sColumnVoltDrop;
				// Call verifyVoltDropValue method and Verify VoltDrop value value
				//verifyVoltDropValue(expectedVoltDrop);
				//Devices_Functions.verifyLoadingDetailsValue(expectedVoltDrop,LoadingDetailsNameForVoltDrop);
				Devices_Functions.verifyLoopLoadingDetailsValue(expectedVoltDrop,LoopA_Details,sColumn);
			
				
				sColumn = sColumnVoltDropWorstCase;
				// Call verifyVoltDropWorstCaseValue method and Verify VoltDrop Worst case value value
				//verifyVoltDropWorstCaseValue(expectedVoltDropWorstcase);
				//Devices_Functions.verifyLoadingDetailsValue(expectedVoltDropWorstcase,LoadingDetailsNameForVoltDropWorstCase);
				Devices_Functions.verifyLoopLoadingDetailsValue(expectedVoltDropWorstcase,LoopA_Details,sColumn);
			
				Common_Functions.clickOnPropertiesTab();
				
				//Click on Points tab
				Common_Functions.clickOnPointsTab();
				
				
			}
			
			// Fetch value from excel sheet and store it
			expectedMaxVoltDrop = ((Range)Excel_Utilities.ExcelRange.Cells[3,4]).Value.ToString();
			expectedMaxVoltDropWorstcase = ((Range)Excel_Utilities.ExcelRange.Cells[4,4]).Value.ToString();
			LoadingDetailsNameForMaxVoltDrop = ((Range)Excel_Utilities.ExcelRange.Cells[3,5]).Value.ToString();
			LoadingDetailsNameForMaxVoltDropWorstCase = ((Range)Excel_Utilities.ExcelRange.Cells[4,5]).Value.ToString();
				
			
			sColumn = sColumnVoltDrop;
			
			Common_Functions.clickOnPanelCalculationsTab();
				
			// Call verifyMaxVoltDrop method and Verify Max volt drop value
			//verifyMaxVoltDrop(expectedMaxVoltDrop);
			//Devices_Functions.verifyMaxLoadingDetailsValue(expectedMaxVoltDrop,LoadingDetailsNameForMaxVoltDrop);
			Devices_Functions.verifyMaxLoopLoadingDetailsValue(expectedMaxVoltDrop,LoopA_Details,sColumn);
			
			sColumn = sColumnVoltDropWorstCase;
			
			// Call verifyMaxVoltDropWorstCaseValue method and Verify Max volt drop worst case value
			//verifyMaxVoltDropWorstCaseValue(expectedMaxVoltDropWorstcase);
			//Devices_Functions.verifyMaxLoadingDetailsValue(expectedMaxVoltDropWorstcase,LoadingDetailsNameForMaxVoltDropWorstCase);
			Devices_Functions.verifyMaxLoopLoadingDetailsValue(expectedMaxVoltDropWorstcase,LoopA_Details,sColumn);
			
			// Close the currently opened excel sheet
			Excel_Utilities.CloseExcel();
			
			//Delete Devices from loop A
			
			// Click on Loop A node
			Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
			//Click on Points tab
			Common_Functions.clickOnPointsTab();
			
			// Open the excel file with mentioned name in argument and sheet
			Excel_Utilities.OpenExcelFile(sFileName,sDeleteDevicesLoopA);
			
			// Count the number of rows in excel
			rows = Excel_Utilities.ExcelRange.Rows.Count;
			
			// For loop to fetch values from the excel sheet and then delete devices using label
			for(int i=9;i<=rows;i++)
			{
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				
				Devices_Functions.DeleteDeviceUsingLabel(sLabelName);
			}
			
			// Fetch value from excel sheet and store it
			expectedVoltDrop = ((Range)Excel_Utilities.ExcelRange.Cells[7,2]).Value.ToString();
			expectedVoltDropWorstcase = ((Range)Excel_Utilities.ExcelRange.Cells[7,4]).Value.ToString();
			LoadingDetailsNameForVoltDrop = ((Range)Excel_Utilities.ExcelRange.Cells[7,1]).Value.ToString();
			LoadingDetailsNameForVoltDropWorstCase = ((Range)Excel_Utilities.ExcelRange.Cells[7,3]).Value.ToString();
				
			
			//Click on Physical Layout tab
			Common_Functions.clickOnPhysicalLayoutTab();
			
			Common_Functions.clickOnPanelCalculationsTab();
			
			sColumn = sColumnVoltDrop;
			// Call verifyVoltDropValue method and Verify volt drop value
			//verifyVoltDropValue(expectedVoltDrop);
			//Devices_Functions.verifyLoadingDetailsValue(expectedVoltDrop,LoadingDetailsNameForVoltDrop);
			Devices_Functions.verifyLoopLoadingDetailsValue(expectedVoltDrop,LoopA_Details,sColumn);
			
			sColumn = sColumnVoltDropWorstCase;
			// Call verifyVoltDropWorstCaseValue method and Verify volt drop worst case value
			//verifyVoltDropWorstCaseValue(expectedVoltDropWorstcase);
			//Devices_Functions.verifyLoadingDetailsValue(expectedVoltDropWorstcase,LoadingDetailsNameForVoltDropWorstCase);
			Devices_Functions.verifyLoopLoadingDetailsValue(expectedVoltDropWorstcase,LoopA_Details,sColumn);
			
			Common_Functions.clickOnPropertiesTab();
				
		}
		
		/********************************************************************
		 * Function Name:
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner:
		 * Last Update :
		 ********************************************************************/
		// Verify max volt Drop value
		[UserCodeMethod]
		public static void verifyMaxVoltDrop(string expectedVoltDropMaxValue)
		{
			//Click on Physical Layout tab
			Common_Functions.clickOnPhysicalLayoutTab();
			
			//Fetch max volt drop text value and storing it in string
			string maxVoltDropValue = repo.ProfileConsys1.MaxVoltDropValue.TextValue;
			
			//Comparing expected and actual maxVoltDrop value
			if(maxVoltDropValue.Equals(expectedVoltDropMaxValue))
			{
				Report.Log(ReportLevel.Success,"Max Volt Drop value are displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Max Volt Drop value are not displayed correctly");
			}
			
			//Click on Points tab
			Common_Functions.clickOnPointsTab();
			
			
		}
		
		/********************************************************************
		 * Function Name:
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner:
		 * Last Update :
		 ********************************************************************/
		// Verify volt Drop value
		[UserCodeMethod]
		public static void verifyVoltDropValue(string expectedVoltDropValue)
		{
			//Click on Physical Layout tab
			Common_Functions.clickOnPhysicalLayoutTab();
			
			//Fetch volt drop text value and storing it in string
			string ActualVoltDropValue = repo.ProfileConsys1.VoltDropValue.TextValue;
			
			//Comparing expected and actual VoltDrop value
			if(ActualVoltDropValue.Equals(expectedVoltDropValue))
			{
				Report.Log(ReportLevel.Success,"Volt Drop Value " + ActualVoltDropValue + " displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Volt Drop Value are not displayed correctly, Volt Drop Value displayed as: " +ActualVoltDropValue +" instead of : "+expectedVoltDropValue);
			}
			
			//Click on Points tab
			Common_Functions.clickOnPointsTab();
			
			
		}
		
		/********************************************************************
		 * Function Name:
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner:
		 * Last Update :
		 ********************************************************************/
		// Verify max volt Drop worst case value
		[UserCodeMethod]
		public static void verifyMaxVoltDropWorstCaseValue(string expectedVoltDropWorstCaseMaxValue)
		{
			//Click on Physical Layout tab
			Common_Functions.clickOnPhysicalLayoutTab();
			
			//Fetch max volt drop worst case text value and storing it in string
			string maxVoltDropWorstCaseValue = repo.ProfileConsys1.MaxVoltDropWorstCaseValue.TextValue;
			
			//Comparing expected and actual maxVoltDrop Worst case value
			if(maxVoltDropWorstCaseValue.Equals(expectedVoltDropWorstCaseMaxValue))
			{
				Report.Log(ReportLevel.Success,"Max Volt Drop worst case value are displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Max Volt Drop worst case value are not displayed correctly");
			}
			
			//Click on Points tab
			Common_Functions.clickOnPointsTab();
			
		}
		
		/********************************************************************
		 * Function Name:
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner:
		 * Last Update :
		 ********************************************************************/
		// Verify volt Drop worst case value
		[UserCodeMethod]
		public static void verifyVoltDropWorstCaseValue(string expectedVoltDropWorstCaseValue)
		{
			//Click on Physical Layout tab
			Common_Functions.clickOnPhysicalLayoutTab();
			
			//Fetch volt drop worst case text value and storing it in string
			string ActualVoltDropWorstCaseValue = repo.ProfileConsys1.VoltDropWorstCaseValue.TextValue;
			
			//Comparing expected and actual VoltDrop Worst case value
			if(ActualVoltDropWorstCaseValue.Equals(expectedVoltDropWorstCaseValue))
			{
				Report.Log(ReportLevel.Success,"Volt Drop Worst case Value " + ActualVoltDropWorstCaseValue + " displayed correclty");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Volt Drop Worst case Value are not displayed correctly, Volt Drop worst case Value displayed as: " +ActualVoltDropWorstCaseValue +" instead of : "+expectedVoltDropWorstCaseValue);
			}
			
			//Click on Points tab
			Common_Functions.clickOnPointsTab();
			
		}
		
		/********************************************************************************************************************
		 * Function Name:verifyVoltageDropCalculation
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner:
		 * Last Update : Alpesh Dhakad - 05/12/2019 - Updated test scripts with new method for loading details
		 * Alpesh Dhakad - 19/01/2021 Updated script as per new UI implementation 
		 ********************************************************************************************************************/
		// Verify Voltage Drop Calculation on Adding devices in loops
		[UserCodeMethod]
		public static void verifyVoltageDropCalculation(string sFileName,string sAddDevicesLoop)
		{
			// Open the excel file and sheet with mentioned name in argument
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesLoop);
			
			// Count the number of rows in excel
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared various fields as String type
			string Multichannel,sLabelName,expectedDCUnits,expectedVoltDrop,expectedVoltDropWorstcase,expectedMaxVoltDrop,expectedMaxVoltDropWorstcase;
			string LoadingDetailsNameForVoltDrop,LoadingDetailsNameForVoltDropWorstCase,LoadingDetailsNameForMaxVoltDrop,LoadingDetailsNameForMaxVoltDropWorstCase;
			string LoopA_Details,LoopB_Details,LoopC_Details;
			
				
			LoopA_Details=((Range)Excel_Utilities.ExcelRange.Cells[2,10]).Value.ToString();
			LoopB_Details=((Range)Excel_Utilities.ExcelRange.Cells[3,10]).Value.ToString();
			LoopC_Details=((Range)Excel_Utilities.ExcelRange.Cells[4,10]).Value.ToString();
			//sColumn=((Range)Excel_Utilities.ExcelRange.Cells[2,11]).Value.ToString();
			
				
			// For loop to fetch values from the excel sheet and then add devices
			for(int i=6; i<=rows; i++)
			{
				ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				string sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				Multichannel = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				expectedDCUnits= ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				expectedVoltDrop = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				expectedVoltDropWorstcase = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				LoadingDetailsNameForVoltDrop = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				LoadingDetailsNameForVoltDropWorstCase = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				
				
				// Add devices from the gallery as per test data from the excel sheet
				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
				
				//Go to Physical layout
				Common_Functions.clickOnPhysicalLayoutTab();
				
				Common_Functions.clickOnPanelCalculationsTab();
				
				// Call verifyDCUnitValue method and Verify DC units value
				//DC_Functions.verifyDCUnitsValue(expectedDCUnits);
				//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,"Current (DC Units)");
				Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopA_Details,"2");

				
				
				// Call verifyVoltDropValue method and Verify VoltDrop value value
				//verifyVoltDropValue(expectedVoltDrop);
				//Devices_Functions.verifyLoadingDetailsValue(expectedVoltDrop,LoadingDetailsNameForVoltDrop);
				Devices_Functions.verifyLoopLoadingDetailsValue(expectedVoltDrop,LoopA_Details,"3");
				
				
				// Call verifyVoltDropWorstCaseValue method and Verify VoltDrop Worst case value value
				//verifyVoltDropWorstCaseValue(expectedVoltDropWorstcase);
				//Devices_Functions.verifyLoadingDetailsValue(expectedVoltDropWorstcase,LoadingDetailsNameForVoltDropWorstCase);
				Devices_Functions.verifyLoopLoadingDetailsValue(expectedVoltDropWorstcase,LoopA_Details,"4");
				
				
				//verifyVoltageDropColor();
				//Devices_Functions.verifyLoadingDetailColor(LoadingDetailsNameForVoltDrop);
				Devices_Functions.verifyLoopLoadingDetailColor(LoopA_Details,"3");
				
				//verifyVoltageDropWorstCaseColor();
				//Devices_Functions.verifyLoadingDetailColor(LoadingDetailsNameForVoltDropWorstCase);
				Devices_Functions.verifyLoopLoadingDetailColor(LoopA_Details,"4");
				
				Common_Functions.clickOnPropertiesTab();
				
				//Click on Points tab
				Common_Functions.clickOnPointsTab();
				
			}
			
			// Fetch value from excel sheet and store it
			expectedMaxVoltDrop = ((Range)Excel_Utilities.ExcelRange.Cells[3,4]).Value.ToString();
			expectedMaxVoltDropWorstcase = ((Range)Excel_Utilities.ExcelRange.Cells[4,4]).Value.ToString();
			LoadingDetailsNameForMaxVoltDrop = ((Range)Excel_Utilities.ExcelRange.Cells[3,5]).Value.ToString();
			LoadingDetailsNameForMaxVoltDropWorstCase = ((Range)Excel_Utilities.ExcelRange.Cells[4,5]).Value.ToString();
			
			Common_Functions.clickOnPanelCalculationsTab();
				
			// Call verifyMaxVoltDrop method and Verify Max volt drop value
			//verifyMaxVoltDrop(expectedMaxVoltDrop);
			//Devices_Functions.verifyMaxLoadingDetailsValue(expectedMaxVoltDrop,LoadingDetailsNameForMaxVoltDrop);
			Devices_Functions.verifyMaxLoopLoadingDetailsValue(expectedMaxVoltDrop,LoopA_Details,"3");
				
			
			// Call verifyMaxVoltDropWorstCaseValue method and Verify Max volt drop worst case value
			//verifyMaxVoltDropWorstCaseValue(expectedMaxVoltDropWorstcase);
			//Devices_Functions.verifyMaxLoadingDetailsValue(expectedMaxVoltDropWorstcase,LoadingDetailsNameForMaxVoltDropWorstCase);
			Devices_Functions.verifyMaxLoopLoadingDetailsValue(expectedMaxVoltDropWorstcase,LoopA_Details,"4");
				
			Common_Functions.clickOnPropertiesTab();
			
			// Close the currently opened excel sheet
			Excel_Utilities.CloseExcel();
			
		}
		
		/************************************************************************************************************************
		 * Function Name:
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner:
		 * Last Update : Alpesh Dhakad - 05/12/2019 - Updated test scripts with new methods for loading details
		 * Alpesh Dhakad - 19/01/2021 Updated script as per new UI implementation
		 ************************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyVoltageDropPercentage(string sFileName, string noLoadVoltDrop)
		{
			string LoopA_Details,LoopB_Details,LoopC_Details;
			
			//Go to Physical layout
			Common_Functions.clickOnPhysicalLayoutTab();
			
			Common_Functions.clickOnPanelCalculationsTab();

			// Check Voltage drop is set to 0.00 when no devices are added
			//verifyVoltDropValue(noLoadVoltDrop);
			//verifyVoltDropWorstCaseValue(noLoadVoltDrop);
			
			//Devices_Functions.verifyLoadingDetailsValue(noLoadVoltDrop,"Volt Drop (V)");
			//Devices_Functions.verifyLoadingDetailsValue(noLoadVoltDrop,"Volt Drop (worst case)");

			Devices_Functions.verifyLoopLoadingDetailsValue(noLoadVoltDrop,"Built-in Loop-A","3");
			Devices_Functions.verifyLoopLoadingDetailsValue(noLoadVoltDrop,"Built-in Loop-A","4");
					
			
			// Navigate to Points tab
			Common_Functions.clickOnPointsTab();
			
			
			Common_Functions.clickOnPropertiesTab();
			
			// Add devices such that voltage drop will show green color
			Devices_Functions.AddMultipleDevices(sFileName, "GreenColorPercentage");
			
			// Verify progress bar color as per percentage
			Report.Log(ReportLevel.Info, "Verifying Volt drop progress bar color is LightGreen.");
			//verifyVoltageDropColor();
			//Devices_Functions.verifyLoadingDetailColor("Volt Drop (V)");
			Devices_Functions.verifyLoopLoadingDetailColor("Built-in Loop-A","3");
				
			
			// Verify progress bar color as per percentage
			Report.Log(ReportLevel.Info, "Verifying Volt drop (worst case) progress bar color is LightGreen.");
			//verifyVoltageDropWorstCaseColor();
			//Devices_Functions.verifyLoadingDetailColor("Volt Drop (worst case)");
			Devices_Functions.verifyLoopLoadingDetailColor("Built-in Loop-A","4");
				
			// Delete All devices
			Devices_Functions.DeleteAllDevices();
			
			// Add devices such that voltage drop will show Yellow color
			Devices_Functions.AddMultipleDevices(sFileName, "VDWorstCaseYellowPercentage");
			
			Common_Functions.clickOnPanelCalculationsTab();
			
			// Verify progress bar color as per percentage
			Report.Log(ReportLevel.Info, "Verifying Volt drop progress bar color is LightGreen.");
			//verifyVoltageDropColor();
			//Devices_Functions.verifyLoadingDetailColor("Volt Drop (V)");
			Devices_Functions.verifyLoopLoadingDetailColor("Built-in Loop-A","3");
			
			
			// Verify progress bar color as per percentage
			Report.Log(ReportLevel.Info, "Verifying Volt drop (worst case) progress bar color is Yellow.");
			//verifyVoltageDropWorstCaseColor();
			//Devices_Functions.verifyLoadingDetailColor("Volt Drop (worst case)");
			Devices_Functions.verifyLoopLoadingDetailColor("Built-in Loop-A","4");
			
			Common_Functions.clickOnPropertiesTab();
			
			// Delete All devices
			Devices_Functions.DeleteAllDevices();

			// Add devices such that voltage drop will show Yellow color
			Devices_Functions.AddMultipleDevices(sFileName, "VtgDropYellowColorPercentage");
		
			Common_Functions.clickOnPanelCalculationsTab();
			
			// Verify progress bar color as per percentage
			Report.Log(ReportLevel.Info, "Verifying Volt drop progress bar color is Yellow.");
			//verifyVoltageDropColor();
			//Devices_Functions.verifyLoadingDetailColor("Volt Drop (V)");
			Devices_Functions.verifyLoopLoadingDetailColor("Built-in Loop-A","3");
			
	
			
			// Verify progress bar color as per percentage
			Report.Log(ReportLevel.Info, "Verifying Volt drop (worst case) progress bar color is Pink.");
			//verifyVoltageDropWorstCaseColor();
			//Devices_Functions.verifyLoadingDetailColor("Volt Drop (worst case)");
			Devices_Functions.verifyLoopLoadingDetailColor("Built-in Loop-A","3");
		
			Common_Functions.clickOnPropertiesTab();
			
			// Delete All devices
			Devices_Functions.DeleteAllDevices();

			// Add devices such that voltage drop will show Yellow color
			Devices_Functions.AddMultipleDevices(sFileName, "RedColorPercentage");
			
			Common_Functions.clickOnPanelCalculationsTab();
			
			// Verify progress bar color as per percentage
			Report.Log(ReportLevel.Info, "Verifying Volt drop progress bar color is Pink.");
			//verifyVoltageDropColor();
			//Devices_Functions.verifyLoadingDetailColor("Volt Drop (V)");
			Devices_Functions.verifyLoopLoadingDetailColor("Built-in Loop-A","3");
			
			
			// Verify progress bar color as per percentage
			Report.Log(ReportLevel.Info, "Verifying Volt drop (worst case) progress bar color is Pink.");
			//verifyVoltageDropWorstCaseColor();
			//Devices_Functions.verifyLoadingDetailColor("Volt Drop (worst case)");
			Devices_Functions.verifyLoopLoadingDetailColor("Built-in Loop-A","4");
			
			Common_Functions.clickOnPropertiesTab();
			
			// Navigate to Points tab
			Common_Functions.clickOnPointsTab();
			
			
			// Delete All devices
			Devices_Functions.DeleteAllDevices();
		}
		
		/********************************************************************
		 * Function Name:
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner:
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void verifyVoltageDropColor()
		{
			string expectedColor;
			
			//Go to Physical layout
			Common_Functions.clickOnPhysicalLayoutTab();
			Delay.Duration(500, false);
			
			float ActualVoltDropValue = float.Parse(repo.ProfileConsys1.VoltDropValue.TextValue);
			
			//repo.ProfileConsys1.DCUnitProgressBar.GetAttributeValue
			string actualColour = repo.ProfileConsys1.VoltDropProgressBar.GetAttributeValue<string>("foreground");
			
			//Fetch max volt drop text value and storing it in string
			float maxVoltDropValue = float.Parse(repo.ProfileConsys1.MaxVoltDropValue.TextValue);
			
			expectedColor = Devices_Functions.calculatePercentage(ActualVoltDropValue, maxVoltDropValue);
			
			Devices_Functions.VerifyPercentage(expectedColor, actualColour);
		}
		
		/********************************************************************
		 * Function Name:
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner:
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void verifyVoltageDropWorstCaseColor()
		{
			string expectedColor;
			
			float ActualVoltDropWorstCaseValue = float.Parse(repo.ProfileConsys1.VoltDropWorstCaseValue.TextValue);

			string actualColour = repo.ProfileConsys1.VoltDropWorstCaseProgressBar.GetAttributeValue<string>("foreground");
			
			//Fetch max volt drop worst case text value and storing it in string
			float maxVoltDropWorstCaseValue = float.Parse(repo.ProfileConsys1.MaxVoltDropWorstCaseValue.TextValue);
			
			expectedColor = Devices_Functions.calculatePercentage(ActualVoltDropWorstCaseValue, maxVoltDropWorstCaseValue);
			
			Devices_Functions.VerifyPercentage(expectedColor, actualColour);
			
			// Navigate to Points tab
			Common_Functions.clickOnPointsTab();
			Delay.Duration(500, false);
		}

		/********************************************************************************************************************
		 * Function Name:verifyVoltageDropCalculation
		 * Function Details:
		 * Parameter/Arguments: filename, Sheetname, LoopDetails
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 20/01/2021
		 ********************************************************************************************************************/
		// Verify Voltage Drop Calculation on Adding devices in loops
		[UserCodeMethod]
		public static void verifyVoltageDropCalculation(string sFileName,string sAddDevicesLoop, string LoopDetails)
		{
		// Open the excel file and sheet with mentioned name in argument
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesLoop);
			
			// Count the number of rows in excel
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared various fields as String type
			string Multichannel,sLabelName,expectedDCUnits,expectedVoltDrop,expectedVoltDropWorstcase,expectedMaxVoltDrop,expectedMaxVoltDropWorstcase;
			string LoadingDetailsNameForVoltDrop,LoadingDetailsNameForVoltDropWorstCase,LoadingDetailsNameForMaxVoltDrop,LoadingDetailsNameForMaxVoltDropWorstCase;
			string LoopA_Details,LoopB_Details,LoopC_Details;
			
				
			LoopA_Details=((Range)Excel_Utilities.ExcelRange.Cells[2,10]).Value.ToString();
			LoopB_Details=((Range)Excel_Utilities.ExcelRange.Cells[3,10]).Value.ToString();
			LoopC_Details=((Range)Excel_Utilities.ExcelRange.Cells[4,10]).Value.ToString();
			//sColumn=((Range)Excel_Utilities.ExcelRange.Cells[2,11]).Value.ToString();
			
				
			// For loop to fetch values from the excel sheet and then add devices
			for(int i=6; i<=rows; i++)
			{
				ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				string sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				Multichannel = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				expectedDCUnits= ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				expectedVoltDrop = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				expectedVoltDropWorstcase = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				LoadingDetailsNameForVoltDrop = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				LoadingDetailsNameForVoltDropWorstCase = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				
				
				// Add devices from the gallery as per test data from the excel sheet
				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
				
				//Go to Physical layout
				Common_Functions.clickOnPhysicalLayoutTab();
				
				Common_Functions.clickOnPanelCalculationsTab();
				
				// Call verifyDCUnitValue method and Verify DC units value
				//DC_Functions.verifyDCUnitsValue(expectedDCUnits);
				//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,"Current (DC Units)");
				Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopDetails,"2");

				
				
				// Call verifyVoltDropValue method and Verify VoltDrop value value
				//verifyVoltDropValue(expectedVoltDrop);
				//Devices_Functions.verifyLoadingDetailsValue(expectedVoltDrop,LoadingDetailsNameForVoltDrop);
				Devices_Functions.verifyLoopLoadingDetailsValue(expectedVoltDrop,LoopDetails,"3");
				
				
				// Call verifyVoltDropWorstCaseValue method and Verify VoltDrop Worst case value value
				//verifyVoltDropWorstCaseValue(expectedVoltDropWorstcase);
				//Devices_Functions.verifyLoadingDetailsValue(expectedVoltDropWorstcase,LoadingDetailsNameForVoltDropWorstCase);
				Devices_Functions.verifyLoopLoadingDetailsValue(expectedVoltDropWorstcase,LoopDetails,"4");
				
				
				//verifyVoltageDropColor();
				//Devices_Functions.verifyLoadingDetailColor(LoadingDetailsNameForVoltDrop);
				Devices_Functions.verifyLoopLoadingDetailColor(LoopDetails,"3");
				
				//verifyVoltageDropWorstCaseColor();
				//Devices_Functions.verifyLoadingDetailColor(LoadingDetailsNameForVoltDropWorstCase);
				Devices_Functions.verifyLoopLoadingDetailColor(LoopDetails,"4");
				
				Common_Functions.clickOnPropertiesTab();
				
				//Click on Points tab
				Common_Functions.clickOnPointsTab();
			}
			
			// Fetch value from excel sheet and store it
			expectedMaxVoltDrop = ((Range)Excel_Utilities.ExcelRange.Cells[3,4]).Value.ToString();
			expectedMaxVoltDropWorstcase = ((Range)Excel_Utilities.ExcelRange.Cells[4,4]).Value.ToString();
			LoadingDetailsNameForMaxVoltDrop = ((Range)Excel_Utilities.ExcelRange.Cells[3,5]).Value.ToString();
			LoadingDetailsNameForMaxVoltDropWorstCase = ((Range)Excel_Utilities.ExcelRange.Cells[4,5]).Value.ToString();
			
			Common_Functions.clickOnPanelCalculationsTab();
				
			// Call verifyMaxVoltDrop method and Verify Max volt drop value
			//verifyMaxVoltDrop(expectedMaxVoltDrop);
			//Devices_Functions.verifyMaxLoadingDetailsValue(expectedMaxVoltDrop,LoadingDetailsNameForMaxVoltDrop);
			Devices_Functions.verifyMaxLoopLoadingDetailsValue(expectedMaxVoltDrop,LoopDetails,"3");
				
			
			// Call verifyMaxVoltDropWorstCaseValue method and Verify Max volt drop worst case value
			//verifyMaxVoltDropWorstCaseValue(expectedMaxVoltDropWorstcase);
			//Devices_Functions.verifyMaxLoadingDetailsValue(expectedMaxVoltDropWorstcase,LoadingDetailsNameForMaxVoltDropWorstCase);
			Devices_Functions.verifyMaxLoopLoadingDetailsValue(expectedMaxVoltDropWorstcase,LoopDetails,"4");
				
			Common_Functions.clickOnPropertiesTab();
			
			// Close the currently opened excel sheet
			Excel_Utilities.CloseExcel();
			
			
		}
	}
}
