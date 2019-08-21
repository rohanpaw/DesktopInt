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
		
		
		/*****************************************************************************************************************
		 * Function Name:
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner:
		 * Last Update : Alpesh Dhakad - 30/07/2019 and 21/08/2019 - Updated script as per new build and xpath
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
				
				// Add devices from the gallery as per test data from the excel sheet
				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
				
				//Go to Physical layout
				repo.ProfileConsys1.tab_PhysicalLayout.Click();
				Delay.Duration(500, false);
				
				// Call verifyDCUnitValue method and Verify DC units value
				DC_Functions.verifyDCUnitsValue(expectedDCUnits);
				
				// Call verifyVoltDropValue method and Verify VoltDrop value value
				verifyVoltDropValue(expectedVoltDrop);
				
				// Call verifyVoltDropWorstCaseValue method and Verify VoltDrop Worst case value value
				verifyVoltDropWorstCaseValue(expectedVoltDropWorstcase);
				
				//Click on Points tab
				repo.ProfileConsys1.tab_Points.Click();
			}
			
			// Fetch value from excel sheet and store it
			expectedMaxVoltDrop = ((Range)Excel_Utilities.ExcelRange.Cells[3,4]).Value.ToString();
			expectedMaxVoltDropWorstcase = ((Range)Excel_Utilities.ExcelRange.Cells[4,4]).Value.ToString();
			
			
			// Call verifyMaxVoltDrop method and Verify Max volt drop value
			verifyMaxVoltDrop(expectedMaxVoltDrop);
			
			// Call verifyMaxVoltDropWorstCaseValue method and Verify Max volt drop worst case value
			verifyMaxVoltDropWorstCaseValue(expectedMaxVoltDropWorstcase);
			
			// Close the currently opened excel sheet
			Excel_Utilities.CloseExcel();
			
			//Delete Devices from loop A
			
			// Click on Loop A node
			Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
			//Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
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
			
			//Click on Physical Layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			// Call verifyVoltDropValue method and Verify volt drop value
			verifyVoltDropValue(expectedVoltDrop);
			
			// Call verifyVoltDropWorstCaseValue method and Verify volt drop worst case value
			verifyVoltDropWorstCaseValue(expectedVoltDropWorstcase);
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
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
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
			repo.ProfileConsys1.tab_Points.Click();
			
			
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
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
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
			repo.ProfileConsys1.tab_Points.Click();
			
			
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
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
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
			repo.ProfileConsys1.tab_Points.Click();
			
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
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
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
			repo.ProfileConsys1.tab_Points.Click();
			
		}
		
		/********************************************************************
		 * Function Name:
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner:
		 * Last Update :
		 ********************************************************************/
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
				
				// Add devices from the gallery as per test data from the excel sheet
				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
				
				//Go to Physical layout
				repo.ProfileConsys1.tab_PhysicalLayout.Click();
				Delay.Duration(500, false);
				
				// Call verifyDCUnitValue method and Verify DC units value
				DC_Functions.verifyDCUnitsValue(expectedDCUnits);
				
				// Call verifyVoltDropValue method and Verify VoltDrop value value
				verifyVoltDropValue(expectedVoltDrop);
				
				// Call verifyVoltDropWorstCaseValue method and Verify VoltDrop Worst case value value
				verifyVoltDropWorstCaseValue(expectedVoltDropWorstcase);
				
				verifyVoltageDropColor();
				verifyVoltageDropWorstCaseColor();
				
				//Click on Points tab
				repo.ProfileConsys1.tab_Points.Click();
			}
			
			// Fetch value from excel sheet and store it
			expectedMaxVoltDrop = ((Range)Excel_Utilities.ExcelRange.Cells[3,4]).Value.ToString();
			expectedMaxVoltDropWorstcase = ((Range)Excel_Utilities.ExcelRange.Cells[4,4]).Value.ToString();
			
			
			// Call verifyMaxVoltDrop method and Verify Max volt drop value
			verifyMaxVoltDrop(expectedMaxVoltDrop);
			
			// Call verifyMaxVoltDropWorstCaseValue method and Verify Max volt drop worst case value
			verifyMaxVoltDropWorstCaseValue(expectedMaxVoltDropWorstcase);
			
			// Close the currently opened excel sheet
			Excel_Utilities.CloseExcel();
			
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
		public static void verifyVoltageDropPercentage(string sFileName, string noLoadVoltDrop)
		{
			//Go to Physical layout
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			Delay.Duration(500, false);

			// Check Voltage drop is set to 0.00 when no devices are added
			verifyVoltDropValue(noLoadVoltDrop);
			verifyVoltDropWorstCaseValue(noLoadVoltDrop);
			
			// Navigate to Points tab
			repo.ProfileConsys1.tab_Points.Click();
			Delay.Duration(500, false);
			
			// Add devices such that voltage drop will show green color
			Devices_Functions.AddMultipleDevices(sFileName, "GreenColorPercentage");
			
			// Verify progress bar color as per percentage
			Report.Log(ReportLevel.Info, "Verifying Volt drop progress bar color is LightGreen.");
			verifyVoltageDropColor();
			
			// Verify progress bar color as per percentage
			Report.Log(ReportLevel.Info, "Verifying Volt drop (worst case) progress bar color is LightGreen.");
			verifyVoltageDropWorstCaseColor();
			
			// Delete All devices
			Devices_Functions.DeleteAllDevices();
			
			// Add devices such that voltage drop will show Yellow color
			Devices_Functions.AddMultipleDevices(sFileName, "VDWorstCaseYellowPercentage");
			
			// Verify progress bar color as per percentage
			Report.Log(ReportLevel.Info, "Verifying Volt drop progress bar color is LightGreen.");
			verifyVoltageDropColor();
			
			// Verify progress bar color as per percentage
			Report.Log(ReportLevel.Info, "Verifying Volt drop (worst case) progress bar color is Yellow.");
			verifyVoltageDropWorstCaseColor();
			
			// Delete All devices
			Devices_Functions.DeleteAllDevices();

			// Add devices such that voltage drop will show Yellow color
			Devices_Functions.AddMultipleDevices(sFileName, "VtgDropYellowColorPercentage");
			
			// Verify progress bar color as per percentage
			Report.Log(ReportLevel.Info, "Verifying Volt drop progress bar color is Yellow.");
			verifyVoltageDropColor();
			
			// Verify progress bar color as per percentage
			Report.Log(ReportLevel.Info, "Verifying Volt drop (worst case) progress bar color is Pink.");
			verifyVoltageDropWorstCaseColor();
			
			// Delete All devices
			Devices_Functions.DeleteAllDevices();

			// Add devices such that voltage drop will show Yellow color
			Devices_Functions.AddMultipleDevices(sFileName, "RedColorPercentage");
			
			// Verify progress bar color as per percentage
			Report.Log(ReportLevel.Info, "Verifying Volt drop progress bar color is Pink.");
			verifyVoltageDropColor();
			
			// Verify progress bar color as per percentage
			Report.Log(ReportLevel.Info, "Verifying Volt drop (worst case) progress bar color is Pink.");
			verifyVoltageDropWorstCaseColor();

			// Navigate to Points tab
			repo.ProfileConsys1.tab_Points.Click();
			Delay.Duration(500, false);
			
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
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
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
			repo.ProfileConsys1.tab_Points.Click();
			Delay.Duration(500, false);
		}

		
	}
}
