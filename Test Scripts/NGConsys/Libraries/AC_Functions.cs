/*
 * Created by Ranorex
 * User: jdhakaa
 * Date: 11/22/2018
 * Time: 4:21 PM
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
	public class AC_Functions
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
		
		
		/********************************************************************************************************************
		 * Function Name: VerifyACCalculation
		 * Function Details: To verify AC calculation after adding and deleting devices
		 * Parameter/Arguments: sFileName, Add device, Delete Device
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update : Alpesh Dhakad - 30/07/2019 & 21/08/2019 - Updated scripts as per new build and xpaths
		 * Last Update:Poonam Kadam-12/9/19- updated loading details method
		 ********************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyACCalculation(string sFileName,string sAddDevicesSheet, string sDeleteDevicesSheet)
		{
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
			//Excel_Utilities.OpenSheet(sAddDevicesSheet);
			float fAcUnits,fMaxACUnits;
			int calculatedAcUnits=0,calculatedAcUnitsofLoopB,calculatedAcUnitsofLoopA;
			string actualColour,ACUnitsLoadingDetails;
			
			//Get Max AC Units and AC Units for loop A using input excel sheet
			calculatedAcUnitsofLoopA = calculateACUnits(sFileName,sAddDevicesSheet);
			
			sACUnits = calculatedAcUnitsofLoopA.ToString();
			sMaxACUnits =  ((Range)Excel_Utilities.ExcelRange.Cells[8,4]).Value.ToString();
			ACUnitsLoadingDetails=((Range)Excel_Utilities.ExcelRange.Cells[8,5]).Value.ToString();
			//Convert to float and calculate AC Units percentage to identify Color of progress bar
			float.TryParse(sACUnits,out fAcUnits);
			float.TryParse(sMaxACUnits,out fMaxACUnits);
			string expectedColorCode = Devices_Functions.calculatePercentage(fAcUnits,fMaxACUnits);
			
			//Go to Physical layout
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			Delay.Duration(500, false);
			
			//Verify Max AC Units and AC Units
			verifyMaxLoadingDetailsValue(sMaxACUnits,ACUnitsLoadingDetails);
			//verifyMaxACUnits();
			verifyLoadingDetailsValue(sACUnits,ACUnitsLoadingDetails);
			//verifyACUnits();
			
			//Get and verify progressbar color from UI
			repo.ProfileConsys1.cell_ACUnits.Click();
			actualColour = repo.ProfileConsys1.ACUnitProgressBar.GetAttributeValue<string>("foreground");
			Devices_Functions.VerifyPercentage(expectedColorCode,actualColour);
			
			//Verify AC units displayed for loop B
			Report.Log(ReportLevel.Info,"Verification of AC Units of Loop B");
			
			// Click on Loop B node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-B");
				
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			//verify Max AC units and AC Units
			verifyMaxLoadingDetailsValue(sMaxACUnits,ACUnitsLoadingDetails);
			//verifyMaxACUnits();
			verifyLoadingDetailsValue(sACUnits,ACUnitsLoadingDetails);
			//verifyACUnits();
			
			//Get and verify progressbar color from UI
			repo.ProfileConsys1.cell_ACUnits.Click();
			actualColour = repo.ProfileConsys1.ACUnitProgressBar.GetAttributeValue<string>("foreground");
			Devices_Functions.VerifyPercentage(expectedColorCode,actualColour);
			
			//Add devices in loop B and calculate AC Units percentage
			repo.ProfileConsys1.tab_Points.Click();
			calculatedAcUnitsofLoopB = calculateACUnits(sFileName,sAddDevicesSheet);
			
			calculatedAcUnits= calculatedAcUnitsofLoopB + calculatedAcUnitsofLoopA;
			sACUnits = calculatedAcUnits.ToString();
			float.TryParse(sACUnits,out fAcUnits);
			expectedColorCode = Devices_Functions.calculatePercentage(fAcUnits,fMaxACUnits);
			
			//verify Actual AC Units displayed for loop B after addition of devices
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			Delay.Duration(500, false);
			Report.Log(ReportLevel.Info,"Verification of AC Units of Loop B after addition of devices");
			verifyACUnits();
			
			//Get and verify progressbar color from UI
			repo.ProfileConsys1.cell_ACUnits.Click();
			actualColour = repo.ProfileConsys1.ACUnitProgressBar.GetAttributeValue<string>("foreground");
			Devices_Functions.VerifyPercentage(expectedColorCode,actualColour);
			
			//Verify AC units displayed for loop A
			Report.Log(ReportLevel.Info,"Verification of AC Units of Loop A");
			// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				repo.ProfileConsys1.tab_PhysicalLayout.Click();
			Delay.Duration(500, false);
			verifyLoadingDetailsValue(sACUnits,ACUnitsLoadingDetails);
			//verifyACUnits();
			
			//Get and verify progressbar color from UI
			repo.ProfileConsys1.cell_ACUnits.Click();
			actualColour = repo.ProfileConsys1.ACUnitProgressBar.GetAttributeValue<string>("foreground");
			Devices_Functions.VerifyPercentage(expectedColorCode,actualColour);
			
			//Close Excel
			Excel_Utilities.ExcelWB.Close(false, null, null);
			Excel_Utilities.ExcelAppl.Quit();
			
			//Delete devices from loop
			repo.ProfileConsys1.tab_Points.Click();
			Devices_Functions.DeleteDevices(sFileName,sDeleteDevicesSheet);
			float ACUnits;
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			float deletedACUnits=0,actualACUnits;
			
			//Delete device and calculate AC Units of deleted devices
			for(int i=8;i<=rows;i++)
			{
				sACUnits = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				float.TryParse(sACUnits,out ACUnits);
				deletedACUnits = deletedACUnits + ACUnits;
			}
			
			//Substract AC units from earlier AC Units
			actualACUnits = calculatedAcUnits-deletedACUnits;
			sACUnits= actualACUnits.ToString();
			
			//Verify AC units displayed for loop A
			Report.Log(ReportLevel.Info,"Verification of AC Units of Loop A");
			// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				repo.ProfileConsys1.tab_PhysicalLayout.Click();
			Delay.Duration(500, false);
			verifyLoadingDetailsValue(sACUnits,ACUnitsLoadingDetails);
			//verifyACUnits();
			
			//Verify AC units displayed for loop B
			Report.Log(ReportLevel.Info,"Verification of AC Units of Loop B");
			
			// Click on Loop B node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-B");
				
				repo.ProfileConsys1.tab_PhysicalLayout.Click();
			Delay.Duration(500, false);
		verifyLoadingDetailsValue(sACUnits,ACUnitsLoadingDetails);
			//verifyACUnits();
			
			
			Excel_Utilities.ExcelWB.Close(false, null, null);
			Excel_Utilities.ExcelAppl.Quit();
		}
		
		/********************************************************************
		 * Function Name: calculateACUnits
		 * Function Details: To calculate AC Unit
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :
		 ********************************************************************/
		public static int calculateACUnits(string sFileName, string sAddDevicesSheet)
		{
			int calculatedAcUnits=0;
			string sDeviceName,sType;
			int rows = Excel_Utilities.ExcelRange.Rows.Count;
			for(int i=8;i<=rows;i++)
			{
				sRow=(i-1).ToString();
				sDeviceName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				Devices_Functions.AddDevicesfromGallery(sDeviceName,sType);
				
				int DeviceACUnits = int.Parse(((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString());
				calculatedAcUnits = calculatedAcUnits + DeviceACUnits;
			}
			
			return calculatedAcUnits;
		}
		
		
		/********************************************************************
		 * Function Name: verifyACUnits
		 * Function Details: To verify AC unit
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :
		 ********************************************************************/
		public static void verifyACUnits()
		{
			if(repo.ProfileConsys1.actualACUnits.EnsureVisible())
			{
				Report.Log(ReportLevel.Success,"AC units displayed correctly "+ sACUnits);
			}
			else
			{
				Report.Log(ReportLevel.Failure,"AC units displayed incorrectly "+ sACUnits);
			}
		}
		
		/********************************************************************
		 * Function Name: verifyMaxACUnits
		 * Function Details: To verify MAX AC units
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :
		 ********************************************************************/
		public static void verifyMaxACUnits()
		{
			if(repo.ProfileConsys1.actualMaxACunits.EnsureVisible())
			{
				Report.Log(ReportLevel.Success,"Max AC units displayed correctly "+ sMaxACUnits);
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Max AC units displayed correctly incorrectly "+ sMaxACUnits);
			}
		}
		
		/**************************************************************************************************
		 * Function Name: VerifyACCalculationforFIM
		 * Function Details: verify AC Calculation for FIM loop
		 * Parameter/Arguments: sFileName, Add device, Delete Device
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update : Alpesh Dhakad - 30/07/2019 & 21/08/2019 - Updated scripts as per new build and xpaths
		 **************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyACCalculationforFIM(string sFileName,string sAddDevicesSheet, string sDeleteDevicesSheet)
		{
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
			//Excel_Utilities.OpenSheet(sAddDevicesSheet);
			float fAcUnits,fMaxACUnits;
			int calculatedAcUnitsofLoopB,calculatedAcUnitsofLoopA;
			string actualColour;
			
			//Get Max AC Units and AC Units for loop A using input excel sheet
			calculatedAcUnitsofLoopA = calculateACUnits(sFileName,sAddDevicesSheet);
			sACUnits = calculatedAcUnitsofLoopA.ToString();
			sMaxACUnits =  ((Range)Excel_Utilities.ExcelRange.Cells[8,4]).Value.ToString();
			
			//Convert to float and calculate AC Units percentage to identify Color of progress bar
			float.TryParse(sACUnits,out fAcUnits);
			float.TryParse(sMaxACUnits,out fMaxACUnits);
			string expectedColorCode = Devices_Functions.calculatePercentage(fAcUnits,fMaxACUnits);
			
			//Go to Physical layout
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			Delay.Duration(500, false);
			
			//Verify Max AC Units and AC Units of Loop A
			Report.Log(ReportLevel.Info,"Verification of AC Units of Loop A");
			verifyMaxACUnits();
			verifyACUnits();
			
			//Get and verify progressbar color from UI
			repo.ProfileConsys1.cell_ACUnits.Click();
			actualColour = repo.ProfileConsys1.ACUnitProgressBar.GetAttributeValue<string>("foreground");
			Devices_Functions.VerifyPercentage(expectedColorCode,actualColour);
			
			//Verify AC units displayed for loop B : AC units of Loop A should not be reflected in Loop B
			Report.Log(ReportLevel.Info,"Verification of AC Units of Loop B");
			
			// Click on Loop B node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-B");
				
				repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			//verify Max AC units and AC Units
			verifyMaxACUnits();
			sACUnits="0";
			verifyACUnits();
			
			//Get and verify progressbar color from UI
//			float.TryParse(sACUnits,out fAcUnits);
//			float.TryParse(sMaxACUnits,out fMaxACUnits);
//			expectedColorCode = calculatePercentage(fAcUnits,fMaxACUnits);
//			repo.ProfileConsys1.cell_ACUnits.Click();
//			actualColour = repo.ProfileConsys1.ACUnitProgressBar.GetAttributeValue<string>("foreground");
//			VerifyPercentage(expectedColorCode,actualColour);
//
			//Add devices in loop B and calculate AC Units percentage
			repo.ProfileConsys1.tab_Points.Click();
			calculatedAcUnitsofLoopB = calculateACUnits(sFileName,sAddDevicesSheet);
			//calculatedAcUnits= calculatedAcUnitsofLoopB;
			sACUnits = calculatedAcUnitsofLoopB.ToString();
			float.TryParse(sACUnits,out fAcUnits);
			expectedColorCode = Devices_Functions.calculatePercentage(fAcUnits,fMaxACUnits);
			
			//verify Actual AC Units displayed for loop B after addition of devices
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			Delay.Duration(500, false);
			Report.Log(ReportLevel.Info,"Verification of AC Units of Loop B after addition of devices");
			verifyACUnits();
			
			//Get and verify progressbar color from UI
			repo.ProfileConsys1.cell_ACUnits.Click();
			actualColour = repo.ProfileConsys1.ACUnitProgressBar.GetAttributeValue<string>("foreground");
			Devices_Functions.VerifyPercentage(expectedColorCode,actualColour);
			
			//Verify AC units displayed for loop A
			Report.Log(ReportLevel.Info,"Verification of AC Units of Loop A");
			// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
			
				repo.ProfileConsys1.tab_PhysicalLayout.Click();
			Delay.Duration(500, false);
			sACUnits = calculatedAcUnitsofLoopA.ToString();
			float.TryParse(sACUnits,out fAcUnits);
			expectedColorCode = Devices_Functions.calculatePercentage(fAcUnits,fMaxACUnits);
			verifyACUnits();
			
			//Get and verify progressbar color from UI
			repo.ProfileConsys1.cell_ACUnits.Click();
			actualColour = repo.ProfileConsys1.ACUnitProgressBar.GetAttributeValue<string>("foreground");
			Devices_Functions.VerifyPercentage(expectedColorCode,actualColour);
			
			//Close Excel
			Excel_Utilities.ExcelWB.Close(false, null, null);
			Excel_Utilities.ExcelAppl.Quit();
			
			//Delete devices from loop
			repo.ProfileConsys1.tab_Points.Click();
			Devices_Functions.DeleteDevices(sFileName,sDeleteDevicesSheet);
			float ACUnits;
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			float deletedACUnits=0,actualACUnits;
			
			//Delete device and calculate AC Units of deleted devices
			for(int i=8;i<=rows;i++)
			{
				sACUnits = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				float.TryParse(sACUnits,out ACUnits);
				deletedACUnits = deletedACUnits + ACUnits;
			}
			
			//Substract AC units from earlier AC Units
			actualACUnits = calculatedAcUnitsofLoopA-deletedACUnits;
			sACUnits= actualACUnits.ToString();
			
			//Verify AC units displayed for loop A
			Report.Log(ReportLevel.Info,"Verification of AC Units of Loop A");
			// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			Delay.Duration(500, false);
			verifyACUnits();
			
			//Verify AC units displayed for loop B
			Report.Log(ReportLevel.Info,"Verification of AC Units of Loop B");
			
			// Click on Loop B node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-B");
			
			
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			Delay.Duration(500, false);
			sACUnits = calculatedAcUnitsofLoopB.ToString();
			verifyACUnits();
			
			
			Excel_Utilities.ExcelWB.Close(false, null, null);
			Excel_Utilities.ExcelAppl.Quit();
		}
		
		/********************************************************************
		 * Function Name: verifyMaxACUnitsValue
		 * Function Details: Verify Max AC Units Value
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update : 19/04/2019 - Alpesh Dhakad - Added last line in this method
		 ********************************************************************/
		[UserCodeMethod]
		public static void verifyMaxACUnitsValue(string expectedMaxACUnits)
		{
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			string maxACUnits = repo.ProfileConsys1.MaxACUnitsValue.TextValue;
			
			if(maxACUnits.Equals(expectedMaxACUnits))
			{
				Report.Log(ReportLevel.Success,"Max AC Units are displayed correctly " +maxACUnits);
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Max AC Units are not displayed correctly " +maxACUnits);
			}
			repo.ProfileConsys1.tab_Points.Click();
		}
		
		/********************************************************************
		 * Function Name: verifyACUnitsValue
		 * Function Details: To verify AC Units value
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void verifyACUnitsValue(string expectedACUnits)
		{
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			string ActualAcUnits = repo.ProfileConsys1.ACUnits.TextValue;
			
			if(ActualAcUnits.Equals(expectedACUnits))
			{
				Report.Log(ReportLevel.Success,"AC Units " + ActualAcUnits + " displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"AC Units are not displayed correctly " + ", Expected AC Units:  " + expectedACUnits  + " Actual AC Units: "+ ActualAcUnits);
			}
			
		}
		
		/*****************************************************************************************************************
		 * Function Name: VerifyACUnitIndicationWithISDevices
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 19/04/2019
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyACUnitIndicationWithISDevices()
		{
			string actualColour,expectedColor;
			
			//Go to Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			//Go to Physical layout
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			

			float ActualACUnits = float.Parse(repo.ProfileConsys1.ACUnits.TextValue);
			
			//Retrieve foreground color
			actualColour = repo.ProfileConsys1.ACUnitProgressBar.GetAttributeValue<string>("foreground");
			
			
			//Fetch max AC value drop text value and storing it in string
			float maxACUnitsValue = float.Parse(repo.ProfileConsys1.MaxACUnits.TextValue);
			
			// To calculate and get the expected color value
			expectedColor = Devices_Functions.calculatePercentage(ActualACUnits, maxACUnitsValue);
			
			// To verify Percentage
			Devices_Functions.VerifyPercentage(expectedColor, actualColour);
			
			//Go to Points tab
			repo.ProfileConsys1.tab_Points.Click();
		}
		
		
	}
}
