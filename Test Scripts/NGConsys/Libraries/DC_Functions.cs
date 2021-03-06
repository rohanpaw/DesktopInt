/*
 * Created by Ranorex
 * User: jdhakaa
 * Date: 11/22/2018
 * Time: 5:53 PM
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
	public class DC_Functions
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
		
		/********************************************************************
		 * Function Name: verifyMaxDCUnits
		 * Function Details: Verify maximum DC unit value
		 * Parameter/Arguments: Expected Maximum DC unit value
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 29/11/2018
		 ********************************************************************/
		[UserCodeMethod]
		public static void verifyMaxDCUnits(string expectedMaxDCUnits)
		{
			//repo.ProfileConsys1.tab_PhysicalLayout.Click();
			Common_Functions.clickOnPhysicalLayoutTab();
			
			string maxDCUnits = repo.ProfileConsys1.MaxDCUnits.TextValue;
			
			if(maxDCUnits.Equals(expectedMaxDCUnits))
			{
				Report.Log(ReportLevel.Success,"Max DC Units " + maxDCUnits + " is displayed correctly " );
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Max DC Units is not displayed correctly, it is displayed as: " + maxDCUnits + " instead of : " +expectedMaxDCUnits);
			}
			
			Common_Functions.clickOnPointsTab();
			
			
		}
		
		/********************************************************************
		 * Function Name: Verify DC Units value
		 * Function Details: Expected DC Units value
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 29/11/2018
		 ********************************************************************/
		[UserCodeMethod]
		public static void verifyDCUnitsValue(string expectedDCUnits)
		{
			//repo.ProfileConsys1.tab_PhysicalLayout.Click();
			Common_Functions.clickOnPhysicalLayoutTab();
			
			string ActualDcUnits = repo.ProfileConsys1.DCUnits.TextValue;
			
			if(ActualDcUnits.Equals(expectedDCUnits))
			{
				Report.Log(ReportLevel.Success,"DC Units " + ActualDcUnits + " is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"DC Units is not displayed correctly, DC Units displayed as: " +ActualDcUnits + " instead of : "+expectedDCUnits);
			}
			Common_Functions.clickOnPointsTab();
		}
		
		/********************************************************************
		 * Function Name: verifyDCUnitsWorstCaseValue
		 * Function Details: Verify DC units worst case value
		 * Parameter/Arguments: expected DC units worst case value
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 29/11/2018
		 ********************************************************************/
		[UserCodeMethod]
		public static void verifyDCUnitsWorstCaseValue(string expectedWorstCaseDCUnits)
		{
			//repo.ProfileConsys1.tab_PhysicalLayout.Click();
			Common_Functions.clickOnPhysicalLayoutTab();
			
			string ActualDcUnitsWorstCase = repo.ProfileConsys1.DCUnitsWorstCase.TextValue;
			
			if(ActualDcUnitsWorstCase.Equals(expectedWorstCaseDCUnits))
			{
				Report.Log(ReportLevel.Success,"DC Units worst case value " + ActualDcUnitsWorstCase + " is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"DC Units worst case value is not displayed correctly, DC Units displayed as: " +ActualDcUnitsWorstCase + " instead of : "+expectedWorstCaseDCUnits);
			}
			Common_Functions.clickOnPointsTab();
		}
		
		/********************************************************************
		 * Function Name: verifyMaxDCUnitsWorstCaseValue
		 * Function Details: Verify max DC units worst case value
		 * Parameter/Arguments: expected max DC units worst case value
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 29/11/2018
		 ********************************************************************/
		[UserCodeMethod]
		public static void verifyMaxDCUnitsWorstCaseValue(string expectedMaxDCUnitsWorstCase)
		{
			//repo.ProfileConsys1.tab_PhysicalLayout.Click();
			Common_Functions.clickOnPhysicalLayoutTab();
			
			string maxDCUnitsWorstCase = repo.ProfileConsys1.MaxDCUnitsWorstCase.TextValue;
			
			if(maxDCUnitsWorstCase.Equals(expectedMaxDCUnitsWorstCase))
			{
				Report.Log(ReportLevel.Success,"Max DC Units worst case value " + maxDCUnitsWorstCase + "is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Max DC Units worst case value is not displayed correctly, Max DC Unit is displayed as: " +maxDCUnitsWorstCase+ " instead of : " +expectedMaxDCUnitsWorstCase);
			}
			Common_Functions.clickOnPointsTab();
		}
		
		/********************************************************************
		 * Function Name: VerifyDCUnitsIndicators
		 * Function Details: To verify DC unit indicators
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyDCUnitsIndicators(string sFileName,string sAddDevicesSheet)
		{
			string expectedColorCode,sType,sDeviceName;
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			for(int j=8;j<=rows;j++)
			{
				sDeviceName =  ((Range)Excel_Utilities.ExcelRange.Cells[j,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[j,2]).Value.ToString();
				int Qty = int.Parse(((Range)Excel_Utilities.ExcelRange.Cells[j,3]).Value.ToString());
				for( int i=1;i<=Qty;i++)
				{
					Devices_Functions.AddDevicesfromGallery(sDeviceName,sType);
				}
				
				float expectedDCUnits = float.Parse(((Range)Excel_Utilities.ExcelRange.Cells[j,4]).Value.ToString());
				
				float maxDCUnits = float.Parse(((Range)Excel_Utilities.ExcelRange.Cells[j,5]).Value.ToString());
				verifyDCUnitsValue(expectedDCUnits.ToString());
				expectedColorCode = Devices_Functions.calculatePercentage(expectedDCUnits, maxDCUnits);
				//repo.ProfileConsys1.cell_ACUnits.Click();
				string actualColour = Devices_Functions.getProgressBarColor("Current (DC Units)");
				Devices_Functions.VerifyPercentage(expectedColorCode,actualColour);
				Common_Functions.clickOnPointsTab();
				Devices_Functions.DeleteAllDevices();
			}
		}
		
		

		/**************************************************************************************************************************************
		 * Function Name: verifyPanelLEDEffectOnDC
		 * Function Details: Verification of DC Units of on changing Panel LED
		 * Parameter/Arguments: Excel sheet name to use and its sheet name
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :Alpesh Dhakad 29/11/2018   Alpesh Dhakad - 30/07/2019 & 21/08/2019 - Updated test scripts as per new build and xpaths
		 * Alpesh Dhakad - 11/01/2021 Updated script as per new UI Changes
		 **************************************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyPanelLEDEffectOnDC(string sFileName,string sPanelLED)
		{
			Excel_Utilities.OpenExcelFile(sFileName,sPanelLED);
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			string expectedDCUnits,DCUnitLoadingDetailName,DCUnitWorstCaseLoadingDetailName,LoopA_Details,LoopB_Details,LoopC_Details;
				
			
			LoopA_Details=((Range)Excel_Utilities.ExcelRange.Cells[2,8]).Value.ToString();
			LoopB_Details=((Range)Excel_Utilities.ExcelRange.Cells[3,8]).Value.ToString();
			LoopC_Details=((Range)Excel_Utilities.ExcelRange.Cells[4,8]).Value.ToString();
			sColumn=((Range)Excel_Utilities.ExcelRange.Cells[2,9]).Value.ToString();
			
			
			for (int i=8; i<=rows;i++)
			{
				int PanelLED;
				string sPanelLEDCount =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				int.TryParse(sPanelLEDCount,out PanelLED);
				
				// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem("Node");
			
				Panel_Functions.changePanelLED(PanelLED);
				
				// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				
			//	repo.ProfileConsys1.tab_PhysicalLayout.Click();
				Common_Functions.clickOnPhysicalLayoutTab();
			
				expectedDCUnits = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				DCUnitLoadingDetailName = ((Range)Excel_Utilities.ExcelRange.Cells[2,6]).Value.ToString();
				DCUnitWorstCaseLoadingDetailName = ((Range)Excel_Utilities.ExcelRange.Cells[3,6]).Value.ToString();

				
				//verifyDCUnitsValue(expectedDCUnits);
				//verifyDCUnitsWorstCaseValue(expectedDCUnits);
				
				//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitLoadingDetailName);
				//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitWorstCaseLoadingDetailName);
				
				Common_Functions.clickOnPanelCalculationsTab();
				
				Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopA_Details,sColumn);
			
				Common_Functions.clickOnPropertiesTab();
				

			}
			Excel_Utilities.CloseExcel();
		}
		
		
		/************************************************************************************************************
		 * Function Name: changeDeviceSensitivityAndVerifyDCUnit
		 * Function Details: To change device sensitivity, day mode and verify DC unit
		 * Parameter/Arguments: fileName, sheetName of add device
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : Alpesh Dhakad - 01/08/2019 & 23/08/2019 - Updated test scripts as per new build and xpaths
		 * Alpesh Dhakad - 12/12/2019 - Updated test scripts with new method for loading details
		 ************************************************************************************************************/
		//Method to verify DC Unit after changing Device Sensitivity and Device Mode
		[UserCodeMethod]
		public static void changeDeviceSensitivityAndVerifyDCUnit(string sFileName,string sAddDevicesSheet)
		{
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			string expectedDCUnits,changeDeviceSensitivity,changeDeviceMode,changeCheckboxState,changeDaySensitivity,changeDayMode,Multichannel,DCUNitLoadingName;
			bool changeCheckboxStateTo,isMultichannel;
			
			for(int i=7; i<=rows; i++)
			{
				ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				string sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				
				sDeviceSensitivity = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				changeDeviceSensitivity = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				sDeviceMode = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				changeDeviceMode = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				
				
				changeCheckboxState = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				
				sDaySensitivity = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				changeDaySensitivity = ((Range)Excel_Utilities.ExcelRange.Cells[i,11]).Value.ToString();
				sDayMode = ((Range)Excel_Utilities.ExcelRange.Cells[i,10]).Value.ToString();
				changeDayMode = ((Range)Excel_Utilities.ExcelRange.Cells[i,12]).Value.ToString();
				
				Multichannel = ((Range)Excel_Utilities.ExcelRange.Cells[i,13]).Value.ToString();
				bool.TryParse(Multichannel, out isMultichannel);
				bool.TryParse(changeCheckboxState, out changeCheckboxStateTo);
				
				string[] splitLabelName  = sLabelName.Split(',');
				if(!isMultichannel)
				{
					// Add devices from the gallery as per test data from the excel sheet
					Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
					
					// Click on Points tab
					Common_Functions.clickOnPointsTab();
					sLabelName = splitLabelName[0];
					
					// Click on Label name for the device
					Devices_Functions.SelectRowUsingLabelName(sLabelName);
					
					// Verify the label name visibility   
					
					if(repo.FormMe.txt_LabelName1Info.Exists())
					{
						// Call VerifyDeviceSensitivity & VerifyDeviceMode method to verify its values
						Devices_Functions.VerifyDeviceSensitivity(sDeviceSensitivity);
						Devices_Functions.VerifyDeviceMode(sDeviceMode);
						
						// Call ChangeDeviceSensitivity & ChangeDeviceMode method to verify its values
						Devices_Functions.ChangeDeviceSensitivity(changeDeviceSensitivity);
						Devices_Functions.ChangeDeviceMode(changeDeviceMode);
						
						// Call CheckUncheckDayMatchesNight to check/uncheck the checkbox and then verify and change its values
						Devices_Functions.CheckUncheckDayMatchesNight(changeCheckboxStateTo);
						
						// Verify and change Day Sensitivity & Day mode
						Devices_Functions.VerifyDaySensitivity(sDaySensitivity);
						Devices_Functions.VerifyDayMode(sDayMode);
						
						if(!changeCheckboxStateTo)
						{
							Devices_Functions.ChangeDayMode(changeDayMode);
							Devices_Functions.ChangeDaySensitivity(changeDaySensitivity);
						}
						
					}
					else
					{
						Report.Log(ReportLevel.Failure, "Device "+sLabelName+" not found");
					}
					
					// Click on Points tab
					Common_Functions.clickOnPointsTab();
					
					// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
				}
				
				else
				{
					// Add devices from the gallery as per test data from the excel sheet
					Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
					
					// Click on Points tab
					Common_Functions.clickOnPointsTab();
					sLabelName = splitLabelName[0];
					
					// Click on Label name for the device
					Devices_Functions.SelectRowUsingLabelName(sLabelName);
					
					Devices_Functions.VerifyDeviceSensitivity(sDeviceSensitivity);
					Devices_Functions.ChangeDeviceSensitivity(changeDeviceSensitivity);
					
					
					// Click on Label name for the device
					sLabelName = splitLabelName[1];
					Devices_Functions.SelectRowUsingLabelName(sLabelName);
					
					Devices_Functions.VerifyDeviceMode(sDeviceMode);
					Devices_Functions.ChangeDeviceMode(changeDeviceMode);
					
					Devices_Functions.CheckUncheckDayMatchesNight(changeCheckboxStateTo);
					
					// Verify and change Day Sensitivity & Day mode
					Devices_Functions.VerifyDaySensitivity(sDaySensitivity);
					Devices_Functions.VerifyDayMode(sDayMode);
					
					if(!changeCheckboxStateTo)
					{
						Devices_Functions.ChangeDayMode(changeDayMode);
						Devices_Functions.ChangeDaySensitivity(changeDaySensitivity);
					}
					
					// Enter the Day Matches night text in Search Properties fields to view day matches night related text;
					repo.ProfileConsys1.txt_SearchProperties.PressKeys("Day Matches Night" +"{ENTER}" );

					// CLick on checkbox cell lower left corner
					repo.ProfileConsys1.PARTItemsPresenter.cell_DayMatchesNight.Click(Location.LowerLeft);
					
					// Click on Day Matches night checkbox
					repo.ProfileConsys1.PARTItemsPresenter.chkbox_DayMatchesNight.Click();
					
					// To retrieve the attribute value as boolean by its ischecked properties and store in actual state
					bool actualState =  repo.ProfileConsys1.PARTItemsPresenter.chkbox_DayMatchesNight.GetAttributeValue<bool>("ischecked");
					
					if(actualState)
					{
						Report.Log(ReportLevel.Success,"User not able to uncheck checkbox and displayed correctly");
					}
					else
					{
						Report.Log(ReportLevel.Failure,"User able to uncheck checkbox and displayed incorrectly");
					}
					
					Devices_Functions.VerifyDayModeField(true);
					Devices_Functions.VerifyDaySensitivityField(true);
					
					//Click on Points tab
					Common_Functions.clickOnPointsTab();
					
					// Click on SearchProperties text field
					repo.ProfileConsys1.txt_SearchProperties.Click();
					
					// Select the text in SearchProperties text field and delete it
					Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
					
					
					
					// Click on Points tab
					Common_Functions.clickOnPointsTab();
					
					// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
					
				}
			}
			//Verify the expected DC units value after changing various properties
			expectedDCUnits= ((Range)Excel_Utilities.ExcelRange.Cells[2,2]).Value.ToString();
			DCUNitLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[2,3]).Value.ToString();
			
			//verifyDCUnitsValue(expectedDCUnits);
			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUNitLoadingName);
			
		}
		
		
		/*****************************************************************************************************************************
		 * Function Name: VerifyDCUnitsAndWorstCaseIndicators
		 * Function Details: To verify DC unit,  worst cases indicators and its color
		 * Parameter/Arguments: fileName, sheetName of add device
		 * Output:
		 * Function Owner: Poonam Kadam
		 * Last Update : Alpesh Dhakad - 23/06/2020 Added EditPoweredValue method with Loop Powered steps as per new implementation
		 * Alpesh Dhakad - 11/01/2021 Updated script as per new UI Changes and new Power calculation update
		 *****************************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyDCUnitsAndWorstCaseIndicators(string sFileName,string sAddDevicesSheet)
		{
			string expectedColorCodeDC, expectedColorCodeWorstCase, sType,sDeviceName,DCUnitLoadingDetailName,DCUnitWorstCaseLoadingDetailName;
			string LoopA_Details,LoopB_Details,LoopC_Details;
			
			
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			DCUnitLoadingDetailName= ((Range)Excel_Utilities.ExcelRange.Cells[2,6]).Value.ToString();
			DCUnitWorstCaseLoadingDetailName= ((Range)Excel_Utilities.ExcelRange.Cells[3,6]).Value.ToString();
			LoopA_Details=((Range)Excel_Utilities.ExcelRange.Cells[2,8]).Value.ToString();
			LoopB_Details=((Range)Excel_Utilities.ExcelRange.Cells[3,8]).Value.ToString();
			LoopC_Details=((Range)Excel_Utilities.ExcelRange.Cells[4,8]).Value.ToString();
			sColumn=((Range)Excel_Utilities.ExcelRange.Cells[2,9]).Value.ToString();
			
			
			
			for(int j=8;j<=rows;j++)
			{
				
				sDeviceName =  ((Range)Excel_Utilities.ExcelRange.Cells[j,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[j,2]).Value.ToString();
				int Qty = int.Parse(((Range)Excel_Utilities.ExcelRange.Cells[j,3]).Value.ToString());
				
				if(j==8)
				{

				for( int i=1;i<=Qty;i++)
				{
					Devices_Functions.AddDevicesfromGallery(sDeviceName,sType);
					
					Devices_Functions.EditPoweredValue("Powered","Loop Powered");
				}
				
				}
				else
				{
					Devices_Functions.AddDevicesfromMultiplePointWizard(sDeviceName,Qty);
				}
				float expectedDCUnits = float.Parse(((Range)Excel_Utilities.ExcelRange.Cells[j,4]).Value.ToString());
				float maxDCUnits = float.Parse(((Range)Excel_Utilities.ExcelRange.Cells[j,5]).Value.ToString());
				float expectedWorstCaseUnits = float.Parse(((Range)Excel_Utilities.ExcelRange.Cells[j,6]).Value.ToString());
				float maxWorstCaseUnits = float.Parse(((Range)Excel_Utilities.ExcelRange.Cells[j,7]).Value.ToString());
				
				//verifyDCUnitsValue(expectedDCUnits.ToString());
				//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits.ToString(),DCUnitLoadingDetailName);

				Common_Functions.clickOnPanelCalculationsTab();
				
				Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits.ToString(),LoopA_Details,sColumn);
			

				Devices_Functions.verifyLoopLoadingDetailColor(LoopA_Details,sColumn);
				
				Common_Functions.clickOnPropertiesTab();
//				//verifyWorstCaseValue
//				expectedColorCodeDC = Devices_Functions.calculatePercentage(expectedDCUnits, maxDCUnits);
//				//expectedColorCodeWorstCase = Devices_Functions.calculatePercentage(expectedWorstCaseUnits, maxWorstCaseUnits);
//				//repo.ProfileConsys1.cell_ACUnits.Click();
//				string actualColourDC = Devices_Functions.getProgressBarColors("Current (DC Units)");
//				
//				
//				//string actualColourWorstCase = Devices_Functions.getProgressBarColors("Current (worst case)");
//				Devices_Functions.VerifyPercentage(expectedColorCodeDC,actualColourDC);
//				//Devices_Functions.VerifyPercentage(expectedColorCodeWorstCase,actualColourWorstCase);
				Common_Functions.clickOnPointsTab();
				//Devices_Functions.DeleteAllDevices();
			}
		}
		
		/***************************************************************************************************************************************************************
		 * Function Name: verifyTripCurrentForDCCalculation
		 * Function Details: To verify trip current DC calculation value by adding devices
		 					and also verify other loop DC value
		 * Parameter/Arguments: fileName, sheetName for Add devices in loop A and add other devices
		 * Output:
		 * Function Owner: Poonam Kadam
		 * Last Update : 30/11/2018  Updated on 22/01/2018 - Alpesh Dhakad Alpesh Dhakad - 30/07/2019 & 21/08/2019 - Updated test scripts as per new build and xpaths
		 *  Alpesh Dhakad - 10/12/2019 - Updated test scripts with new method for loading details
		 * Alpesh Dhakad - 11/01/2021 Updated script as per new UI Changes
 		 **************************************************************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyTripCurrentForDCCalculation(string sFileName, string sAddDevicesLoopA, string sAddOtherDevices)
		{
			// Declared various fields as String type
			string sLabelName,expectedDCUnits,DCUnitLoadingDetailName;
			string LoopA_Details,LoopB_Details,LoopC_Details;
			
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesLoopA);
			
			LoopA_Details=((Range)Excel_Utilities.ExcelRange.Cells[2,10]).Value.ToString();
			LoopB_Details=((Range)Excel_Utilities.ExcelRange.Cells[3,10]).Value.ToString();
			LoopC_Details=((Range)Excel_Utilities.ExcelRange.Cells[4,10]).Value.ToString();
			sColumn=((Range)Excel_Utilities.ExcelRange.Cells[2,11]).Value.ToString();
				
			
			// Count the number of rows in excel
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			//Select Physical Layout tab
			//repo.ProfileConsys1.tab_PhysicalLayout.Click();
			Common_Functions.clickOnPhysicalLayoutTab();
			
			Common_Functions.clickOnPanelCalculationsTab();
			
			//Fetch value from excel sheet and store it
			expectedDCUnits = ((Range)Excel_Utilities.ExcelRange.Cells[4,7]).Value.ToString();
			DCUnitLoadingDetailName= ((Range)Excel_Utilities.ExcelRange.Cells[2,9]).Value.ToString();
			
			//verifyDCUnitsValue(expectedDCUnits);
			//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitLoadingDetailName);
			Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopA_Details,sColumn);
	
			Report.Log(ReportLevel.Info, "Verified Default DC units");
			
			Common_Functions.clickOnPropertiesTab();
			
			//Select Points tab
			Common_Functions.clickOnPointsTab();
			
			// Click on Loop A node
			Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
			int rowNumber=8;
			ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[rowNumber,1]).Value.ToString();
			string sType = ((Range)Excel_Utilities.ExcelRange.Cells[rowNumber,2]).Value.ToString();
			sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[rowNumber,3]).Value.ToString();
			
			sBase = ((Range)Excel_Utilities.ExcelRange.Cells[rowNumber,9]).Value.ToString();
			sRowIndex= ((Range)Excel_Utilities.ExcelRange.Cells[rowNumber,10]).Value.ToString();
			
			// Add devices from the gallery as per test data from the excel sheet
			Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
			
			//Assign Base to devices
			if(sBase!=null && sBase !="NA")
			{
				Devices_Functions.AssignDeviceBase(sLabelName,sBase,sRowIndex);
				Report.Log(ReportLevel.Info, "Base " + sBase + " assigned to "+ "ModelNumber");
			}
			
			// For loop to fetch values from the excel sheet and then add devices
			for(int i=9; i<=rows; i++)
			{
				ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				
				sBase = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				sRowIndex= ((Range)Excel_Utilities.ExcelRange.Cells[i,10]).Value.ToString();
				
				// Add devices from the gallery as per test data from the excel sheet
				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
				
				//Assign base to devices
				if(sBase!=null && sBase !="NA")
				{
					Devices_Functions.AssignDeviceBaseForMultipleDevices(sLabelName,sBase,sRowIndex);
					Report.Log(ReportLevel.Info, "Base " + sBase + " assigned to "+ "ModelNumber");
				}
				
				// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
			}
			//Select Physical Layout tab
			Common_Functions.clickOnPhysicalLayoutTab();
			
			Common_Functions.clickOnPanelCalculationsTab();
			
			//Fetch value from excel sheet and store it
			expectedDCUnits = ((Range)Excel_Utilities.ExcelRange.Cells[1,7]).Value.ToString();
			//verifyDCUnitsValue(expectedDCUnits);
			//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitLoadingDetailName);
			Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopA_Details,sColumn);
	

			Report.Log(ReportLevel.Info, "Verified DC units after adding Devices and Base");
			
			Common_Functions.clickOnPropertiesTab();
			
			//Select Points tab
			Common_Functions.clickOnPointsTab();
			
			// Click on Loop A node
			Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
			//Close excel
			Excel_Utilities.CloseExcel();
			
			//Open Excel sheet
			Excel_Utilities.OpenExcelFile(sFileName,sAddOtherDevices);
			
			// Count the number of rows in excel
			rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			LoopA_Details=((Range)Excel_Utilities.ExcelRange.Cells[2,10]).Value.ToString();
			LoopB_Details=((Range)Excel_Utilities.ExcelRange.Cells[3,10]).Value.ToString();
			LoopC_Details=((Range)Excel_Utilities.ExcelRange.Cells[4,10]).Value.ToString();
			sColumn=((Range)Excel_Utilities.ExcelRange.Cells[2,11]).Value.ToString();
			
			//Select Points tab
			Common_Functions.clickOnPointsTab();
			
			// For loop to fetch values from the excel sheet and then add devices
			for(int i=8; i<=rows; i++)
			{
				ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				
				// Add devices from the gallery as per test data from the excel sheet
				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
				
				// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
				// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
				//Select Points tab
				Common_Functions.clickOnPointsTab();
			
				
			}
			//Select Physical Layout tab
			//repo.ProfileConsys1.tab_PhysicalLayout.Click();
			Common_Functions.clickOnPhysicalLayoutTab();
			
			Common_Functions.clickOnPanelCalculationsTab();
			
			//Fetch value from excel sheet and store it
			String expectedDCUnits1 = ((Range)Excel_Utilities.ExcelRange.Cells[1,7]).Value.ToString();
			DCUnitLoadingDetailName= ((Range)Excel_Utilities.ExcelRange.Cells[2,9]).Value.ToString();
			
			//verifyDCUnitsValue(expectedDCUnits1);
			//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits1,DCUnitLoadingDetailName);
			Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits1,LoopA_Details,sColumn);

			Common_Functions.clickOnPropertiesTab();

			Report.Log(ReportLevel.Info, "Verified DC units after adding other Devices");
			
			//Select Points tab
			Common_Functions.clickOnPointsTab();
			
			for(int i=8; i<=rows; i++)
			{
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				//Delete devices
				Devices_Functions.DeleteDeviceUsingLabel(sLabelName);
			}
			
			//Fetch value from excel sheet and store it
			expectedDCUnits = ((Range)Excel_Utilities.ExcelRange.Cells[2,7]).Value.ToString();
			
			Common_Functions.clickOnPanelCalculationsTab();
			
			
			//verifyDCUnitsValue(expectedDCUnits);
			//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitLoadingDetailName);
			Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopA_Details,sColumn);


			Report.Log(ReportLevel.Info, "Verified DC units after adding other Devices");
			
			//Close excel
			Excel_Utilities.CloseExcel();
			
		}

		
		/**********************************************************************************************************
		 * Function Name: verifyTripCurrentWithMultipleBase()
		 * Function Details: Verify Trip current with changing base of devices
		 * Parameter/Arguments: fileName, sheetName
		 * Output:
		 * Function Owner: Devendra Kulkarni
		 * Last Update : 30/11/2018
		 * Alpesh Dhakad - 10/12/2019 - Updated test scripts with new method for loading details
		 * Alpesh Dhakad - 12/01/2021 Updated script as per new UI Changes
		 **********************************************************************************************************/
		[UserCodeMethod]
		public static void verifyTripCurrentWithMultipleBase(string fileName, string sheetName)
		{
			// Declared various fields as String type
			string sLabelName, expectedDCUnits, sType,DCUnitLoadingDetailName,DCUnitWorstCaseLoadingDetailName;
			string LoopA_Details,LoopB_Details,LoopC_Details;
			
			Excel_Utilities.OpenExcelFile(fileName,sheetName);
			
			LoopA_Details=((Range)Excel_Utilities.ExcelRange.Cells[2,10]).Value.ToString();
			LoopB_Details=((Range)Excel_Utilities.ExcelRange.Cells[3,10]).Value.ToString();
			LoopC_Details=((Range)Excel_Utilities.ExcelRange.Cells[4,10]).Value.ToString();
			sColumn=((Range)Excel_Utilities.ExcelRange.Cells[2,11]).Value.ToString();
			
			// Count the number of rows in excel
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			Report.Log(ReportLevel.Info, "No of rows: "+rows);
			
			for (int i=8; i<=rows; i++)
			{
				ModelNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				
				// Add devices from the gallery as per test data from the excel sheet
				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
				Report.Log(ReportLevel.Info, "Device "+ModelNumber+" added to Panel");
			}
			
			//Fetch value from excel sheet and store it
			expectedDCUnits = ((Range)Excel_Utilities.ExcelRange.Cells[1,7]).Value.ToString();
			DCUnitLoadingDetailName= ((Range)Excel_Utilities.ExcelRange.Cells[2,9]).Value.ToString();
			DCUnitWorstCaseLoadingDetailName= ((Range)Excel_Utilities.ExcelRange.Cells[3,9]).Value.ToString();
			
			//verifyDCUnitsValue(expectedDCUnits);
			//verifyDCUnitsWorstCaseValue(expectedDCUnits);
			
			//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitLoadingDetailName);
			//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitWorstCaseLoadingDetailName);
			
			Common_Functions.clickOnPanelCalculationsTab();
			
			Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopA_Details,sColumn);

			
			Report.Log(ReportLevel.Info, "Verified default DC units.");
			
			Common_Functions.clickOnPropertiesTab();
			
			//Select Points tab
			Common_Functions.clickOnPointsTab();
			
			sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[9,3]).Value.ToString();
			sBase = ((Range)Excel_Utilities.ExcelRange.Cells[9,9]).Value.ToString();
			sRowIndex= ((Range)Excel_Utilities.ExcelRange.Cells[9,10]).Value.ToString();
			Devices_Functions.AssignDeviceBaseForMultipleDevices(sLabelName,sBase,sRowIndex);
			Report.Log(ReportLevel.Info, "Base " + sBase + " assigned to "+ sLabelName);
			
			//Fetch value from excel sheet and store it
			expectedDCUnits = ((Range)Excel_Utilities.ExcelRange.Cells[5,7]).Value.ToString();
			//verifyDCUnitsValue(expectedDCUnits);
			//verifyDCUnitsWorstCaseValue(expectedDCUnits);
			
			//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitLoadingDetailName);
			//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitWorstCaseLoadingDetailName);

			Common_Functions.clickOnPanelCalculationsTab();
			
			Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopA_Details,sColumn);

			
			Report.Log(ReportLevel.Info, "Verified DC units changing base.");
			
			Common_Functions.clickOnPropertiesTab();
			
			//Select Points tab
			Common_Functions.clickOnPointsTab();
			
			sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[10,3]).Value.ToString();
			sRowIndex= ((Range)Excel_Utilities.ExcelRange.Cells[10,10]).Value.ToString();
			Devices_Functions.RemoveBase(sLabelName, sRowIndex);

			//Fetch value from excel sheet and store it
			expectedDCUnits = ((Range)Excel_Utilities.ExcelRange.Cells[2,7]).Value.ToString();
			//verifyDCUnitsValue(expectedDCUnits);
			//verifyDCUnitsWorstCaseValue(expectedDCUnits);
			
			//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitLoadingDetailName);
			//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitWorstCaseLoadingDetailName);
			
			Common_Functions.clickOnPanelCalculationsTab();
			
			Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopA_Details,sColumn);

			
			Report.Log(ReportLevel.Info, "Verified DC units after deleting base.");
			
			//Close excel
			Excel_Utilities.CloseExcel();

		}
		
		/********************************************************************
		 * Function Name: verifyDCUnitsAfterReopen()
		 * Function Details: This function verified DC units after reopening of project
		 *                   to ensure data saved correctly.
		 * Parameter/Arguments: fileName, sheetName, expectedDCUnits
		 * Output:
		 * Function Owner: Devendra Kulkarni
		 * Last Update : 30/11/2018
		 ********************************************************************/
		[UserCodeMethod]
		public static void verifyDCUnitsAfterReopen(string fileName, string sheetName, string expectedDCUnits)
		{
			
			Excel_Utilities.OpenExcelFile(fileName,sheetName);
			
			//Select Physical Layout tab
			//repo.ProfileConsys1.tab_PhysicalLayout.Click();
			Common_Functions.clickOnPhysicalLayoutTab();
			

			string ActualDcUnits = repo.FormMe.Current_DC_Units.TextValue;
			string ActualWorstCaseCurrent = repo.FormMe.CurrentWrstCase.TextValue;
			
			if((ActualDcUnits.Equals(expectedDCUnits)) && (ActualWorstCaseCurrent.Equals(expectedDCUnits)))
			{
				Report.Log(ReportLevel.Success,"DC Units " + ActualDcUnits + " displayed correclty");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"DC Units are not displayed correctly, DC Units displayed as: " +ActualDcUnits +" instead of : "+expectedDCUnits);
			}
			Report.Log(ReportLevel.Info, "Verified DC units after reopening application.");
			
			//Close excel
			Excel_Utilities.CloseExcel();
		}

		/****************************************************************************************************************************************
		 * Function Name: verifyTripCurrentWithMultipleLoop()
		 * Function Details: This function verifies trip current with devices
		 * 					 connected in Loop A and Loop B
		 * Parameter/Arguments: fileName, sheetNameA, sheetNameB
		 * Output:
		 * Function Owner: Devendra Kulkarni
		 * Last Update : 30/11/2018   Alpesh Dhakad - 30/07/2019 & 21/08/2019 - Updated test scripts as per new build and xpaths
		 * Alpesh Dhakad - 10/12/2019 - Updated test scripts with new method for loading details
		 * Alpesh Dhakad - 12/01/2021 Updated script as per new UI Changes
		 ****************************************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyTripCurrentWithMultipleLoop(string fileName, string sheetNameA, string sheetNameB)
		{
			string expectedDCUnits,DCUnitLoadingDetailName,DCUnitWorstCaseLoadingDetailName;
			string LoopA_Details,LoopB_Details,LoopC_Details;
			
			AddDevicesFromExcel(fileName, sheetNameA);
			Excel_Utilities.OpenExcelFile(fileName,sheetNameA);
			//Fetch value from excel sheet and store it
			expectedDCUnits = ((Range)Excel_Utilities.ExcelRange.Cells[1,7]).Value.ToString();
			DCUnitLoadingDetailName= ((Range)Excel_Utilities.ExcelRange.Cells[2,9]).Value.ToString();
			DCUnitWorstCaseLoadingDetailName= ((Range)Excel_Utilities.ExcelRange.Cells[3,9]).Value.ToString();
			
			LoopA_Details=((Range)Excel_Utilities.ExcelRange.Cells[2,10]).Value.ToString();
			LoopB_Details=((Range)Excel_Utilities.ExcelRange.Cells[3,10]).Value.ToString();
			LoopC_Details=((Range)Excel_Utilities.ExcelRange.Cells[4,10]).Value.ToString();
			sColumn=((Range)Excel_Utilities.ExcelRange.Cells[2,11]).Value.ToString();
			
			//verifyDCUnitsValue(expectedDCUnits);
			//verifyDCUnitsWorstCaseValue(expectedDCUnits);
			
			//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitLoadingDetailName);
			//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitWorstCaseLoadingDetailName);

			Common_Functions.clickOnPanelCalculationsTab();
			
			Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopA_Details,sColumn);

			
			Report.Log(ReportLevel.Info, "Verified default DC units.");
			
			Common_Functions.clickOnPropertiesTab();
			
			//Close excel
			Excel_Utilities.CloseExcel();
			
			//Select Points tab
			Common_Functions.clickOnPointsTab();
			
			// Click on Loop B node
			Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-B");
					
			AddDevicesFromExcel(fileName, sheetNameB);
			
			Excel_Utilities.OpenExcelFile(fileName,sheetNameB);
			
			//Fetch value from excel sheet and store it
			expectedDCUnits = ((Range)Excel_Utilities.ExcelRange.Cells[2,7]).Value.ToString();
			DCUnitLoadingDetailName= ((Range)Excel_Utilities.ExcelRange.Cells[2,9]).Value.ToString();
			DCUnitWorstCaseLoadingDetailName= ((Range)Excel_Utilities.ExcelRange.Cells[3,9]).Value.ToString();
			
			LoopA_Details=((Range)Excel_Utilities.ExcelRange.Cells[2,10]).Value.ToString();
			LoopB_Details=((Range)Excel_Utilities.ExcelRange.Cells[3,10]).Value.ToString();
			LoopC_Details=((Range)Excel_Utilities.ExcelRange.Cells[4,10]).Value.ToString();
			sColumn=((Range)Excel_Utilities.ExcelRange.Cells[2,11]).Value.ToString();
			
			//verifyDCUnitsValue(expectedDCUnits);
			//verifyDCUnitsWorstCaseValue(expectedDCUnits);
			//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitLoadingDetailName);
			//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitWorstCaseLoadingDetailName);

			Common_Functions.clickOnPanelCalculationsTab();
			
			Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopB_Details,sColumn);

			Report.Log(ReportLevel.Info, "Verified DC units after adding devices in Loop B.");
			
			Common_Functions.clickOnPropertiesTab();
			
			//Select Points tab
			Common_Functions.clickOnPointsTab();
			
			// Remove base from Loop B
			sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[9,3]).Value.ToString();
			sRowIndex= ((Range)Excel_Utilities.ExcelRange.Cells[9,10]).Value.ToString();
			Devices_Functions.RemoveBase(sLabelName, sRowIndex);
			
			//Fetch value from excel sheet and store it
			expectedDCUnits = ((Range)Excel_Utilities.ExcelRange.Cells[3,7]).Value.ToString();
			//verifyDCUnitsValue(expectedDCUnits);
			//verifyDCUnitsWorstCaseValue(expectedDCUnits);
			
			//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitLoadingDetailName);
			//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitWorstCaseLoadingDetailName);

			Common_Functions.clickOnPanelCalculationsTab();
			
			Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopB_Details,sColumn);

			
			Report.Log(ReportLevel.Info, "Verified DC units after removing base in Loop B.");

			//Close excel
			Excel_Utilities.CloseExcel();
		}

		/********************************************************************
		 * Function Name: AddDevicesFromExcel()
		 * Function Details: This function adds devices from Excel sheet mentioned
		 * Parameter/Arguments: fileName, sheetName
		 * Output: None
		 * Function Owner: Devendra Kulkarni
		 * Last Update : 30/11/2018
		 ********************************************************************/
		[UserCodeMethod]
		public static void AddDevicesFromExcel(string fileName, string sheetName)
		{
			Excel_Utilities.OpenExcelFile(fileName,sheetName);
			string sType;
			
			// Count the number of rows in excel
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			for (int i=8; i<=rows; i++)
			{
				ModelNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				
				// Add devices from the gallery as per test data from the excel sheet
				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
				Report.Log(ReportLevel.Info, "Device "+ModelNumber+" added to Panel");
			}
			
			//Close excel
			Excel_Utilities.CloseExcel();
		}

		/****************************************************************************************************************
		 * Function Name: VerifyDCCalculationOnAddingDevices()
		 * Function Details: To verify DC calculation on adding devices on Loop A and Loop B
		 * Parameter/Arguments: sFileName, sAddDevicesLoopA, sAddDevicesLoopB
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : Alpesh Dhakad - 30/07/2019 & 21/08/2019 - Updated test scripts as per new build and xpaths
		 * Alpesh Dhakad - 06/12/2019 - Updated test scripts with new method for loading details
		 * Alpesh Dhakad - 11/01/2021 Updated script as per new UI Changes
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyDCCalculationOnAddingDevices(string sFileName, string sAddDevicesLoopA, string sAddDevicesLoopB)
		{
			//Add devies in loop A,
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesLoopA);
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			string expectedDCUnits, sType, sLabelName,DCUnitLoadingDetailName,DCUnitWorstCaseLoadingDetailName;
			string LoopA_Details,LoopB_Details,LoopC_Details;
			
			LoopA_Details=((Range)Excel_Utilities.ExcelRange.Cells[2,8]).Value.ToString();
			LoopB_Details=((Range)Excel_Utilities.ExcelRange.Cells[3,8]).Value.ToString();
			LoopC_Details=((Range)Excel_Utilities.ExcelRange.Cells[4,8]).Value.ToString();
			sColumn=((Range)Excel_Utilities.ExcelRange.Cells[2,9]).Value.ToString();
			
			
			
			for(int i=8; i<=rows; i++)
			{
				ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				sBase = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				//sRowIndex= ((Range)Excel_Utilities.ExcelRange.Cells[i,10]).Value.ToString();
				
				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
				
			}
			
			//Verify DC Units of Loop A
			Report.Log(ReportLevel.Info,"Verification of DC Units of Loop A on addition of devices in Loop A");
			expectedDCUnits= ((Range)Excel_Utilities.ExcelRange.Cells[2,2]).Value.ToString();
			DCUnitLoadingDetailName= ((Range)Excel_Utilities.ExcelRange.Cells[2,3]).Value.ToString();
			DCUnitWorstCaseLoadingDetailName= ((Range)Excel_Utilities.ExcelRange.Cells[3,3]).Value.ToString();
			
			//verifyDCUnitsValue(expectedDCUnits);
			//verifyDCUnitsWorstCaseValue(expectedDCUnits);
			
			//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitLoadingDetailName);
			//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitWorstCaseLoadingDetailName);
			
			Common_Functions.clickOnPanelCalculationsTab();
				
			Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopA_Details,sColumn);
			
			Common_Functions.clickOnPropertiesTab();


			
			//Verify DC Units of Loop B
			Report.Log(ReportLevel.Info,"Verification of DC Units of Loop B on addition of devices in Loop A");
			
			// Click on Loop B node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-B");
				 
			
			expectedDCUnits= ((Range)Excel_Utilities.ExcelRange.Cells[3,2]).Value.ToString();
			//verifyDCUnitsValue(expectedDCUnits);
			//verifyDCUnitsWorstCaseValue(expectedDCUnits);
			
			//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitLoadingDetailName);
			//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitWorstCaseLoadingDetailName);
			
			
			Common_Functions.clickOnPanelCalculationsTab();
				
			Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopB_Details,sColumn);
			
			Common_Functions.clickOnPropertiesTab();
			
			Excel_Utilities.CloseExcel();
			
			//Add devices in loop B
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesLoopB);
			rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			LoopA_Details=((Range)Excel_Utilities.ExcelRange.Cells[2,8]).Value.ToString();
			LoopB_Details=((Range)Excel_Utilities.ExcelRange.Cells[3,8]).Value.ToString();
			LoopC_Details=((Range)Excel_Utilities.ExcelRange.Cells[4,8]).Value.ToString();
			sColumn=((Range)Excel_Utilities.ExcelRange.Cells[2,9]).Value.ToString();
			
			
			
			// Click on Loop B node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-B");
			
			
			for(int i=8; i<=rows; i++)
			{
				ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				sBase = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				//sRowIndex= ((Range)Excel_Utilities.ExcelRange.Cells[i,10]).Value.ToString();
				
				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
				
			}
			
			//Verify DC Units of Loop B
			Report.Log(ReportLevel.Info,"Verification of DC Units of Loop B on addition of devices in Loop B");
			expectedDCUnits= ((Range)Excel_Utilities.ExcelRange.Cells[3,2]).Value.ToString();
			DCUnitLoadingDetailName= ((Range)Excel_Utilities.ExcelRange.Cells[2,3]).Value.ToString();
			DCUnitWorstCaseLoadingDetailName= ((Range)Excel_Utilities.ExcelRange.Cells[3,3]).Value.ToString();
			
			
			
			//verifyDCUnitsValue(expectedDCUnits);
			//verifyDCUnitsWorstCaseValue(expectedDCUnits);
			
			//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitLoadingDetailName);
			//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitWorstCaseLoadingDetailName);

			Common_Functions.clickOnPanelCalculationsTab();
				
			Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopB_Details,sColumn);
			
			Common_Functions.clickOnPropertiesTab();
			
			
			//Verify DC Units of Loop A
			// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				 
				
			// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
							
			
			Report.Log(ReportLevel.Info,"Verification of DC Units of Loop A on addition of devices in Loop B");
			expectedDCUnits= ((Range)Excel_Utilities.ExcelRange.Cells[2,2]).Value.ToString();
			//verifyDCUnitsValue(expectedDCUnits);
			//verifyDCUnitsWorstCaseValue(expectedDCUnits);
			
			//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitLoadingDetailName);
			//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitWorstCaseLoadingDetailName);

			Common_Functions.clickOnPanelCalculationsTab();
				
			Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopA_Details,sColumn);
			
			Common_Functions.clickOnPropertiesTab();
			
			Excel_Utilities.CloseExcel();

		}
		
		
		/****************************************************************************************************************
		 * Function Name: VerifyDCCalculationOnChangingBase
		 * Function Details:  Verify DC calculation on changing base of device and adding sounder with base
		 * Parameter/Arguments: sFileName, sAddDevicesLoopA, sAddSounderBaseDevices
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 19/12/2018
		 ****************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyDCCalculationOnChangingBase(string sFileName, string sAddDevicesLoopA, string sAddSounderBaseDevices)
		{
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesLoopA);
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			string sType, sLabelName;
			
			for(int i=8; i<=rows; i++)
			{
				ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				sBase = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				sRowIndex= ((Range)Excel_Utilities.ExcelRange.Cells[i,10]).Value.ToString();
				
				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
				
			}
			sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[10,3]).Value.ToString();
			sBase = ((Range)Excel_Utilities.ExcelRange.Cells[10,9]).Value.ToString();
			sRowIndex= ((Range)Excel_Utilities.ExcelRange.Cells[10,10]).Value.ToString();
			Devices_Functions.AssignDeviceBase(sLabelName,sBase,sRowIndex);
			
			Excel_Utilities.CloseExcel();
			
			// Add Sounder Base devices
			Excel_Utilities.OpenExcelFile(sFileName,sAddSounderBaseDevices);
			rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			for(int i=8; i<=rows; i++)
			{
				ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				sBase = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				sRowIndex= ((Range)Excel_Utilities.ExcelRange.Cells[i,10]).Value.ToString();
				
				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
				
			}
			Excel_Utilities.CloseExcel();
			
		}
		
		/*****************************************************************************************************************
		 * Function Name:  VerifyCurrentDCUnitscalculation
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Purvi Bhasim
		 * Last Update : 08/01/2019 Alpesh Dhakad - 30/07/2019 & 21/08/2019 - Updated scripts as per new build and xpaths
		 * Alpesh Dhakad - 06/12/2019 - Updated test scripts with new method for loading details
		 * Alpesh Dhakad - 15/05/2020 Updated script as per new implementation changes
		 * Alpesh Dhakad - 11/01/2021 Updated script as per new UI Changes
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyCurrentDCUnitscalculation(string sFileName,string sAddPanelSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared variables
			string ModelNumber,sType,sLabelName,sAssignedBase,expectedDCUnits,DefaultDCUnits,ChangedDCUnit,sPanelLEDCount,LoadingDetailsName;
			string LoopA_Details,LoopB_Details,LoopC_Details;
			int PanelLED;
			
			LoopA_Details=((Range)Excel_Utilities.ExcelRange.Cells[2,6]).Value.ToString();
				LoopB_Details=((Range)Excel_Utilities.ExcelRange.Cells[3,6]).Value.ToString();
				LoopC_Details=((Range)Excel_Utilities.ExcelRange.Cells[4,6]).Value.ToString();
				sColumn=((Range)Excel_Utilities.ExcelRange.Cells[2,7]).Value.ToString();
			
			PanelLED=0;
			ChangedDCUnit=string.Empty;
			expectedDCUnits=string.Empty;
			DefaultDCUnits=string.Empty;
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				
				ModelNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				sAssignedBase = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sRowIndex= ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				expectedDCUnits = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				DefaultDCUnits = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				sPanelLEDCount = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				ChangedDCUnit = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				LoadingDetailsName = ((Range)Excel_Utilities.ExcelRange.Cells[i,10]).Value.ToString();
				

				
				int.TryParse(sPanelLEDCount,out PanelLED);
				
				// Click on Expander node
				Common_Functions.ClickOnNavigationTreeExpander("Node");
				
				// Click on Loop Card node
				//Common_Functions.ClickOnNavigationTreeExpander("PFI");
				
				// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				 
				
				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
				
				//Assign Base to the device
				Devices_Functions.AssignDeviceBase(sLabelName,sAssignedBase,sRowIndex);
				
				Common_Functions.clickOnPanelCalculationsTab();
				
				//Verify Default DC Units
				//verifyDCUnitsValue(expectedDCUnits);
				//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,LoadingDetailsName);
				Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopA_Details,sColumn);
			
				Common_Functions.clickOnPropertiesTab();

				// Click on Site node
				Common_Functions.ClickOnNavigationTreeItem("Site");
				
			}
			// Click on Loop A node
			Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
			//Click on  points tab
			Common_Functions.clickOnPointsTab();
			
			Keyboard.Press("{LControlKey down}{Akey}{LControlKey up}");
			
			//Copy Devices
			repo.FormMe.btn_Copy.Click();
			
			// Click on Loop C node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-C");
				
				//Click on  points tab
			Common_Functions.clickOnPointsTab();
			
			//Paste the devices
			Common_Functions.clickOnPasteButton();
			
			Common_Functions.clickOnPanelCalculationsTab();
			//Verify DC Units
			//verifyDCUnitsValue(expectedDCUnits);
			//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,"Current (DC Units)");
			Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopC_Details,sColumn);
			
			Common_Functions.clickOnPropertiesTab();
			
			// Click on Site node
			Common_Functions.ClickOnNavigationTreeItem("Site");
				
			
			// Click on Loop C node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-C");
				
			
			//go to points grid
			Common_Functions.clickOnPointsTab();
			
			Keyboard.Press("{LControlKey down}{Akey}{LControlKey up}");
			
			//Copy Devices
			Common_Functions.clickOnCutButton();
			
			Common_Functions.clickOnPanelCalculationsTab();
			
			//Verify Default DC Units
			//verifyDCUnitsValue(DefaultDCUnits);
			//Devices_Functions.verifyLoadingDetailsValue(DefaultDCUnits,"Current (DC Units)");
			Devices_Functions.verifyLoopLoadingDetailsValue(DefaultDCUnits,LoopC_Details,sColumn);
			
			Common_Functions.clickOnPropertiesTab();
			
			
			
			// Click on Site node
				Common_Functions.ClickOnNavigationTreeItem("Site");
				
			
			// Click on Panel node
				//Common_Functions.ClickOnNavigationTreeExpander("Node");
				
			
			Panel_Functions.changePanelLED(PanelLED);
			
			
			// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				Common_Functions.clickOnPanelCalculationsTab();
				
			//Verify Default DC Units
			//verifyDCUnitsValue(ChangedDCUnit);
			//Devices_Functions.verifyLoadingDetailsValue(ChangedDCUnit,"Current (DC Units)");
			Devices_Functions.verifyLoopLoadingDetailsValue(ChangedDCUnit,LoopA_Details,sColumn);
			
			Common_Functions.clickOnPropertiesTab();
			

		}
		
		/*************************************************************************************************************************
		 * Function Name: verifyTripCurrentCalculationForFIMLoopFC
		 * Function Details: Verify Trip current with changing base of devices on both the loops
		 * Parameter/Arguments: fileName, loopAdevices,loopBdevices
		 * Output:
		 * Function Owner: Alpesh Dhakad 
		 * Last Update : 11/09/2019 Alpesh Dhakad - 12/12/2019 - Updated test scripts with new method for loading details
		 * Alpesh Dhakad - 12/01/2021 Updated script as per new UI Changes
		 *************************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyTripCurrentCalculationForFIMLoopFC(string fileName, string loopAdevices, string loopBdevices)
		{
			// Declared various fields as String type
			string sLabelName, expectedDCUnits, sType,DCUnitLoopLoadingName,DCUnitWorstCaseLoopLoadingName;
			string LoopA_Details,LoopB_Details,LoopC_Details;
			
			Excel_Utilities.OpenExcelFile(fileName,loopAdevices);
			
			LoopA_Details=((Range)Excel_Utilities.ExcelRange.Cells[2,8]).Value.ToString();
			LoopB_Details=((Range)Excel_Utilities.ExcelRange.Cells[3,8]).Value.ToString();
			LoopC_Details=((Range)Excel_Utilities.ExcelRange.Cells[4,8]).Value.ToString();
			sColumn=((Range)Excel_Utilities.ExcelRange.Cells[2,9]).Value.ToString();
				
						
			// Count the number of rows in excel
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			Report.Log(ReportLevel.Info, "No of rows: "+rows);
			
			for (int i=8; i<=rows; i++)
			{
				ModelNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				
				// Add devices from the gallery as per test data from the excel sheet
				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
				
			}
			
			//Fetch value from excel sheet and store it
			expectedDCUnits = ((Range)Excel_Utilities.ExcelRange.Cells[1,7]).Value.ToString();
			DCUnitLoopLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[2,5]).Value.ToString();
			DCUnitWorstCaseLoopLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[3,5]).Value.ToString();
			
			Common_Functions.clickOnPanelCalculationsTab();
			
			//verifyDCUnitsValue(expectedDCUnits);
			//verifyDCUnitsWorstCaseValue(expectedDCUnits);
//			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitLoopLoadingName);
//			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitWorstCaseLoopLoadingName);
	
			Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopA_Details,sColumn);

			Report.Log(ReportLevel.Info, "Verified default DC units.");
			
			Common_Functions.clickOnPropertiesTab();
			
			//Select Points tab
			Common_Functions.clickOnPointsTab();
			
			sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[9,3]).Value.ToString();
			sBase = ((Range)Excel_Utilities.ExcelRange.Cells[9,9]).Value.ToString();
			sRowIndex= ((Range)Excel_Utilities.ExcelRange.Cells[9,10]).Value.ToString();
			Devices_Functions.AssignDeviceBaseForMultipleDevices(sLabelName,sBase,sRowIndex);
			Report.Log(ReportLevel.Info, "Base " + sBase + " assigned to "+ sLabelName);
			
			//Fetch value from excel sheet and store it
			expectedDCUnits = ((Range)Excel_Utilities.ExcelRange.Cells[5,7]).Value.ToString();
			//verifyDCUnitsValue(expectedDCUnits);
			//verifyDCUnitsWorstCaseValue(expectedDCUnits);
			
			Common_Functions.clickOnPanelCalculationsTab();
			
//			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitLoopLoadingName);
//			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitWorstCaseLoopLoadingName);
	
			Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopA_Details,sColumn);

			
			Report.Log(ReportLevel.Info, "Verified DC units changing base.");
			
			Common_Functions.clickOnPropertiesTab();
			
			//Select Points tab
			Common_Functions.clickOnPointsTab();
			
			
			//Close excel
			Excel_Utilities.CloseExcel();
			
			//Select Points tab
			Common_Functions.clickOnPointsTab();
			
			// Click on Loop B node
			Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-B");
					
			AddDevicesFromExcel(fileName, loopBdevices);
			
			Excel_Utilities.OpenExcelFile(fileName,loopBdevices);
			
			//Fetch value from excel sheet and store it
			expectedDCUnits = ((Range)Excel_Utilities.ExcelRange.Cells[2,7]).Value.ToString();
			DCUnitLoopLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[2,5]).Value.ToString();
			DCUnitWorstCaseLoopLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[3,5]).Value.ToString();
			
			LoopA_Details=((Range)Excel_Utilities.ExcelRange.Cells[2,8]).Value.ToString();
			LoopB_Details=((Range)Excel_Utilities.ExcelRange.Cells[3,8]).Value.ToString();
			LoopC_Details=((Range)Excel_Utilities.ExcelRange.Cells[4,8]).Value.ToString();
			sColumn=((Range)Excel_Utilities.ExcelRange.Cells[2,9]).Value.ToString();
				
			//verifyDCUnitsValue(expectedDCUnits);
			//verifyDCUnitsWorstCaseValue(expectedDCUnits);
			
//			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitLoopLoadingName);
//			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitWorstCaseLoopLoadingName);
			Common_Functions.clickOnPanelCalculationsTab();
			
			Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopB_Details,sColumn);

			
			Report.Log(ReportLevel.Info, "Verified DC units after adding devices in Loop B.");
			
			Common_Functions.clickOnPropertiesTab();
			
			//Select Points tab
			Common_Functions.clickOnPointsTab();
			
			sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[10,3]).Value.ToString();
			sBase = ((Range)Excel_Utilities.ExcelRange.Cells[10,9]).Value.ToString();
			sRowIndex= ((Range)Excel_Utilities.ExcelRange.Cells[10,10]).Value.ToString();
			Devices_Functions.AssignDeviceBaseForMultipleDevices(sLabelName,sBase,sRowIndex);
			Report.Log(ReportLevel.Info, "Base " + sBase + " assigned to "+ sLabelName);
			
			//Fetch value from excel sheet and store it
			expectedDCUnits = ((Range)Excel_Utilities.ExcelRange.Cells[5,7]).Value.ToString();
			//verifyDCUnitsValue(expectedDCUnits);
			//verifyDCUnitsWorstCaseValue(expectedDCUnits);
			
//			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitLoopLoadingName);
//			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitWorstCaseLoopLoadingName);
	
			Common_Functions.clickOnPanelCalculationsTab();
			
			Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopB_Details,sColumn);

			
			Report.Log(ReportLevel.Info, "Verified DC units changing base.");
			
			Common_Functions.clickOnPropertiesTab();
			
			// Remove base from Loop B
			sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[10,3]).Value.ToString();
			sRowIndex= ((Range)Excel_Utilities.ExcelRange.Cells[10,10]).Value.ToString();
			Devices_Functions.RemoveBase(sLabelName, sRowIndex);
			
			//Fetch value from excel sheet and store it
			expectedDCUnits = ((Range)Excel_Utilities.ExcelRange.Cells[3,7]).Value.ToString();
			//verifyDCUnitsValue(expectedDCUnits);
			//verifyDCUnitsWorstCaseValue(expectedDCUnits);
			
//			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitLoopLoadingName);
//			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitWorstCaseLoopLoadingName);
			
			Common_Functions.clickOnPanelCalculationsTab();
			
			Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopB_Details,sColumn);

	
			Report.Log(ReportLevel.Info, "Verified DC units after removing base in Loop B.");
			
			Common_Functions.clickOnPropertiesTab();
			
			//Select Points tab
			Common_Functions.clickOnPointsTab();
			
			// Click on Loop A node
			Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
			
			//Fetch value from excel sheet and store it
			expectedDCUnits = ((Range)Excel_Utilities.ExcelRange.Cells[1,7]).Value.ToString();
			//verifyDCUnitsValue(expectedDCUnits);
			//verifyDCUnitsWorstCaseValue(expectedDCUnits);
			
//			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitLoopLoadingName);
//			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitWorstCaseLoopLoadingName);
	
			Common_Functions.clickOnPanelCalculationsTab();
			
			Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopA_Details,sColumn);

			
			Report.Log(ReportLevel.Info, "Verified DC units in Loop A.");
			
			

			//Close excel
			Excel_Utilities.CloseExcel();

			
		}
		
		
		/***************************************************************************************************************************************************************
		 * Function Name: verifyTripCurrentOnAdditionDeletionOfIsolatorBase
		 * Function Details: To verify trip current DC calculation value by adding devices
		 					and also verify other loop DC value
		 * Parameter/Arguments: fileName, sheetName for Add devices in loop A and add other devices
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 11/09/2019 Alpesh Dhakad - 12/12/2019 - Updated test scripts with new method for loading details
		 * Alpesh Dhakad - 12/01/2021 Updated script as per new UI Changes
		 ***************************************************************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyTripCurrentOnAdditionDeletionOfIsolatorBase(string sFileName, string sAddDevicesLoopA, string sOtherDevices)
		{
			// Declared various fields as String type
			string sLabelName,expectedDCUnits,DCUnitLoopLoadingName,DCUnitWorstCaseLoopLoadingName;
			string LoopA_Details,LoopB_Details,LoopC_Details;
			
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesLoopA);
			
			// Count the number of rows in excel
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			LoopA_Details=((Range)Excel_Utilities.ExcelRange.Cells[2,8]).Value.ToString();
			LoopB_Details=((Range)Excel_Utilities.ExcelRange.Cells[3,8]).Value.ToString();
			LoopC_Details=((Range)Excel_Utilities.ExcelRange.Cells[4,8]).Value.ToString();
			sColumn=((Range)Excel_Utilities.ExcelRange.Cells[2,9]).Value.ToString();
				
			
			//Select Physical Layout tab
			//repo.ProfileConsys1.tab_PhysicalLayout.Click();
			Common_Functions.clickOnPhysicalLayoutTab();
			
			Common_Functions.clickOnPanelCalculationsTab();
			
			//Fetch value from excel sheet and store it
			expectedDCUnits = ((Range)Excel_Utilities.ExcelRange.Cells[4,7]).Value.ToString();
			DCUnitLoopLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[2,5]).Value.ToString();
			DCUnitWorstCaseLoopLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[3,5]).Value.ToString();
	
			//verifyDCUnitsValue(expectedDCUnits);
			//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitLoopLoadingName);
			Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopA_Details,sColumn);

			
			Report.Log(ReportLevel.Info, "Verified Default DC units");
			
			Common_Functions.clickOnPropertiesTab();
			
			//Select Points tab
			Common_Functions.clickOnPointsTab();
			
			// Click on Loop A node
			Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
			int rowNumber=8;
			ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[rowNumber,1]).Value.ToString();
			string sType = ((Range)Excel_Utilities.ExcelRange.Cells[rowNumber,2]).Value.ToString();
			sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[rowNumber,3]).Value.ToString();
			
			sBase = ((Range)Excel_Utilities.ExcelRange.Cells[rowNumber,9]).Value.ToString();
			sRowIndex= ((Range)Excel_Utilities.ExcelRange.Cells[rowNumber,10]).Value.ToString();
			
			// Add devices from the gallery as per test data from the excel sheet
			Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
			
			//Assign Base to devices
			if(sBase!=null && sBase !="NA")
			{
				Devices_Functions.AssignDeviceBase(sLabelName,sBase,sRowIndex);
				Report.Log(ReportLevel.Info, "Base " + sBase + " assigned to "+ "ModelNumber");
			}
			
			// For loop to fetch values from the excel sheet and then add devices
			for(int i=9; i<=rows; i++)
			{
				ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				
				sBase = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				sRowIndex= ((Range)Excel_Utilities.ExcelRange.Cells[i,10]).Value.ToString();
				
				// Add devices from the gallery as per test data from the excel sheet
				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
				
				//Assign base to devices
				if(sBase!=null && sBase !="NA")
				{
					Devices_Functions.AssignDeviceBaseForMultipleDevices(sLabelName,sBase,sRowIndex);
					Report.Log(ReportLevel.Info, "Base " + sBase + " assigned to "+ "ModelNumber");
				}
				
				// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
			}
			//Select Physical Layout tab
			//repo.ProfileConsys1.tab_PhysicalLayout.Click();
			Common_Functions.clickOnPhysicalLayoutTab();
			
			
			//Fetch value from excel sheet and store it
			expectedDCUnits = ((Range)Excel_Utilities.ExcelRange.Cells[1,7]).Value.ToString();
			DCUnitLoopLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[2,5]).Value.ToString();
			DCUnitWorstCaseLoopLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[3,5]).Value.ToString();
			
			Common_Functions.clickOnPanelCalculationsTab();
	
			//verifyDCUnitsValue(expectedDCUnits);
			//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitLoopLoadingName);
			Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopA_Details,sColumn);

			Report.Log(ReportLevel.Info, "Verified DC units after adding Devices and Base");
			
			Common_Functions.clickOnPropertiesTab();
			
			//Select Points tab
			Common_Functions.clickOnPointsTab();
			
			// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
			//Close excel
			Excel_Utilities.CloseExcel();
			
			Excel_Utilities.OpenExcelFile(sFileName,sOtherDevices);
			
			//Select Points tab
			Common_Functions.clickOnPointsTab();
			
			LoopA_Details=((Range)Excel_Utilities.ExcelRange.Cells[2,8]).Value.ToString();
			LoopB_Details=((Range)Excel_Utilities.ExcelRange.Cells[3,8]).Value.ToString();
			LoopC_Details=((Range)Excel_Utilities.ExcelRange.Cells[4,8]).Value.ToString();
			sColumn=((Range)Excel_Utilities.ExcelRange.Cells[2,9]).Value.ToString();
				
			
			
			sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[8,3]).Value.ToString();
			sRowIndex= ((Range)Excel_Utilities.ExcelRange.Cells[8,10]).Value.ToString();
			Devices_Functions.RemoveBase(sLabelName, sRowIndex);

			//Fetch value from excel sheet and store it
			expectedDCUnits = ((Range)Excel_Utilities.ExcelRange.Cells[1,7]).Value.ToString();
			DCUnitLoopLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[2,5]).Value.ToString();
			DCUnitWorstCaseLoopLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[3,5]).Value.ToString();
	
			
			//verifyDCUnitsValue(expectedDCUnits);
			//verifyDCUnitsWorstCaseValue(expectedDCUnits);
			
//			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitLoopLoadingName);
//			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitWorstCaseLoopLoadingName);
			
			Common_Functions.clickOnPanelCalculationsTab();
	
			Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopA_Details,sColumn);

			
			
			Report.Log(ReportLevel.Info, "Verified DC units after deleting base.");
			
			Common_Functions.clickOnPropertiesTab();
			
			//Select Points tab
			Common_Functions.clickOnPointsTab();
			
			sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[9,3]).Value.ToString();
			sRowIndex= ((Range)Excel_Utilities.ExcelRange.Cells[9,10]).Value.ToString();
			Devices_Functions.RemoveBase(sLabelName, sRowIndex);
			
			
			//Fetch value from excel sheet and store it
			expectedDCUnits = ((Range)Excel_Utilities.ExcelRange.Cells[2,7]).Value.ToString();
			//verifyDCUnitsValue(expectedDCUnits);
			//verifyDCUnitsWorstCaseValue(expectedDCUnits);
			
//			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitLoopLoadingName);
//			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitWorstCaseLoopLoadingName);
			
			Common_Functions.clickOnPanelCalculationsTab();
	
			Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopA_Details,sColumn);

			Report.Log(ReportLevel.Info, "Verified DC units after deleting base.");
			
			
			
			//Close excel
			Excel_Utilities.CloseExcel();
		}
		
		/********************************************************************
		 * Function Name: verifyDCUnitsValueAfterReopen
		 * Function Details: Expected DC Units value
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 12/09/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void verifyDCUnitsValueAfterReopen(string expectedDCUnits)
		{
			//repo.ProfileConsys1.tab_PhysicalLayout.Click();
			Common_Functions.clickOnPhysicalLayoutTab();
			
			string ActualDcUnits = repo.FormMe.DCUnits_AfterReopen.TextValue;
			
			if(ActualDcUnits.Equals(expectedDCUnits))
			{
				Report.Log(ReportLevel.Success,"DC Units " + ActualDcUnits + " is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"DC Units is not displayed correctly, DC Units displayed as: " +ActualDcUnits + " instead of : "+expectedDCUnits);
			}
			Common_Functions.clickOnPointsTab();
		}
		
		/********************************************************************
		 * Function Name: verifyDCUnitsWorstCaseValueAfterReopen
		 * Function Details: Verify DC units worst case value
		 * Parameter/Arguments: expected DC units worst case value
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 12/09/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void verifyDCUnitsWorstCaseValueAfterReopen(string expectedWorstCaseDCUnits)
		{
			//repo.ProfileConsys1.tab_PhysicalLayout.Click();
			Common_Functions.clickOnPhysicalLayoutTab();
			
			string ActualDcUnitsWorstCase = repo.FormMe.DCUnitsWorstCase_AfterReopen.TextValue;
			
			if(ActualDcUnitsWorstCase.Equals(expectedWorstCaseDCUnits))
			{
				Report.Log(ReportLevel.Success,"DC Units worst case value " + ActualDcUnitsWorstCase + " is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"DC Units worst case value is not displayed correctly, DC Units displayed as: " +ActualDcUnitsWorstCase + " instead of : "+expectedWorstCaseDCUnits);
			}
			Common_Functions.clickOnPointsTab();
		}
	
		/****************************************************************************************************************
		 * Function Name: VerifyCurrentDCCalculation
		 * Function Details:  Verify current DC calculation on changing base of device
		 * Parameter/Arguments: sFileName, sAddDevicesLoopA, sAddSounderBaseDevices
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 12/09/2019 Alpesh Dhakad - 12/12/2019 - Updated test scripts with new method for loading details
		 * Alpesh Dhakad - 12/01/2021 Updated script as per new UI Changes
		 ****************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyCurrentDCCalculation(string sFileName, string sAddDevicesLoopA, string sOtherDevicesLoopA)
		{
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesLoopA);
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			string sType, sLabelName,expectedDCUnits,DCUnitLoopLoadingName,DCUnitWorstCaseLoopLoadingName;
			string LoopA_Details,LoopB_Details,LoopC_Details;
			
			LoopA_Details=((Range)Excel_Utilities.ExcelRange.Cells[2,10]).Value.ToString();
			LoopB_Details=((Range)Excel_Utilities.ExcelRange.Cells[3,10]).Value.ToString();
			LoopC_Details=((Range)Excel_Utilities.ExcelRange.Cells[4,10]).Value.ToString();
			sColumn=((Range)Excel_Utilities.ExcelRange.Cells[2,11]).Value.ToString();
	
			
			
			for(int i=8; i<=rows; i++)
			{
				ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				sBase = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				sRowIndex= ((Range)Excel_Utilities.ExcelRange.Cells[i,10]).Value.ToString();
				
				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
				
			}
			sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[10,3]).Value.ToString();
			sBase = ((Range)Excel_Utilities.ExcelRange.Cells[10,9]).Value.ToString();
			sRowIndex= ((Range)Excel_Utilities.ExcelRange.Cells[10,10]).Value.ToString();
			Devices_Functions.AssignDeviceBase(sLabelName,sBase,sRowIndex);
			
			//Select Physical Layout tab
			//repo.ProfileConsys1.tab_PhysicalLayout.Click();
			Common_Functions.clickOnPhysicalLayoutTab();
			
			//Fetch value from excel sheet and store it
			expectedDCUnits = ((Range)Excel_Utilities.ExcelRange.Cells[1,7]).Value.ToString();
			DCUnitLoopLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[2,5]).Value.ToString();
			DCUnitWorstCaseLoopLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[3,5]).Value.ToString();
			
			Common_Functions.clickOnPanelCalculationsTab();
			
			//verifyDCUnitsValue(expectedDCUnits);
			//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitLoopLoadingName);
			Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopA_Details,sColumn);
			
			Report.Log(ReportLevel.Info, "Verified DC units after adding Devices and Base");
			
			Common_Functions.clickOnPropertiesTab();
			
			Excel_Utilities.CloseExcel();
			
			Excel_Utilities.OpenExcelFile(sFileName,sOtherDevicesLoopA);
			
			LoopA_Details=((Range)Excel_Utilities.ExcelRange.Cells[2,10]).Value.ToString();
			LoopB_Details=((Range)Excel_Utilities.ExcelRange.Cells[3,10]).Value.ToString();
			LoopC_Details=((Range)Excel_Utilities.ExcelRange.Cells[4,10]).Value.ToString();
			sColumn=((Range)Excel_Utilities.ExcelRange.Cells[2,11]).Value.ToString();
	
			
			//Select Points tab
			Common_Functions.clickOnPointsTab();
			
			sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[8,3]).Value.ToString();
			
			Devices_Functions.SelectRowUsingLabelName(sLabelName);
			
			Common_Functions.clickOnDeleteButton();
			
						
			//Fetch value from excel sheet and store it
			expectedDCUnits = ((Range)Excel_Utilities.ExcelRange.Cells[2,7]).Value.ToString();
			DCUnitLoopLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[2,5]).Value.ToString();
			DCUnitWorstCaseLoopLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[3,5]).Value.ToString();
			
			//verifyDCUnitsValue(expectedDCUnits);
			//verifyDCUnitsWorstCaseValue(expectedDCUnits);
			
//			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitLoopLoadingName);
//			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitWorstCaseLoopLoadingName);
//			
			Common_Functions.clickOnPanelCalculationsTab();
			
			//verifyDCUnitsValue(expectedDCUnits);
			Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopA_Details,sColumn);
			
			Common_Functions.clickOnPropertiesTab();
			
			
			Excel_Utilities.CloseExcel();
			
			Excel_Utilities.OpenExcelFile(sFileName,sOtherDevicesLoopA);
			
			LoopA_Details=((Range)Excel_Utilities.ExcelRange.Cells[2,10]).Value.ToString();
			LoopB_Details=((Range)Excel_Utilities.ExcelRange.Cells[3,10]).Value.ToString();
			LoopC_Details=((Range)Excel_Utilities.ExcelRange.Cells[4,10]).Value.ToString();
			sColumn=((Range)Excel_Utilities.ExcelRange.Cells[2,11]).Value.ToString();
	
			
			sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[9,3]).Value.ToString();
			sRowIndex= ((Range)Excel_Utilities.ExcelRange.Cells[9,10]).Value.ToString();
			Devices_Functions.RemoveBase(sLabelName, sRowIndex);

			//Fetch value from excel sheet and store it
			expectedDCUnits = ((Range)Excel_Utilities.ExcelRange.Cells[3,7]).Value.ToString();
			DCUnitLoopLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[2,5]).Value.ToString();
			DCUnitWorstCaseLoopLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[3,5]).Value.ToString();
			
			
			//verifyDCUnitsValue(expectedDCUnits);
			//verifyDCUnitsWorstCaseValue(expectedDCUnits);
			
//			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitLoopLoadingName);
//			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitWorstCaseLoopLoadingName);
			
			Common_Functions.clickOnPanelCalculationsTab();
			
			Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopA_Details,sColumn);

			Report.Log(ReportLevel.Info, "Verified DC units after deleting base.");
			
			Common_Functions.clickOnPropertiesTab();
		
			Excel_Utilities.CloseExcel();
			
			Excel_Utilities.OpenExcelFile(sFileName,sOtherDevicesLoopA);
			
			LoopA_Details=((Range)Excel_Utilities.ExcelRange.Cells[2,10]).Value.ToString();
			LoopB_Details=((Range)Excel_Utilities.ExcelRange.Cells[3,10]).Value.ToString();
			LoopC_Details=((Range)Excel_Utilities.ExcelRange.Cells[4,10]).Value.ToString();
			sColumn=((Range)Excel_Utilities.ExcelRange.Cells[2,11]).Value.ToString();
	
			
			//Select Points tab
			Common_Functions.clickOnPointsTab();
			
			sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[10,3]).Value.ToString();
			
			Devices_Functions.SelectRowUsingLabelName(sLabelName);
			
			Common_Functions.clickOnDeleteButton();
			
			//Fetch value from excel sheet and store it
			expectedDCUnits = ((Range)Excel_Utilities.ExcelRange.Cells[4,7]).Value.ToString();
			DCUnitLoopLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[2,5]).Value.ToString();
			DCUnitWorstCaseLoopLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[3,5]).Value.ToString();
			
			
			//verifyDCUnitsValue(expectedDCUnits);
			//verifyDCUnitsWorstCaseValue(expectedDCUnits);
			
			//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitLoopLoadingName);
			//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitWorstCaseLoopLoadingName);
			
			Common_Functions.clickOnPanelCalculationsTab();
			
			Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopA_Details,sColumn);

			
			Common_Functions.clickOnPropertiesTab();
		
			
			
			Excel_Utilities.CloseExcel();
			
		}
		
		/***********************************************************************************************************************
		 * Function Name: verifyDCUnitCalculationForNonSplitLoops
		 * Function Details: Verify DC Units for Non Split loops
		 * Parameter/Arguments: fileName, loopAdevices,loopBdevices
		 * Output:
		 * Function Owner: Alpesh Dhakad 
		 * Last Update : 13/09/2019 Alpesh Dhakad - 12/12/2019 - Updated test scripts with new method for loading details
		 * Alpesh Dhakad - 12/01/2021 Updated script as per new UI Changes
		 ***********************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyDCUnitCalculationForNonSplitLoops(string fileName, string loopAdevices, string loopBdevices)
		{
			// Declared various fields as String type
			string expectedDCUnits, sType,DCUnitLoopLoadingName,DCUnitWorstCaseLoopLoadingName;
			string LoopA_Details,LoopB_Details,LoopC_Details;
	
			Excel_Utilities.OpenExcelFile(fileName,loopAdevices);
			
			// Count the number of rows in excel
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			Report.Log(ReportLevel.Info, "No of rows: "+rows);
			
			LoopA_Details=((Range)Excel_Utilities.ExcelRange.Cells[2,8]).Value.ToString();
			LoopB_Details=((Range)Excel_Utilities.ExcelRange.Cells[3,8]).Value.ToString();
			LoopC_Details=((Range)Excel_Utilities.ExcelRange.Cells[4,8]).Value.ToString();
			sColumn=((Range)Excel_Utilities.ExcelRange.Cells[2,9]).Value.ToString();

			// Click on Loop A node
			Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
			
			for (int i=8; i<=rows; i++)
			{
				ModelNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				
				// Add devices from the gallery as per test data from the excel sheet
				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
				
			}
			
			//Fetch value from excel sheet and store it
			expectedDCUnits = ((Range)Excel_Utilities.ExcelRange.Cells[1,7]).Value.ToString();
			DCUnitLoopLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[2,5]).Value.ToString();
			DCUnitWorstCaseLoopLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[3,5]).Value.ToString();
			
			
			//verifyDCUnitsValue(expectedDCUnits);
			//verifyDCUnitsWorstCaseValue(expectedDCUnits);
			
//			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitLoopLoadingName);
//			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitWorstCaseLoopLoadingName);
//			
			Common_Functions.clickOnPanelCalculationsTab();
			
			Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopA_Details,sColumn);

			
			Report.Log(ReportLevel.Info, "Verified DC units for Loop A.");
			
			Common_Functions.clickOnPropertiesTab();
			
			//Select Points tab
			Common_Functions.clickOnPointsTab();
			
			// Click on Loop B node
			Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-B");
			
			//Click on Physical Layout tab
			//repo.ProfileConsys1.tab_PhysicalLayout.Click();
			Common_Functions.clickOnPhysicalLayoutTab();
			
			//Fetch value from excel sheet and store it
			expectedDCUnits = ((Range)Excel_Utilities.ExcelRange.Cells[2,7]).Value.ToString();
			DCUnitLoopLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[2,5]).Value.ToString();
			DCUnitWorstCaseLoopLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[3,5]).Value.ToString();
			
			
			//verifyDCUnitsValue(expectedDCUnits);
			//verifyDCUnitsWorstCaseValue(expectedDCUnits);
			
			
//			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitLoopLoadingName);
//			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitWorstCaseLoopLoadingName);
			
			Common_Functions.clickOnPanelCalculationsTab();
			
			Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopB_Details,sColumn);

			
			Report.Log(ReportLevel.Info, "Verified DC units for Loop B");
			
			
			Common_Functions.clickOnPropertiesTab();
			
			
			
			//Close excel
			Excel_Utilities.CloseExcel();
			
			//Click on Points tab
			Common_Functions.clickOnPointsTab();
			
			// Click on Loop B node
			Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-B");
					
			AddDevicesFromExcel(fileName, loopBdevices);
			
			Excel_Utilities.OpenExcelFile(fileName,loopBdevices);
			
			//Fetch value from excel sheet and store it
			expectedDCUnits = ((Range)Excel_Utilities.ExcelRange.Cells[1,7]).Value.ToString();
			DCUnitLoopLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[2,5]).Value.ToString();
			DCUnitWorstCaseLoopLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[3,5]).Value.ToString();
			
			LoopA_Details=((Range)Excel_Utilities.ExcelRange.Cells[2,8]).Value.ToString();
			LoopB_Details=((Range)Excel_Utilities.ExcelRange.Cells[3,8]).Value.ToString();
			LoopC_Details=((Range)Excel_Utilities.ExcelRange.Cells[4,8]).Value.ToString();
			sColumn=((Range)Excel_Utilities.ExcelRange.Cells[2,9]).Value.ToString();

			
			//verifyDCUnitsValue(expectedDCUnits);
			//verifyDCUnitsWorstCaseValue(expectedDCUnits);
			
//			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitLoopLoadingName);
//			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitWorstCaseLoopLoadingName);
			
			
			Common_Functions.clickOnPanelCalculationsTab();
			
			Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopB_Details,sColumn);

			
			Report.Log(ReportLevel.Info, "Verified DC units after adding devices in Loop B.");
			
			
			Common_Functions.clickOnPropertiesTab();
			
			
			//Select Points tab
			Common_Functions.clickOnPointsTab();
			
			// Click on Loop A node
			Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
			
			//Click on Physical Layout tab
			//repo.ProfileConsys1.tab_PhysicalLayout.Click();
			Common_Functions.clickOnPhysicalLayoutTab();
			
			
			//Fetch value from excel sheet and store it
			expectedDCUnits = ((Range)Excel_Utilities.ExcelRange.Cells[2,7]).Value.ToString();
			DCUnitLoopLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[2,5]).Value.ToString();
			DCUnitWorstCaseLoopLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[3,5]).Value.ToString();
			
			
			//verifyDCUnitsValue(expectedDCUnits);
			//verifyDCUnitsWorstCaseValue(expectedDCUnits);
//			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitLoopLoadingName);
//			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitWorstCaseLoopLoadingName);

			
			Common_Functions.clickOnPanelCalculationsTab();
			
			Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits,LoopA_Details,sColumn);

			
			Report.Log(ReportLevel.Info, "Verified DC units for Loop A");
			
			
			Common_Functions.clickOnPropertiesTab();
			
			

			//Close excel
			Excel_Utilities.CloseExcel();

			
		}
		
		/************************************************************************************************************************************
		 * Function Name: VerifyDCUnitsAndWorstCaseIndicators
		 * Function Details: To verify DC unit,  worst cases indicators and its color
		 * Parameter/Arguments: fileName, sheetName of add device
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : Alpesh Dhakad - 12/12/2019 - Updated test scripts with new method for loading details
		 * Alpesh Dhakad - 12/01/2021 Updated script as per new UI Changes
		 ************************************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyDCUnitsAndWorstCaseColorAndValueIndicators(string sFileName,string sAddDevicesSheet)
		{
			string expectedColorCodeDC, expectedColorCodeWorstCase, sType,sDeviceName,DCUnitLoopLoadingName,DCUnitWorstCaseLoopLoadingName;
			string LoopA_Details,LoopB_Details,LoopC_Details;
		
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			for(int j=8;j<=rows;j++)
			{
				sDeviceName =  ((Range)Excel_Utilities.ExcelRange.Cells[j,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[j,2]).Value.ToString();
				//sDeviceQty = ((Range)Excel_Utilities.ExcelRange.Cells[j,3]).Value.ToString();
				int Qty = int.Parse(((Range)Excel_Utilities.ExcelRange.Cells[j,3]).Value.ToString());
				
				LoopA_Details=((Range)Excel_Utilities.ExcelRange.Cells[2,6]).Value.ToString();
				LoopB_Details=((Range)Excel_Utilities.ExcelRange.Cells[3,6]).Value.ToString();
				LoopC_Details=((Range)Excel_Utilities.ExcelRange.Cells[4,6]).Value.ToString();
				sColumn=((Range)Excel_Utilities.ExcelRange.Cells[2,7]).Value.ToString();
					
				
				Devices_Functions.AddDevicesfromMultiplePointWizard(sDeviceName,Qty);
				
				float expectedDCUnits = float.Parse(((Range)Excel_Utilities.ExcelRange.Cells[j,4]).Value.ToString());
				float maxDCUnits = float.Parse(((Range)Excel_Utilities.ExcelRange.Cells[j,5]).Value.ToString());
				float expectedWorstCaseUnits = float.Parse(((Range)Excel_Utilities.ExcelRange.Cells[j,6]).Value.ToString());
				float maxWorstCaseUnits = float.Parse(((Range)Excel_Utilities.ExcelRange.Cells[j,7]).Value.ToString());
				
				DCUnitLoopLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[2,5]).Value.ToString();
				DCUnitWorstCaseLoopLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[3,5]).Value.ToString();
				
				Common_Functions.clickOnPanelCalculationsTab();
				
				//verifyDCUnitsValue(expectedDCUnits.ToString());
				//Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits.ToString(),DCUnitLoopLoadingName);
				Devices_Functions.verifyLoopLoadingDetailsValue(expectedDCUnits.ToString(),LoopA_Details,sColumn);

				
				Devices_Functions.verifyLoopLoadingDetailColor(LoopA_Details,sColumn);
				
				Common_Functions.clickOnPropertiesTab();
				
				//verifyWorstCaseValue
//				expectedColorCodeDC = Devices_Functions.calculatePercentage(expectedDCUnits, maxDCUnits);
//				expectedColorCodeWorstCase = Devices_Functions.calculatePercentage(expectedWorstCaseUnits, maxWorstCaseUnits);
//				//repo.ProfileConsys1.cell_ACUnits.Click();
//				string actualColourDC = Devices_Functions.getProgressBarColors(DCUnitLoopLoadingName);
//				string actualColourWorstCase = Devices_Functions.getProgressBarColors(DCUnitWorstCaseLoopLoadingName);
//				Devices_Functions.VerifyPercentage(expectedColorCodeDC,actualColourDC);
//				Devices_Functions.VerifyPercentage(expectedColorCodeWorstCase,actualColourWorstCase);
				Common_Functions.clickOnPointsTab();
				
			}
		}
		
		/***********************************************************************************************************************
		 * Function Name: verifyDCUnitCalculationForSplitLoops
		 * Function Details: Verify DC Units for Split loops
		 * Parameter/Arguments: fileName, loopAdevices,loopBdevices
		 * Output:
		 * Function Owner: Alpesh Dhakad 
		 * Last Update : 31/08/2020
		 ***********************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyDCUnitCalculationForSplitLoops(string fileName, string loopAdevices, string loopBdevices)
		{
			// Declared various fields as String type
			string expectedDCUnits, sType,DCUnitLoopLoadingName,DCUnitWorstCaseLoopLoadingName;
			
			Excel_Utilities.OpenExcelFile(fileName,loopAdevices);
			
			// Count the number of rows in excel
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			Report.Log(ReportLevel.Info, "No of rows: "+rows);
			
			// Click on Loop A node
			Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
			
			for (int i=8; i<=rows; i++)
			{
				ModelNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				
				// Add devices from the gallery as per test data from the excel sheet
				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
				
			}
			
			//Fetch value from excel sheet and store it
			expectedDCUnits = ((Range)Excel_Utilities.ExcelRange.Cells[1,7]).Value.ToString();
			DCUnitLoopLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[2,5]).Value.ToString();
			DCUnitWorstCaseLoopLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[3,5]).Value.ToString();
			
			
			//verifyDCUnitsValue(expectedDCUnits);
			//verifyDCUnitsWorstCaseValue(expectedDCUnits);
			
			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitLoopLoadingName);
			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitWorstCaseLoopLoadingName);
			
			
			Report.Log(ReportLevel.Info, "Verified DC units for Loop A.");
			
			//Select Points tab
			Common_Functions.clickOnPointsTab();
			
			// Click on Loop B node
			Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-C");
			
			//Click on Physical Layout tab
			//repo.ProfileConsys1.tab_PhysicalLayout.Click();
			Common_Functions.clickOnPhysicalLayoutTab();
			
			//Fetch value from excel sheet and store it
			expectedDCUnits = ((Range)Excel_Utilities.ExcelRange.Cells[2,7]).Value.ToString();
			DCUnitLoopLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[2,5]).Value.ToString();
			DCUnitWorstCaseLoopLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[3,5]).Value.ToString();
			
			
			//verifyDCUnitsValue(expectedDCUnits);
			//verifyDCUnitsWorstCaseValue(expectedDCUnits);
			
			
			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitLoopLoadingName);
			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitWorstCaseLoopLoadingName);
			Report.Log(ReportLevel.Info, "Verified DC units for Loop C");
			
			//Close excel
			Excel_Utilities.CloseExcel();
			
			//Click on Points tab
			Common_Functions.clickOnPointsTab();
			
			// Click on Loop B node
			Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-C");
					
			AddDevicesFromExcel(fileName, loopBdevices);
			
			Excel_Utilities.OpenExcelFile(fileName,loopBdevices);
			
			//Fetch value from excel sheet and store it
			expectedDCUnits = ((Range)Excel_Utilities.ExcelRange.Cells[1,7]).Value.ToString();
			DCUnitLoopLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[2,5]).Value.ToString();
			DCUnitWorstCaseLoopLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[3,5]).Value.ToString();
			
			
			//verifyDCUnitsValue(expectedDCUnits);
			//verifyDCUnitsWorstCaseValue(expectedDCUnits);
			
			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitLoopLoadingName);
			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitWorstCaseLoopLoadingName);
			Report.Log(ReportLevel.Info, "Verified DC units after adding devices in Loop C.");
			
			//Select Points tab
			Common_Functions.clickOnPointsTab();
			
			// Click on Loop A node
			Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
			
			//Click on Physical Layout tab
			//repo.ProfileConsys1.tab_PhysicalLayout.Click();
			Common_Functions.clickOnPhysicalLayoutTab();
			
			
			//Fetch value from excel sheet and store it
			expectedDCUnits = ((Range)Excel_Utilities.ExcelRange.Cells[2,7]).Value.ToString();
			DCUnitLoopLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[2,5]).Value.ToString();
			DCUnitWorstCaseLoopLoadingName = ((Range)Excel_Utilities.ExcelRange.Cells[3,5]).Value.ToString();
			
			
			//verifyDCUnitsValue(expectedDCUnits);
			//verifyDCUnitsWorstCaseValue(expectedDCUnits);
			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitLoopLoadingName);
			Devices_Functions.verifyLoadingDetailsValue(expectedDCUnits,DCUnitWorstCaseLoopLoadingName);
			Report.Log(ReportLevel.Info, "Verified DC units for Loop A");

			//Close excel
			Excel_Utilities.CloseExcel();

			
		}
		
		
		/***********************************************************************************************************************
		 * Function Name: 
		 * Function Details: 
		 * Parameter/Arguments: 
		 * Output:
		 * Function Owner: Alpesh Dhakad 
		 * Last Update : 03/02/2021
		 ***********************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyHPLCalc(string fileName, string loopAdevices)
		{
			Excel_Utilities.OpenExcelFile(fileName,loopAdevices);
			
			
			repo.FormMe.btn_AllGalleryDropdown.Click();
			
			repo.Export.ButtonOK.Click();
			
			repo.AddANewPanel.SecondPSU_txt1Info.Exists();
			
			
		
			
		}
		/***********************************************************************************************************************
		 * Function Name: 
		 * Function Details: 
		 * Parameter/Arguments: fileName,sAddDeviceSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad 
		 * Last Update : ]
		 ***********************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyMTPanelPowerCalculations(string fileName, string sAddDeviceSheet)
		{
			
			ModelNumber=sAddDeviceSheet;

			repo.FormMe.btn_AllGalleryDropdown.Click();
					
			string[] array2 = { "801 CH", "801 CH Car Park", "801 H" };	
			
			
			IList<Ranorex.Text> list = repo.ContextMenu.ListItemsTextInfo.CreateAdapters<Ranorex.Text>();

			
			int i=0;			
			foreach (Ranorex.Text test in list)
			{
				if(test.TextValue.Equals(array2[i]))
				{
					Report.Info(test.TextValue + "matching");
					i++;
				}
				else
				{
					Report.Info(test.TextValue + "Not matching");
					i++;
				}
				
			}	
		
		}
		
		/********************************************************************
		 * Function Name: verifyMaxDCUnitsforMT2
		 * Function Details: Verify maximum DC unit value for MT2 Panel
		 * Parameter/Arguments: Expected Maximum DC unit value
		 * Output:
		 * Function Owner: Rohan Pawar
		 * Last Update : 18/05/2021
		 ********************************************************************/
		[UserCodeMethod]
		public static void verifyMaxDCUnitsforMT2(string expectedMaxDCUnits)
		{
			//repo.ProfileConsys1.tab_PhysicalLayout.Click();
			//Common_Functions.clickOnPhysicalLayoutTab();
			
			string maxDCUnits = repo.FormMe.RP_MaxDC.TextValue;
			
			if(maxDCUnits.Equals(expectedMaxDCUnits))
			{
				Report.Log(ReportLevel.Success,"Max DC Units " + maxDCUnits + " is displayed correctly " );
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Max DC Units is not displayed correctly, it is displayed as: " + maxDCUnits + " instead of : " +expectedMaxDCUnits);
			}
			
			Common_Functions.clickOnPointsTab();
			
			
		}	
	}
}
