/*
 * Created by Ranorex
 * User: jbhosash
 * Date: 8/27/2018
 * Time: 3:20 PM
 * 
 * To change this template use Tools > Options > Coding > Edit standard headers.
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Threading;
using WinForms = System.Windows.Forms;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace TestProject.Libraries
{
	
	[UserCodeCollection]
	public class Devices_Functions
	{
		
		//Create instance of repository to access repository items
		static NGConsysRepository repo = NGConsysRepository.Instance;
		
		static string ModelNumber
		{
			
			get { return repo.ModelNumber; }
			set { repo.ModelNumber = value; }
		}
		
		static string sDeviceOrderRow
		{
			get { return repo.sDeviceOrderRow; }
			set { repo.sDeviceOrderRow = value; }
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
		
		static string sDeviceIndex
		{
			get { return repo.sDeviceIndex; }
			set { repo.sDeviceIndex = value; }
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
		
		static string sMainProcessorGalleryIndex
		{
			get { return repo.sMainProcessorGalleryIndex; }
			set { repo.sMainProcessorGalleryIndex = value; }
		}
		
		static string sRBusGalleryIndex
		{
			get { return repo.sRBusGalleryIndex; }
			set { repo.sRBusGalleryIndex = value; }
		}
		
		static string sDeviceName
		{
			get { return repo.sDeviceName; }
			set { repo.sDeviceName = value; }
		}
		
		static string sAccessoriesGalleryIndex
		{
			get { return repo.sAccessoriesGalleryIndex; }
			set { repo.sAccessoriesGalleryIndex = value; }
		}
		
		static string sDeviceOrderName
		{
			get { return repo.sDeviceOrderName; }
			set { repo.sDeviceOrderName = value; }
		}
		
		static string sListIndex
		{
			get { return repo.sListIndex; }
			set { repo.sListIndex = value; }
		}
		
		static string sColumn
		{
			get { return repo.sColumn; }
			set { repo.sColumn = value; }
		}
		
		static string sPhysicalLayoutDeviceIndex
		{
			get { return repo.sPhysicalLayoutDeviceIndex; }
			set { repo.sPhysicalLayoutDeviceIndex = value; }
		}
		
		
		static string sDeviceAddress
		{
			get { return repo.sDeviceAddress; }
			set { repo.sDeviceAddress = value; }
		}
		
		static string sOtherSlotCardName
		{
			get { return repo.sOtherSlotCardName; }
			set { repo.sOtherSlotCardName = value; }
		}
		
		/********************************************************************
		 * Function Name: AddDevices
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update : 29/03/2019 Alpesh Dhakad- Updated btn_MultiplePointWizard xpath and change script accordingly
		 ********************************************************************/
		[UserCodeMethod]
		public static void AddDevices(string sFileName,string sSheetName)
		{
			
			Excel_Utilities.OpenExcelFile(sFileName,sSheetName);
			//Excel_Utilities.OpenSheet(sSheetName);
			int rows = Excel_Utilities.ExcelRange.Rows.Count;
			
			repo.FormMe.btn_MultiplePointWizard.Click();
			//repo.ProfileConsys1.btn_MultiplePointWizard_DoNotUse.Click();
			repo.AddDevices.txt_AllDevices.Click();
			
			
			for(int i=8;i<=rows;i++)
			{
				string sDeviceName=  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				repo.AddDevices.txt_SearchDevices.Click();
				Keyboard.Press("{LControlKey down}{Akey}{LControlKey up}"+sDeviceName);
				ModelNumber = sDeviceName;
				repo.AddDevices.txt_ModelNumber.Click();
				
			}
			
			repo.AddDevices.btn_AddDevices.Click();
			Delay.Milliseconds(200);
			for(int i=8;i<=rows;i++)
			{
				string sDeviceName = ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				ModelNumber = sDeviceName;
				sRow=(i-7).ToString();
				repo.ProfileConsys1.PanelInvetoryGrid.InventoryGridCell.Click();
				Validate.AttributeEqual(repo.ProfileConsys1.PanelInvetoryGrid.txt_DeviceNameInfo, "Text", sDeviceName);
				Delay.Milliseconds(100);
			}
			
			
			
			Excel_Utilities.ExcelWB.Close(false, null, null);
			//Excel_Utilities.ExcelAppl.Quit();
		}
		
		
		/********************************************************************
		 * Function Name: AddDevicesfromGallery
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void AddDevicesfromGallery(string sDeviceName,string sType)
		{
			sGalleryIndex = SelectGalleryType(sType);
			ModelNumber=sDeviceName;
			repo.ProfileConsys1.btn_DevicesGalleryDropDown.Click();
			repo.ContextMenu.txt_SelectDevice.Click();
			Report.Log(ReportLevel.Success, "Device "+sDeviceName+" Added Successfully");
		}
		
		/********************************************************************
		 * Function Name: AddDevicesfromGalleryNotHavingImages
		 * Function Details: Select devices from gallery using caption displayed for image
		 * Parameter/Arguments: Device name(Model Number) and type of gallery
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Created on :19/2/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void AddDevicesfromGalleryNotHavingImages(string sDeviceName,string sType)
		{
			sGalleryIndex = SelectGalleryType(sType);
			ModelNumber=sDeviceName;
			repo.ProfileConsys1.btn_DevicesGalleryDropDown.Click();
			repo.ContextMenu.txt_galleryItem.Click();
		}
		
		/********************************************************************
		 * Function Name: verifyDevicesfromGalleryNotHavingImages
		 * Function Details: verify Given device exist in gallery
		 * Parameter/Arguments: Device name(Model Number), type of gallery, Visibility
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Created on :11/3/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void verifyDevicesfromGalleryNotHavingImages(string sDeviceName,string sType,bool Visibility)
		{
			sGalleryIndex = SelectGalleryType(sType);
			ModelNumber=sDeviceName;
			repo.ProfileConsys1.btn_DevicesGalleryDropDown.Click();
			if(Visibility)
			{
				if(repo.ContextMenu.txt_galleryItemInfo.Exists())
				{
					Report.Log(ReportLevel.Success, "Device "+sDeviceName+" exist in gallery");
				}
				else
				{
					Report.Log(ReportLevel.Failure, "Device "+sDeviceName+" not exist in gallery");
				}
			}
			else
			{
				if(repo.ContextMenu.txt_galleryItemInfo.Exists())
				{
					Report.Log(ReportLevel.Failure, "Device "+sDeviceName+" exist in gallery");
				}
				else
				{
					Report.Log(ReportLevel.Success, "Device "+sDeviceName+" not exist in gallery");
				}
				
			}
			
		}
		/********************************************************************
		 * Function Name: SelectGalleryType
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static string SelectGalleryType(string sType)
		{
			switch (sType)
			{
				case "Detectors":
					sGalleryIndex="4";
					break;
				case "Call points":
					sGalleryIndex="5";
					break;
				case "Sounders/Beacons":
					sGalleryIndex="6";
					break;
				case "Ancillary":
					sGalleryIndex="7";
					break;
				case "Ancillary Conventional":
					sGalleryIndex="8";
					break;
				case "Ancillary Specific":
					sGalleryIndex="9";
					break;
				case "Other":
					sGalleryIndex="10";
					break;
				case "Conventional Sounders":
					sGalleryIndex="11";
					break;
				default:
					Console.WriteLine("Please specify correct gallery name");
					break;
			}
			return sGalleryIndex;
		}
		
		/********************************************************************
		 * Function Name: SelectDeviceFromPanelAccessories
		 * Function Details: to check if the device is enabled/disabled in panel accessories gallery
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static string SelectDeviceFromPanelAccessories(string DeviceName)
		{
			switch (DeviceName)
			{
				case "IOB800":
					sDeviceIndex="0";
					break;
				case "FB800":
					sDeviceIndex="1";;
					break;
				case "PCS800":
					sDeviceIndex="2";
					break;
				case "POS800-S":
					sDeviceIndex="3";
					break;
				case "POS800-M":
					sDeviceIndex="4";
					break;
				case "PBB801":
					sDeviceIndex="5";
					break;
					
				default:
					Console.WriteLine("Please specify correct Device name");
					break;
			}
			return sDeviceIndex;
		}
		
		/********************************************************************
		 * Function Name: DeleteDevices
		 * Function Details: To delete devices using excel
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void DeleteDevices(string sFileName,string sSheetName)
		{
			Excel_Utilities.OpenExcelFile(sFileName,sSheetName);
			//Excel_Utilities.OpenSheet(sSheetName);
			int rows = Excel_Utilities.ExcelRange.Rows.Count;
			
			for(int i=8;i<=rows;i++)
			{
				sLabelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1.Click();
				
				if(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info.Exists())
				{
					repo.ProfileConsys1.btn_Delete.Click();
					Validate.AttributeEqual(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info, "Text", sLabelName);
					Report.Log(ReportLevel.Success, "Device "+sLabelName+" deleted successfully");
				}
				
				else
				{
					
					Report.Log(ReportLevel.Failure, "Device "+sLabelName+" not found");
				}

			}
			
			//Excel_Utilities.ExcelWB.Close(false, null, null);
			//Excel_Utilities.ExcelAppl.Quit();
			
		}
		

		
		/********************************************************************
		 * Function Name: CableCapacitance
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale/Alpesh Dhakad
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void CableCapacitance(string sFileName,string sSheetName)
		{
			Excel_Utilities.OpenExcelFile(sFileName,sSheetName);
			//Excel_Utilities.OpenSheet(sSheetName);
			int rows = Excel_Utilities.ExcelRange.Rows.Count;
			for(int i=8;i<=rows;i++)
			{
				sRow=(i-7).ToString();
				string sDeviceName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				string sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				AddDevicesfromGallery(sDeviceName,sType);
				
				string state =  ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				VerifyGalleryItem(sType,sDeviceName,state);
				
				repo.ProfileConsys1.PanelInvetoryGrid.InventoryGridRow.Click();
				
				string CableCapacitanceValue =  ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				repo.ProfileConsys1.cell_CableCapacitance.Click();
				string capacitance=repo.ProfileConsys1.txt_CableCapacitance.TextValue;
				
				if(capacitance.Equals(CableCapacitanceValue))
				{
					Report.Log(ReportLevel.Success,"default cable capacitance displayed as "+capacitance);
				}
				
				else
				{
					Report.Log(ReportLevel.Failure,"default cable capacitance displayed as "+capacitance+" instead of "+CableCapacitanceValue);
				}
				
				//=================================
				repo.ProfileConsys1.tab_PhysicalLayout.Click();
				
				string maxACUnits = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				verifyMaxACUnitsValue(maxACUnits);
				
				repo.ProfileConsys1.tab_Points.Click();
				
				//===================================
				string ChangedValue =  ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				repo.ProfileConsys1.cell_CableCapacitance.DoubleClick();
				repo.ProfileConsys1.txt_CableCapacitance.PressKeys((ChangedValue) +"{ENTER}");

				
				repo.ProfileConsys1.PanelInvetoryGrid.InventoryGridRow.Click();
				repo.ProfileConsys1.tab_Points.Click();
				
				state =  ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				VerifyGalleryItem(sType,sDeviceName,state);
				
				repo.ProfileConsys1.PanelInvetoryGrid.InventoryGridRow.Click();
				
				//=================================
				repo.ProfileConsys1.tab_PhysicalLayout.Click();
				
				string changedmaxACUnits = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				verifyMaxACUnitsValue(changedmaxACUnits);
				
				repo.ProfileConsys1.tab_Points.Click();
				
				//===================================
			}
			
			Excel_Utilities.ExcelWB.Close(false, null, null);
			Excel_Utilities.ExcelAppl.Quit();
			
			
		}

		/********************************************************************
		 * Function Name: VerifyGalleryItem
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyGalleryItem(string sType,string deviceName, string state)
		{
			if(state.Equals("Enabled"))
			{
				sGalleryIndex = SelectGalleryType(sType);
				ModelNumber=deviceName;
				repo.ProfileConsys1.btn_DevicesGalleryDropDown.Click();
				if (repo.ContextMenu.txt_SelectDevice.Enabled)
				{
					Report.Log(ReportLevel.Success, "Gallery Item: " + deviceName+ " Enabled in gallery");
				}
				else
				{
					Report.Log(ReportLevel.Failure, "Gallery Item: " + deviceName+ " Disabled in gallery");
				}
			}
			else
			{
				sGalleryIndex = SelectGalleryType(sType);
				ModelNumber=deviceName;
				repo.ProfileConsys1.btn_DevicesGalleryDropDown.Click();
				if (repo.ContextMenu.txt_SelectDevice.Enabled)
				{
					Report.Log(ReportLevel.Failure, "Gallery Item: " + deviceName+ " enabled in gallery");
				}
				else
				{
					Report.Log(ReportLevel.Success, "Gallery Item: " + deviceName+ " disabled in gallery");
				}
			}
			
		}
		
		
		
		/********************************************************************
		 * Function Name: ChangeCableLength
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void ChangeCableLength(String sLoopType,int fCableLength1,int fCableLength2)
		{
			float fMaxACUnits;
			repo.ProfileConsys1.tab_Points.Click();
			repo.ProfileConsys1.PanelNode.Click();
			repo.ProfileConsys1.NavigationTree.Loop_A.Click();
			repo.ProfileConsys1.cell_CableLength.Click();
			
			if(sLoopType.Equals("PFI"))
			{
				Keyboard.Press("{LControlKey down}{Akey}{LControlKey up}"+fCableLength1 + "{Enter}");
				//repo.ProfileConsys1.tab_Points.Click();
				repo.ProfileConsys1.PanelNode.Click();
				Delay.Duration(1000, false);
				repo.ProfileConsys1.NavigationTree.Loop_B.Click();
				Delay.Duration(1000, false);
				repo.ProfileConsys1.cell_CableLength.Click();
				Keyboard.Press("{LControlKey down}{Akey}{LControlKey up}"+fCableLength2 + "{Enter}");
				Delay.Duration(500, false);
				fMaxACUnits = (450-(fCableLength1+fCableLength2)/10);
			}
			else
			{
				Keyboard.Press("{LControlKey down}{Akey}{LControlKey up}"+fCableLength1 + "{Enter}");
				fMaxACUnits = (450-(fCableLength1)/10);
			}
			
		}
		
		
		/********************************************************************
		 * Function Name: calculatePercentage
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static string calculatePercentage(float min,float max)
		{
			string expectedColorCode;
			float percentage =(min/max)*100;
			if(percentage!=0 && percentage<95)
			{
				expectedColorCode="GREEN";
			}
			else if(percentage>=95 && percentage<100)
			{
				expectedColorCode="YELLOW";
			}
			else if(percentage>=100)
			{
				expectedColorCode="PINK";
			}
			else
			{
				expectedColorCode="WHITE";
			}
			
			return expectedColorCode;
		}
		
		/********************************************************************
		 * Function Name: VerifyPercentage
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyPercentage(string expectedColorCode,string actualColour)
		{
			switch (expectedColorCode)
			{
				case "GREEN":
					if(actualColour.Equals("#FF90EE90"))
					{
						Report.Log(ReportLevel.Success,"Colour displayed is LightGreen and Units are less than 95% ");
					}
					else
					{
						Report.Log(ReportLevel.Failure,"Progressbar colour is not displayed as LightGreen" + " Colour code is: "+actualColour);
					}
					break;
					
				case "YELLOW":
					if(actualColour.Equals("#FFFFFF00"))
					{
						Report.Log(ReportLevel.Success,"Colour displayed is Yellow and Units are greater than and equal to 95% ");
					}
					else
					{
						Report.Log(ReportLevel.Failure,"Progressbar colour is not displayed as yellow" + " Colour code is: "+actualColour);
					}
					break;
				case "PINK":
					if(actualColour.Equals("#FFFFC0CB"))
					{
						Report.Log(ReportLevel.Success,"Colour displayed is Pink and Units are greater than and equal to 100% ");
					}
					else
					{
						Report.Log(ReportLevel.Failure,"Progressbar colour is not displayed as Pink" + " Colour code is: "+actualColour);
					}
					break;
				case "WHITE":
					if(actualColour.Equals("#FFFFFFFF"))
					{
						Report.Log(ReportLevel.Success,"Colour displayed is white and Units are 0");
					}
					else
					{
						Report.Log(ReportLevel.Failure,"Progressbar colour is not displayed as white" + " Colour code is: "+actualColour);
					}
					break;
					
			}
			
		}
		
		
		
		/********************************************************************
		 * Function Name: AssignDeviceBase
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void AssignDeviceBase(string DeviceLabel, string sBaseofDevice, string sBasePropertyRowIndex)
		{
			int iRowIndex;
			string sExistingBase;
			sBase = sBaseofDevice;
			sRowIndex = sBasePropertyRowIndex;
			sLabelName = DeviceLabel;
			repo.ProfileConsys1.PanelInvetoryGrid.LabelofDevice.Click();
			repo.ProfileConsys1.BaseofDeviceRow.Click();
			repo.ProfileConsys1.BaseofDeviceRow.PressKeys("{Right}");
			int.TryParse(sRowIndex, out iRowIndex);
			iRowIndex = iRowIndex+1;
			sRowIndex = iRowIndex.ToString();
			repo.ProfileConsys1.Cell_BaseofDevice.Click();
			sExistingBase = repo.ProfileConsys1.SomeText.TextValue;
			//sExistingBase = sExistingBase.Replace(@"\""",string.Empty);
			if(!sExistingBase.Equals(sBase))
			{
				repo.ProfileConsys1.BaseofDeviceRow.MoveTo("760;19");
				repo.ProfileConsys1.BaseofDeviceRow.Click("760;19");
				int.TryParse(sRowIndex, out iRowIndex);
				iRowIndex = iRowIndex-1;
				sRowIndex = iRowIndex.ToString();
				repo.ProfileConsys1.BaseofDeviceRow.MoveTo("760;19");
				repo.ProfileConsys1.BaseofDeviceRow.Click("760;19");
				repo.ContextMenu.btn_BaseSelection.Click();
			}
		}

		/********************************************************************
		 * Function Name: AssignAdditionalBase
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void AssignAdditionalBase(string DeviceLabel, string sBaseofDevice, string sBasePropertyRowIndex)
		{
			sLabelName = DeviceLabel;
			repo.ProfileConsys1.tab_Points.Click();
			repo.ProfileConsys1.PanelInvetoryGrid.LabelofDevice.Click();
			sBase = sBaseofDevice;
			sRowIndex = sBasePropertyRowIndex;
			repo.ProfileConsys1.BaseofDeviceRow.MoveTo("760;19");
			repo.ProfileConsys1.BaseofDeviceRow.Click("760;19");
			repo.ContextMenu.btn_BaseSelection.Click();
		}
		
		/********************************************************************
		 * Function Name: RemoveBase
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void RemoveBase(string DeviceLabel, string sBasePropertyRowIndex)
		{
			int iRowIndex;
			sLabelName = DeviceLabel;
			repo.ProfileConsys1.tab_Points.Click();
			repo.ProfileConsys1.PanelInvetoryGrid.LabelofDevice.Click();
			sRowIndex = sBasePropertyRowIndex;
			repo.ProfileConsys1.BaseofDeviceRow.Click();
			repo.ProfileConsys1.BaseofDeviceRow.PressKeys("{Right}");
			int.TryParse(sRowIndex, out iRowIndex);
			iRowIndex = iRowIndex+1;
			sRowIndex = iRowIndex.ToString();
			repo.ProfileConsys1.BaseofDeviceRow.MoveTo("760;19");
			repo.ProfileConsys1.BaseofDeviceRow.Click("760;19");
		}

		/********************************************************************
		 * Function Name: changeAndVerifyNumberOfAlarmLED
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void changeAndVerifyNumberOfAlarmLED(int LEDNumber, string rangeState, int expectedResult)
		{
			
			int Value,actualValue,revertTo;
			string sActualValue;
			repo.ProfileConsys1.SiteNode.Click();
			repo.ProfileConsys1.PanelNode.Click();
			Delay.Duration(500);
			// repo.ProfileConsys1.NavigationTree.Node1Pro32xD.Click();
			
			repo.ProfileConsys1.cell_NumberOfAlarmLeds.Click();
			
			
			if((rangeState.Equals("Valid")))
			{
				Keyboard.Press("{LControlKey down}{Akey}{LControlKey up}"+LEDNumber +"{ENTER}");
				
				//Ranorex.Keyboard.Press(System.Windows.Forms.Keys.Return);
				
				repo.ProfileConsys1.cell_NumberOfAlarmLeds.Click();
				sActualValue = repo.ProfileConsys1.txt_NumberOfAlarmLeds.TextValue;
				int.TryParse(sActualValue, out Value);
				if(Value==LEDNumber)
				{
					Report.Log(ReportLevel.Success,"Number of Alarm LEDs is changed to: "+LEDNumber);
				}
				else
				{
					Report.Log(ReportLevel.Failure,"Number of Alarm LEDs is not changed to: "+LEDNumber);
				}
			}
			else if((rangeState.Equals("InvalidRange")))
			{
				string initialValue = repo.ProfileConsys1.txt_NumberOfAlarmLeds.TextValue;
				int.TryParse(initialValue,out revertTo);
				Keyboard.Press("{LControlKey down}{Akey}{LControlKey up}"+LEDNumber +"{ENTER}");
				repo.ProfileConsys1.cell_NumberOfAlarmLeds.Click();
				string revertedValue = repo.ProfileConsys1.txt_NumberOfAlarmLeds.TextValue;
				int.TryParse(revertedValue, out actualValue);
				if(actualValue==revertTo)
				{
					Report.Log(ReportLevel.Success,"Number of Alarm LEDs is reverted to: "+revertTo);
				}
				else
				{
					Report.Log(ReportLevel.Failure,"Number of Alarm LEDs is not reverted to: "+revertTo);
				}
			}
			else
			{
				
				Keyboard.Press("{LControlKey down}{Akey}{LControlKey up}"+LEDNumber +"{ENTER}");
				repo.ProfileConsys1.cell_NumberOfAlarmLeds.Click();
				string revertedValue = repo.ProfileConsys1.txt_NumberOfAlarmLeds.TextValue;
				int.TryParse(revertedValue,out actualValue);
				if(actualValue==expectedResult)
				{
					Report.Log(ReportLevel.Success,"Number of Alarm LEDs is reverted to: "+expectedResult);
				}
				else
				{
					Report.Log(ReportLevel.Failure,"Number of Alarm LEDs is not reverted to: "+expectedResult);
				}
				
			}
			
		}
		
		/********************************************************************
		 * Function Name: verifyMinMaxThroughSpinControl
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void verifyMinMaxThroughSpinControl(string minLimit,string maxLimit)
		{
			Keyboard.Press("{LControlKey down}{Akey}{LControlKey up}"+ maxLimit +"{ENTER}");
			repo.ProfileConsys1.cell_NumberOfAlarmLeds.Click();
			repo.ProfileConsys1.btn_NumberOfAlarmLedsSpinUpButton.Click();
			string actualValue = repo.ProfileConsys1.txt_NumberOfAlarmLeds.TextValue;
			if(actualValue.Equals(maxLimit))
			{
				Report.Log(ReportLevel.Success,"Spin control accepts values within specified max limit");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Spin control doesnot accepts values within specified max limit");
			}
			Keyboard.Press("{ENTER}");
			
			Keyboard.Press("{LControlKey down}{Akey}{LControlKey up}"+ minLimit +"{ENTER}");
			repo.ProfileConsys1.cell_NumberOfAlarmLeds.Click();
			repo.ProfileConsys1.btn_NumberOfAlarmLedsSpinDownButton.Click();
			actualValue = repo.ProfileConsys1.txt_NumberOfAlarmLeds.TextValue;
			if(actualValue.Equals(minLimit))
			{
				Report.Log(ReportLevel.Success,"Spin control accepts values within specified min limit");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Spin control does not accepts values within specified min limit");
			}
		}
		
		
		/********************************************************************
		 * Function Name: getProgressBarColor
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static string getProgressBarColor(string LoadingType)
		{
			string actualColour;
			switch (LoadingType)
			{
				case "Signal (AC Units)":
					sRowIndex = "1";
					break;
				case "Current (DC Units)":
					sRowIndex = "2";
					break;
				case "Current (worst case)":
					sRowIndex = "3";
					break;
				default:
					Console.WriteLine("Specified loading type doesn't exist");
					break;
			}
			
			return actualColour = repo.ProfileConsys1.DCUnitProgressBar.GetAttributeValue<string>("foreground");
			
		}
		
		
		/********************************************************************
		 * Function Name: DeleteAllDevices
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void DeleteAllDevices()
		{
			Keyboard.Press("{LControlKey down}{Akey}{LControlKey up}");
			repo.ProfileConsys1.btn_Delete.Click();
		}
		
		
		/********************************************************************
		 * Function Name: VerifyDCCalculationforPFI
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public void VerifyDCCalculationforPFI(string sFileName, string sAddDevicesLoopA, string sAddDevicesLoopB,string sPanelLED, string sDeleteDevicesLoopA, string sDeleteDevicesLoopB)
		{
			//Add devies in loop A
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesLoopA);
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			string expectedDCUnits, sType, LabelName;
			for(int i=7; i<=rows; i++)
			{
				ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				sBase = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				sRowIndex= ((Range)Excel_Utilities.ExcelRange.Cells[i,10]).Value.ToString();
				//sDeviceVolume = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				//sFlashRate = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				
				AddDevicesfromGallery(ModelNumber,sType);
				if(sBase!=null && sBase !="NA")
				{
					AssignDeviceBase(sLabelName,sBase,sRowIndex);
				}
				
				repo.ProfileConsys1.NavigationTree.Loop_A.Click();
				Delay.Milliseconds(500);
			}
			
			//Verify DC Units of Loop A
			Report.Log(ReportLevel.Info,"Verification of DC Units of Loop A on addition of devices in Loop A");
			expectedDCUnits= ((Range)Excel_Utilities.ExcelRange.Cells[2,2]).Value.ToString();
			verifyDCUnitsValue(expectedDCUnits);
			
			//Verify DC Units of Loop B
			Report.Log(ReportLevel.Info,"Verification of DC Units of Loop B on addition of devices in Loop A");
			repo.ProfileConsys1.NavigationTree.Loop_B.Click();
			expectedDCUnits= ((Range)Excel_Utilities.ExcelRange.Cells[3,2]).Value.ToString();
			verifyDCUnitsValue(expectedDCUnits);
			
			Excel_Utilities.CloseExcel();
			
			//Add devices in loop B
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesLoopB);
			rows= Excel_Utilities.ExcelRange.Rows.Count;
			repo.ProfileConsys1.NavigationTree.Loop_B.Click();
			for(int i=7; i<=rows; i++)
			{
				ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				sBase = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				sRowIndex= ((Range)Excel_Utilities.ExcelRange.Cells[i,10]).Value.ToString();
				//sDeviceVolume = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				//sFlashRate = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				
				AddDevicesfromGallery(ModelNumber,sType);
				if(sBase!=null && sBase !="NA")
				{
					AssignDeviceBase(sLabelName,sBase,sRowIndex);
				}
				
				repo.ProfileConsys1.NavigationTree.Loop_B.Click();
				Delay.Milliseconds(500);
			}
			
			//Verify DC Units of Loop B
			Report.Log(ReportLevel.Info,"Verification of DC Units of Loop A on addition of devices in Loop B");
			expectedDCUnits= ((Range)Excel_Utilities.ExcelRange.Cells[2,2]).Value.ToString();
			verifyDCUnitsValue(expectedDCUnits);
			
			//Verify DC Units of Loop A
			repo.ProfileConsys1.NavigationTree.Loop_A.Click();
			Report.Log(ReportLevel.Info,"Verification of DC Units of Loop B on addition of devices in Loop B");
			expectedDCUnits= ((Range)Excel_Utilities.ExcelRange.Cells[3,2]).Value.ToString();
			verifyDCUnitsValue(expectedDCUnits);
			
			Excel_Utilities.CloseExcel();
			Report.Log(ReportLevel.Info,"Verification of DC Units of on changing Panel LED");
			verifyPanelLEDEffectOnDC(sFileName,sPanelLED);
			
			Excel_Utilities.CloseExcel();
			
			//Delete Devices from loop A
			repo.ProfileConsys1.NavigationTree.Loop_A.Click();
			repo.ProfileConsys1.tab_Points.Click();
			
			Excel_Utilities.OpenExcelFile(sFileName,sDeleteDevicesLoopA);
			rows = Excel_Utilities.ExcelRange.Rows.Count;
			
			for(int i=3;i<=rows;i++)
			{
				LabelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				DeleteDeviceUsingLabel(LabelName);
			}
			
			//Verify DC Units of Loop A
			Report.Log(ReportLevel.Info,"Verification of DC Units of Loop A on deletion of devices from Loop A");
			expectedDCUnits= ((Range)Excel_Utilities.ExcelRange.Cells[1,2]).Value.ToString();
			verifyDCUnitsValue(expectedDCUnits);
			
			//Verify DC Units of Loop B
			Report.Log(ReportLevel.Info,"Verification of DC Units of Loop B on deletion of devices from Loop A");
			repo.ProfileConsys1.NavigationTree.Loop_B.Click();
			verifyDCUnitsValue(expectedDCUnits);
			
			Excel_Utilities.CloseExcel();
			
			//Delete Devices from loop B
			repo.ProfileConsys1.NavigationTree.Loop_B.Click();
			repo.ProfileConsys1.tab_Points.Click();
			
			Excel_Utilities.OpenExcelFile(sFileName,sDeleteDevicesLoopB);
			rows = Excel_Utilities.ExcelRange.Rows.Count;
			
			for(int i=3;i<=rows;i++)
			{
				LabelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				DeleteDeviceUsingLabel(LabelName);
			}
			
			//Verify DC Units of Loop B
			Report.Log(ReportLevel.Info,"Verification of DC Units of Loop A on deletion of devices from Loop B");
			expectedDCUnits= ((Range)Excel_Utilities.ExcelRange.Cells[1,2]).Value.ToString();
			verifyDCUnitsValue(expectedDCUnits);
			
			//Verify DC Units of Loop A
			Report.Log(ReportLevel.Info,"Verification of DC Units of Loop B on deletion of devices from Loop B");
			repo.ProfileConsys1.NavigationTree.Loop_A.Click();
			verifyDCUnitsValue(expectedDCUnits);
		}

		/********************************************************************
		 * Function Name: DeleteDeviceUsingLabel
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void DeleteDeviceUsingLabel(string LabelName)
		{
			sLabelName = LabelName;
			
			repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1.Click();
			
			if(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info.Exists())
			{
				repo.ProfileConsys1.btn_Delete.Click();
				Validate.AttributeEqual(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info, "Text", sLabelName);
				Report.Log(ReportLevel.Success, "Device "+sLabelName+" deleted successfully");
			}
			
			else
			{
				
				Report.Log(ReportLevel.Failure, "Device "+sLabelName+" not found");
			}
		}
		
		/********************************************************************
		 * Function Name: VerifyDeviceSensitivity
		 * Function Details: To verify device sensitivity value
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :
		 ********************************************************************/
		// To verify device sensitivity value as per the argument
		[UserCodeMethod]
		public static void VerifyDeviceSensitivity(string sDeviceSensitivity)
		{
			// Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Enter the Device text in Search Properties fields to view device related text
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("Device" +"{ENTER}");
			
			// Click on Device Sensitivity cell
			repo.ProfileConsys1.PARTItemsPresenter.cell_DeviceSensitivity.Click();
			
			// Get the text value of Device Sensitivity field
			string DeviceSensitivity = repo.ProfileConsys1.DeviceSensitivity.TextValue;
			
			//Comparing expected and actual Device Sensitivity value
			if(DeviceSensitivity.Equals(sDeviceSensitivity))
			{
				Report.Log(ReportLevel.Success,"Device Sensitivity " +DeviceSensitivity + " is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Device Sensitivity is not displayed correctly");
			}
			
			//Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
			
			
		}
		
		/********************************************************************
		 * Function Name: ChangeDeviceSensitivity
		 * Function Details: To change device sensitivity
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :
		 ********************************************************************/
		// To change device sensitivity value as per the argument
		[UserCodeMethod]
		public static void ChangeDeviceSensitivity(string changeDeviceSensitivity)
		{
			//Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Enter the Device text in Search Properties fields to view device related text
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("Device" +"{ENTER}" );
			
			// Click on Device Sensitivity cell
			repo.ProfileConsys1.PARTItemsPresenter.cell_DeviceSensitivity.Click();
			
			// Enter the value to change device sensitivity
			repo.ProfileConsys1.PARTItemsPresenter.txt_changeDeviceSensitivity.PressKeys((changeDeviceSensitivity) +"{ENTER}" + "{ENTER}");
			
			// Click on Device Sensitivity cell
			repo.ProfileConsys1.PARTItemsPresenter.cell_DeviceSensitivity.Click();
			
			// Get the text value of changed Device Sensitivity field
			string DeviceSensitivity = repo.ProfileConsys1.PARTItemsPresenter.txt_changeDeviceSensitivity.TextValue;
			
			//Comparing expected and actual changed Device Sensitivity value
			if(DeviceSensitivity.Equals(changeDeviceSensitivity))
			{
				Report.Log(ReportLevel.Success,"Device Sensitivity changed successfully to " +DeviceSensitivity + " and is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Device Sensitivity is not changed to " +changeDeviceSensitivity + "and displayed incorrectly");
			}
			
			
			//Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
			
		}
		
		/********************************************************************
		 * Function Name: VerifyDeviceMode
		 * Function Details: To verify device mode
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :
		 ********************************************************************/
		// To verify device mode value as per the argument
		[UserCodeMethod]
		public static void VerifyDeviceMode(string sDeviceMode)
		{
			//Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Enter the Device text in Search Properties fields to view device related text
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("Device" +"{ENTER}" );
			
			// Click on Device Mode cell
			repo.ProfileConsys1.PARTItemsPresenter.cell_DeviceMode.Click();
			
			// Get the text value of changed Device Mode field
			string DeviceMode = repo.ProfileConsys1.DeviceMode.TextValue;
			
			//Comparing expected and actual changed Device Mode value
			if(DeviceMode.Equals(sDeviceMode))
			{
				Report.Log(ReportLevel.Success,"Device mode " +DeviceMode+ " is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Device mode is not displayed correctly");
			}
			
			//Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
			
		}
		
		/********************************************************************
		 * Function Name: ChangeDeviceMode
		 * Function Details: To change device mode
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :
		 ********************************************************************/
		// To change and verify device mode value as per the argument
		[UserCodeMethod]
		public static void ChangeDeviceMode(string changeDeviceMode)
		{
			//Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Enter the Device text in Search Properties fields to view device related text
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("Device" +"{ENTER}" );
			
			// Click on Device Mode cell
			repo.ProfileConsys1.PARTItemsPresenter.cell_DeviceMode.Click();
			
			// Enter the value to change device mode
			repo.ProfileConsys1.PARTItemsPresenter.txt_changeDeviceMode.PressKeys((changeDeviceMode) +"{ENTER}" + "{ENTER}");
			
			// Click on Device Mode cell
			repo.ProfileConsys1.PARTItemsPresenter.cell_DeviceMode.Click();
			
			// Get the text value of changed Device Mode field
			string DeviceMode = repo.ProfileConsys1.PARTItemsPresenter.txt_changeDeviceMode.TextValue;
			
			//Comparing expected and actual changed Device Mode value
			if(DeviceMode.Equals(changeDeviceMode))
			{
				Report.Log(ReportLevel.Success,"Device mode changed successfully to " +DeviceMode+ " and is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Device mode is not changed to " +changeDeviceMode+ " and displayed incorrectly");
			}
			
			//Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
			
		}
		
		
		/********************************************************************
		 * Function Name: CheckUncheckDayMatchesNight
		 * Function Details: To check and uncheck day matches night checkbox
		 * Parameter/Arguments: boolean value
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :
		 ********************************************************************/
		// To verify check box state of Day matches night field
		[UserCodeMethod]
		public static void CheckUncheckDayMatchesNight(bool ExpectedState)
		{
			//Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Enter the Day Matches night text in Search Properties fields to view day matches night related text;
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("Day Matches Night" +"{ENTER}" );

			// CLick on checkbox cell lower left corner
			repo.ProfileConsys1.PARTItemsPresenter.cell_DayMatchesNight.Click(Location.LowerLeft);
			// repo.ProfileConsys1.PARTItemsPresenter.chkbox_DayMatchesNight.EnsureVisible();
			
			// To retrieve the attribute value as boolean by its ischecked properties and store in actual state
			bool actualState =  repo.ProfileConsys1.PARTItemsPresenter.chkbox_DayMatchesNight.GetAttributeValue<bool>("ischecked");
			
			//As per actual state and expected state values verfiying day mode and day sensitivity field state and action performed on checkbox
			if(actualState)
			{
				if(ExpectedState)
				{
					// Verify Day mode field state
					VerifyDayModeField(ExpectedState);
					
					// Verify Day Sensitivity field state
					VerifyDaySensitivityField(ExpectedState);
				}
				else
				{
					// Click on Day Matches night checkbox
					repo.ProfileConsys1.PARTItemsPresenter.chkbox_DayMatchesNight.Click();
					
					// Verify Day mode field state
					VerifyDayModeField(ExpectedState);
					
					// Verify Day Sensitivity field state
					VerifyDaySensitivityField(ExpectedState);
				}
			}
			
			else
			{
				if(ExpectedState)
				{
					// Click on Day Matches night checkbox
					repo.ProfileConsys1.PARTItemsPresenter.chkbox_DayMatchesNight.Click();
					
					// Verify Day mode field state as disabled state
					VerifyDayModeField(ExpectedState);
					
					// Verify Day Sensitivity field state
					VerifyDaySensitivityField(ExpectedState);
				}
				else
				{
					// Verify Day mode field state as enabled state
					VerifyDayModeField(ExpectedState);
					
					// Verify Day Sensitivity field state
					VerifyDaySensitivityField(ExpectedState);
				}
			}
			//Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
		}
		
		
		/********************************************************************
		 * Function Name: VerifyDayModeField
		 * Function Details: To verify day mode field
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :
		 ********************************************************************/
		// To verify day mode value field as per the argument
		[UserCodeMethod]
		public static void VerifyDayModeField(bool ExpectedDayModeState)
		{
			// To retrieve the attribute value as boolean by its "isreadonly" properties and store in verifyReadOnly
			bool verifyReadOnly = repo.ProfileConsys1.PARTItemsPresenter.row_DayModeField.GetAttributeValue<bool>("isreadonly");
			
			// Comparing verifyReadOnly and ExpectedDayModeState values
			if(verifyReadOnly.Equals(ExpectedDayModeState))
			{
				Report.Log(ReportLevel.Success,"Day mode field is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Day mode field is displayed incorrectly");
			}
		}
		
		/********************************************************************
		 * Function Name: VerifyDaySensitivityField
		 * Function Details: To verify day sensitivity field
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :
		 ********************************************************************/
		// To verify day sensitivity value field as per the argument
		[UserCodeMethod]
		public static void VerifyDaySensitivityField(bool ExpectedDaySenstivityState)
		{
			// To retrieve the attribute value as boolean by its "isreadonly" properties and store in verifyReadOnly
			bool verifyReadOnly = repo.ProfileConsys1.PARTItemsPresenter.row_DaySensitivityField.GetAttributeValue<bool>("isreadonly");
			
			// Comparing verifyReadOnly and ExpectedDayModeState values
			if(verifyReadOnly.Equals(ExpectedDaySenstivityState))
			{
				Report.Log(ReportLevel.Success,"Day Sensitivity field is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Day Sensitivity field is displayed incorrectly");
			}
		}
		
		/********************************************************************
		 * Function Name: VerifyDayMode
		 * Function Details: To verify day mode
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :
		 ********************************************************************/
		// To verify day mode value field as per the argument
		[UserCodeMethod]
		public static void VerifyDayMode(string sDayMode)
		{
			//Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Enter the Day Matches night text in Search Properties fields to view day matches night related text;
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("Day Matches Night" +"{ENTER}" );
			
			// Click on Day mode cell
			repo.ProfileConsys1.PARTItemsPresenter.cell_DayMode.Click();
			
			// Retrieve value of Day mode and store in DayMode
			string DayMode = repo.ProfileConsys1.PARTItemsPresenter.txt_DayMode.Text;
			
			// Comparing DayMode and sDayMode values
			if(DayMode.Equals(sDayMode))
			{
				Report.Log(ReportLevel.Success,"Day mode " +DayMode+ " is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Day mode is not displayed correctly");
			}
			//Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
		}

		
		
		/********************************************************************
		 * Function Name: ChangeDayMode
		 * Function Details: To change day mode
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		//  Purpose : To change and verify day mode value field as per the argument
		public static void ChangeDayMode(string changeDayMode)
		{
			//Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Enter the Day Matches night text in Search Properties fields to view day matches night related text;
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("Day Matches Night" +"{ENTER}" );
			
			// Click on Day mode cell
			repo.ProfileConsys1.PARTItemsPresenter.cell_DayMode.Click();
			
			// Enter the changeDayMode value and click Enter twice
			repo.ProfileConsys1.PARTItemsPresenter.txt_DayMode.PressKeys((changeDayMode) +"{ENTER}" + "{ENTER}");
			
			// Click on Day mode cell
			repo.ProfileConsys1.PARTItemsPresenter.cell_DayMode.Click();
			
			//Retrieve value of changed Day mode text and store in DayMode
			string DayMode = repo.ProfileConsys1.PARTItemsPresenter.txt_changeDayMode.TextValue;
			
			// Comparing DayMode and changeDayMode values
			if(DayMode.Equals(changeDayMode))
			{
				Report.Log(ReportLevel.Success,"Day mode changed successfully to " +DayMode+ " and is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Day mode is not changed to " +changeDayMode+ " and displayed incorrectly");
			}
			//Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
		}
		
		/********************************************************************
		 * Function Name: VerifyDaySensitivity
		 * Function Details: To Verify day sensitivity
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :
		 ********************************************************************/
		// To verify day sensitivity value field as per the argument
		[UserCodeMethod]
		public static void VerifyDaySensitivity(string sDaySensitivity)
		{
			//Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Enter the Day Matches night text in Search Properties fields to view day matches night related text;
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("Day Matches Night" +"{ENTER}" );
			
			// Click on Day Sensitivity cell
			repo.ProfileConsys1.PARTItemsPresenter.cell_DaySensitivity.Click();
			
			//Retrieve value of Day Sensitivity text and store in DaySensitivity
			string DaySensitivity = repo.ProfileConsys1.PARTItemsPresenter.txt_DaySensitivity.Text;
			
			// Comparing DaySensitivity and sDaySensitivity values
			if(DaySensitivity.Equals(sDaySensitivity))
			{
				Report.Log(ReportLevel.Success,"Day Sensitivity " +DaySensitivity+ " is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Day Sensitivity is not displayed correctly");
			}
			//Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
		}
		
		/********************************************************************
		 * Function Name: ChangeDaySensitivity
		 * Function Details: To change day sensitivity
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :
		 ********************************************************************/
		// To change and verify day sensitivity value field as per the argument
		[UserCodeMethod]
		public static void ChangeDaySensitivity(string changeDaySensitivity)
		{
			//Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Enter the Day Matches night text in Search Properties fields to view day matches night related text;
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("Day Matches Night" +"{ENTER}" );
			
			// Click on Day Sensitivity cell
			repo.ProfileConsys1.PARTItemsPresenter.cell_DaySensitivity.Click();
			
			// Enter the changeDaySensitivity value and click Enter twice
			repo.ProfileConsys1.PARTItemsPresenter.txt_DaySensitivity.PressKeys((changeDaySensitivity) +"{ENTER}" + "{ENTER}");
			
			// Click on Day Sensitivity cell
			repo.ProfileConsys1.PARTItemsPresenter.cell_DaySensitivity.Click();
			
			//Retrieve value of changed Day Sensitivity text and store in DaySensitivity
			string DaySensitivity = repo.ProfileConsys1.PARTItemsPresenter.txt_changeDaySensitivity.TextValue;
			
			// Comparing DaySensitivity and changeDaySensitivity values
			if(DaySensitivity.Equals(changeDaySensitivity))
			{
				Report.Log(ReportLevel.Success,"Day Sensitivity changed successfully to " +DaySensitivity+ " and is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Day Sensitivity is not changed to " +changeDaySensitivity+ " and displayed incorrectly");
			}
			//Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
		}
		
		/*************************************************************************************************************
		 *                Below functions are mapped to respective code libraries as per functionality
		 * ***********************************************************************************************************
		 *                                     NOTE : DO NOT REMOVE IT FROM HERE
		 * ************************************************************************************************************/
		
		public static void VerifyACCalculation(string sFileName,string sAddDevicesSheet, string sDeleteDevicesSheet)
		{
			AC_Functions.VerifyACCalculation(sFileName,sAddDevicesSheet,sDeleteDevicesSheet);
		}
		
		
		public static void VerifyACCalculationforFIM(string sFileName,string sAddDevicesSheet, string sDeleteDevicesSheet)
		{
			AC_Functions.VerifyACCalculationforFIM(sFileName,sAddDevicesSheet,sDeleteDevicesSheet);
			
		}
		
		public static void verifyMaxACUnitsValue(string expectedMaxACUnits)
		{
			AC_Functions.verifyMaxACUnitsValue(expectedMaxACUnits);
			
		}
		
		
		public static void verifyACUnitsValue(string expectedACUnits)
		{
			AC_Functions.verifyACUnitsValue(expectedACUnits);
		}
		
		public static void verifyMaxDCUnits(string expectedMaxDCUnits)
		{
			DC_Functions.verifyMaxDCUnits(expectedMaxDCUnits);
			
		}
		
		
		public static void verifyDCUnitsValue(string expectedDCUnits)
		{
			DC_Functions.verifyDCUnitsValue(expectedDCUnits);
		}
		
		
		public static void VerifyDCUnitsIndicators(string sFileName,string sAddDevicesSheet)
		{
			DC_Functions.VerifyDCUnitsIndicators(sFileName,sAddDevicesSheet);
		}
		
		
		public static void verifyPanelLEDEffectOnDC(string sFileName,string sPanelLED)
		{
			DC_Functions.verifyPanelLEDEffectOnDC(sFileName,sPanelLED);
		}
		
		
		public static void changeDeviceSensitivityAndVerifyDCUnit(string sFileName,string sAddDevicesSheet)
		{
			DC_Functions.changeDeviceSensitivityAndVerifyDCUnit(sFileName,sAddDevicesSheet);
		}
		
		
		public static void VerifyDCUnitsAndWorstCaseIndicators(string sFileName,string sAddDevicesSheet)
		{
			DC_Functions.VerifyDCUnitsAndWorstCaseIndicators(sFileName,sAddDevicesSheet);
		}
		
		
		// To verify voltage drop value on adding and removing devices
		public static void verifyVoltageDropOnAddingAndRemovingDevices(string sFileName,string sAddDevicesLoopA, string sDeleteDevicesLoopA)
		{
			VoltageDrop_Functions.verifyVoltageDropOnAddingAndRemovingDevices(sFileName,sAddDevicesLoopA,sDeleteDevicesLoopA);
		}
		
		// Verify max volt Drop value
		public static void verifyMaxVoltDrop(string expectedVoltDropMaxValue)
		{
			VoltageDrop_Functions.verifyMaxVoltDrop(expectedVoltDropMaxValue);
		}
		
		// Verify volt Drop value
		public static void verifyVoltDropValue(string expectedVoltDropValue)
		{
			VoltageDrop_Functions.verifyVoltDropValue(expectedVoltDropValue);
		}
		
		// Verify max volt Drop worst case value
		public static void verifyMaxVoltDropWorstCaseValue(string expectedVoltDropWorstCaseMaxValue)
		{
			VoltageDrop_Functions.verifyMaxVoltDropWorstCaseValue(expectedVoltDropWorstCaseMaxValue);
		}
		
		// Verify volt Drop worst case value
		public static void verifyVoltDropWorstCaseValue(string expectedVoltDropWorstCaseValue)
		{
			VoltageDrop_Functions.verifyVoltDropWorstCaseValue(expectedVoltDropWorstCaseValue);
		}
		
		// Verify Voltage Drop Calculation on Adding devices in loops
		public static void verifyVoltageDropCalculation(string sFileName,string sAddDevicesLoop)
		{
			VoltageDrop_Functions.verifyVoltageDropCalculation(sFileName,sAddDevicesLoop);
		}
		
		// Verify Voltage Drop percentage
		public static void verifyVoltageDropPercentage(string sFileName, string noLoadVoltDrop)
		{
			VoltageDrop_Functions.verifyVoltageDropPercentage(sFileName,noLoadVoltDrop);
		}
		
		

		/*************************************************************************************************************
		 *                Above functions are mapped to respective code libraries as per functionality
		 * ***********************************************************************************************************
		 *                                     NOTE : DO NOT REMOVE IT FROM HERE
		 * ************************************************************************************************************/
		

		
		/********************************************************************
		 * Function Name: AddMultipleDevices
		 * Function Details: To add multiple devices
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Poonam Kadam
		 * Last Update : 29/03/2019 Alpesh Dhakad- Updated btn_MultiplePointWizard xpath and change script accordingly
		 ********************************************************************/
		[UserCodeMethod]
		public static void AddMultipleDevices(string sFileName, string sSheetName)
		{
			
			Excel_Utilities.OpenExcelFile(sFileName,sSheetName);
			//Excel_Utilities.OpenSheet(sSheetName);
			int rows = Excel_Utilities.ExcelRange.Rows.Count;
			repo.FormMe.btn_MultiplePointWizard.Click();
			
			//repo.ProfileConsys1.btn_MultiplePointWizard_DoNotUse.Click();
			repo.AddDevices.txt_AllDevices.Click();
			
			string sDeviceName=  ((Range)Excel_Utilities.ExcelRange.Cells[8,1]).Value.ToString();
			int DeviceQty=  int.Parse(((Range)Excel_Utilities.ExcelRange.Cells[8,2]).Value.ToString());
			
			repo.AddDevices.txt_SearchDevices.Click();
			Keyboard.Press("{LControlKey down}{Akey}{LControlKey up}"+sDeviceName);
			ModelNumber = sDeviceName;
			repo.AddDevices.txt_ModelNumber.Click();
			repo.AddDevices.txt_Quantity.Click();
			Keyboard.Press("{LControlKey down}{Akey}{LControlKey up}"+DeviceQty.ToString());
			
			repo.AddDevices.btn_AddDevices.Click();
			Report.Log(ReportLevel.Success,+DeviceQty+" \""+sDeviceName+ "\" Device Added successfully");
			Delay.Milliseconds(200);/*
			for(int i=1;i<=DeviceQty;i++)
			{
				string DeviceName = ((Range)Excel_Utilities.ExcelRange.Cells[2,1]).Value.ToString();
				ModelNumber = DeviceName;
				sRow = (i).ToString();
				repo.ProfileConsys1.PanelInvetoryGrid.InventoryGridCell.Click();
				Validate.AttributeEqual(repo.ProfileConsys1.PanelInvetoryGrid.txt_DeviceNameInfo, "Text", DeviceName);
				Delay.Milliseconds(100);
				Report.Log(ReportLevel.Success,"Number "+i+" "+sDeviceName+ "Device verified in points grid successfully");
			}*/
			
			Excel_Utilities.ExcelWB.Close(false, null, null);
			//Excel_Utilities.ExcelAppl.Quit();
		}



		
//			[UserCodeMethod]
//			public static void GetList()
//			{
//
//				//Click on Points tab
//			repo.ProfileConsys1.tab_Points.Click();
//
//			// Click on SearchProperties text field
//			repo.ProfileConsys1.txt_SearchProperties.Click();
//
//			// Enter the Device text in Search Properties fields to view device related text
//			repo.ProfileConsys1.txt_SearchProperties.PressKeys("Device" +"{ENTER}" );
//
//			repo.ProfileConsys1.PARTItemsPresenter.cell_DeviceSensitivity.Click();
//			repo.ProfileConsys1.PARTItemsPresenter.PARTItem.Click();
//			repo.ProfileConsys1.PARTItemsPresenter.PARTItem.Click();
//
//			IList<Ranorex.ListItem> dropDownListText = repo.ContextMenu.PARTContent.FindDescendants();
//
//			foreach (Ranorex.ListItem element in dropDownListText)
//			{
//				string s1 = dropDownListText.ElementAt(0);
//				int ab = dropDownListText.IndexOf("Low (60dB)");
//				Ranorex.ListItem s = repo.ContextMenu.PARTContent.Items;
//				Report.Log(ReportLevel.Success,s.ToString());
//			}
//			if (repo.ContextMenu.PARTContent.Items.Count > 0)
//			{
//				Report.Log(ReportLevel.Success,"Success ");
//			}
//			else
//			{
//				Report.Log(ReportLevel.Failure,"Failed");
//			}
//
		
//			7
		
//			//Click on Points tab
//			repo.ProfileConsys1.tab_Points.Click();
//
//			// Click on SearchProperties text field
//			repo.ProfileConsys1.txt_SearchProperties.Click();
//
//			// Enter the Device text in Search Properties fields to view device related text
//			repo.ProfileConsys1.txt_SearchProperties.PressKeys("Device" +"{ENTER}" );
		//====================
//			// Click on Device Sensitivity cell
//			repo.ProfileConsys1.PARTItemsPresenter.cell_DeviceSensitivity.Click();
//			repo.ProfileConsys1.PARTItemsPresenter.PARTItem.Click();
//
//
//			repo.ProfileConsys1.PARTItemsPresenter.PARTItem.Click();
//
//			============================
		////			int count = repo.ProfileConsys1.PARTItemsPresenter.ComboboxList.Element.Children.Count;
		////			for (int i=0; i<=count; i++)
//			{
//				string sensitivity = repo.ProfileConsys1.PARTItemsPresenter.ComboboxList.Element.Children.IndexOf(i).ToString();
//				//string sensitivity = repo.ProfileConsys1.PARTItemsPresenter.ComboboxList.Element.Children.IndexOf(i);
//			}
		
//
//			IList<ListItem> List = repo.ProfileConsys1.PARTItemsPresenter.ComboboxList.FindChildren
//			foreach (ListItem Item in List)
//			{
//				string sensitivity = repo.ProfileConsys1.PARTItemsPresenter.ComboboxList.Element.Children.(Item);
//			}
		//string expectedValue= "Low (60dB)";
		//Select(item => item.Text).
		//List<string> dropDownListText = repo.ContextMenu.PARTContent.Items;
//
//			//repo.ContextMenu.SomeListItem.Text();
//			try{
//				//string txt1 = repo.ContextMenu.PARTContent.Items;
//				Thread.Sleep(500);
//			IList<Ranorex.ListItem> dropDownListText = repo.ContextMenu.PARTContent.Items;
//
//			foreach (Ranorex.ListItem element in dropDownListText)
//			{
//				ListItem s = repo.ContextMenu.PARTContent.Items.ElementAt(0);
//				Report.Log(ReportLevel.Success,s.Text);
//
//			}
//
//			if (repo.ContextMenu.PARTContent.Items.Count > 0)
//			{
//				Report.Log(ReportLevel.Success,"Day Sensitivity changed successfully to ");
//			}
//			else
//			{
//				Report.Log(ReportLevel.Failure,"Day Sensitivity failed to ");
//			}
//			}
//			catch(Exception e){
//
//			}
//
		
//			if(dropDownListText.Where(item => item == expectedValue).Count() >0)
//			{
//				Report.Log(ReportLevel.Success,"Day Sensitivity changed successfully to ");
//			}
//			else
//			{
//				Report.Log(ReportLevel.Failure,"Day Sensitivity failed to ");
//			}
//			================================================
//			int count = repo.ProfileConsys1.PARTItemsPresenter.ComboboxList.Element.Children.Count;
//			IList<ListItem> cmb = repo.ContextMenu.PARTContent.Items;
//
//			//string s = cmb.Text;
//			Report.Log(ReportLevel.Success, s);
//			for (int i=0; i<count; i++)
//			{
//			   foreach (Ranorex.Text txt in cmb.FindDescendants<Ranorex.Text>())
//				{
//						Report.Log(ReportLevel.Success, txt.TextValue);
//				}
//			}
		//}
		//===================================
//
//
		
		
		
//			foreach (ListItem listItem in repo.ContextMenu.PARTContent.Items)
//			{
//				dropDownListText.Add(listItem.Text);
//			}
//
		
//
//			//for (int i=0; i<=count; i++)
//			{
//				Ranorex.ComboBox cmb = repo.ProfileConsys1.PARTItemsPresenter.cell_DeviceSensitivity.FindSingle(".//combobox", 60000);
//				cmb.Click();
//				IList<Ranorex.ListItem> Items = cmb.FindDescendants<Ranorex.ListItem>();
//
		//            	foreach (Ranorex.ListItem ListI in Items)
		//            	{
		//               	Report.Log(ReportLevel.Success, ListI.Text.ToString);
		//            	}
//
//
//			}

		
		
		//repo.ProfileConsys1.PARTItemsPresenter.ComboboxList.
		

		//myCombobox.GetItemText(myCombobox.Items[index])
		

		//repo.ProfileConsys1.PARTItemsPresenter.ComboboxList.Element.Children.IndexOf(ListItem i)
		
//			for (int i=0; i<=count; i++)
//			{
//				Ranorex.ComboBox cmb = repo.ProfileConsys1.PARTItemsPresenter.ComboboxList;
//
//
//
//
//				foreach (Ranorex.Text txt in cmb.FindDescendants<Ranorex.Text>())
//				{
//				Report.Log(ReportLevel.Success, txt.TextValue);
//				}
//
		////			}
		
//			Ranorex.ListItem
//
//			Ranorex.Container cont = "yourPath";
//			foreach(Ranorex.Text txt in cont.FindChildren<Ranorex.Text>())
//			{
		//    		Report.Log(ReportLevel.Success, txt.TextValue);
//			}
//
		
		//ComboBox cmb = repo.ProfileConsys1.PARTItemsPresenter.ComboboxList.Element.Children;
		
		// IList<Ranorex.ListItem> MyListItems = cmb.FindDescendants<Ranorex.ListItem>();

		//            foreach (Ranorex.ComboBox ThisListItem in repo.ProfileConsys1.PARTItemsPresenter.ComboboxList.Element.Children)
		//            {
		//           	Report.Log(ReportLevel.Success, ThisListItem.Text);
//
		//            }
//	}
//		}
//		}
		
		//	}
		
		
		/********************************************************************
		 * Function Name: ChangeCableLength
		 * Function Details: To change cable length
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :
		 ********************************************************************/
		// Change cable length method
		[UserCodeMethod]
		public static void ChangeCableLength(int fchangeCableLength)
		{
			//Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			//Click on Panel Node
			repo.ProfileConsys1.PanelNode.Click();
			
			//Click on Loop A in Navigation tree tab
			repo.ProfileConsys1.NavigationTree.Loop_A.Click();
			
			//Click on cable length cell
			repo.ProfileConsys1.cell_CableLength.Click();
			
			//Change the value of cable length
			Keyboard.Press("{LControlKey down}{Akey}{LControlKey up}"+fchangeCableLength + "{Enter}");
			
			//Click on Panel Node
			repo.ProfileConsys1.PanelNode.Click();
			Delay.Duration(1000, false);
		}
		
		/********************************************************************
		 * Function Name: ChangeCableResistance
		 * Function Details: To change cable resistance
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :
		 ********************************************************************/
		// Change cable resistance method
		[UserCodeMethod]
		public static void ChangeCableResistance(string fchangeCableResistance)
		{
			//Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			//Click on Panel Node
			repo.ProfileConsys1.PanelNode.Click();
			
			//Click on Loop A in Navigation tree tab
			repo.ProfileConsys1.NavigationTree.Loop_A.Click();
			
			//Click on cable resistance cell
			repo.ProfileConsys1.cell_CableResistance.Click();
			
			//Change the value of cable length
			Keyboard.Press("{LControlKey down}{Akey}{LControlKey up}"+fchangeCableResistance + "{Enter}");
			
			//Click on Panel Node
			repo.ProfileConsys1.PanelNode.Click();
			Delay.Duration(1000, false);
		}
		
		
		
		
		
		
//		 Report.Log(ReportLevel.Info, "Mouse", "Mouse Right Click item 'ProfileConsys1.Row11' at 11;10.", repo.ProfileConsys1.Row11Info, new RecordItemIndex(5));
		//            repo.ProfileConsys1.Row11.Click(System.Windows.Forms.MouseButtons.Right, "11;10");
		//            Delay.Milliseconds(200);
//
		//            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Down item 'ProfileConsys1.Row11' at 12;11.", repo.ProfileConsys1.Row11Info, new RecordItemIndex(6));
		//            repo.ProfileConsys1.Row11.MoveTo("12;11");
		//            Mouse.ButtonDown(System.Windows.Forms.MouseButtons.Left);
		//            Delay.Milliseconds(200)
//

		/********************************************************************
		 * Function Name: AddDevicesForBVT
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Devendra Kulkarni
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void AddDevicesForBVT(string sFileName, string singleDevice, string multipleDevices)
		{
			if(!singleDevice.IsEmpty())
			{
				// Open the excel file and sheet with mentioned name in argument
				Excel_Utilities.OpenExcelFile(sFileName, singleDevice);
				
				// Count the number of rows in excel
				int rows= Excel_Utilities.ExcelRange.Rows.Count;
				
				// Declared various fields as String type
				string modelNumber, sType;
				
				// For loop to fetch values from the excel sheet and then add devices
				for(int i=8; i<=rows; i++)
				{
					modelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
					sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
					
					// Add devices from the gallery as per test data from the excel sheet
					AddDevicesfromGallery(modelNumber, sType);
				}
				
				// Close the currently opened excel sheet
				Excel_Utilities.CloseExcel();
			}
			
			if(!multipleDevices.IsEmpty())
			{
				AddMultipleDevices(sFileName, multipleDevices);
			}
		}
		
		/********************************************************************
		 * Function Name: VerifyProperties_BVT
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Devendra Kulkarni
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyProperties_BVT()
		{/*
			repo.FormMe.PointsModel1.Click();
			string modelProperties = repo.FormMe.PropertyModel1.TextValue;
			if (model.Equals(modelProperties))
			{
				Report.Log(ReportLevel.Success, "Properties window displayed for device 1.");
			}
			else
			{
				Report.Log(ReportLevel.Error, "Properties window not displayed for device 1.");
			}
			
			model = repo.FormMe.PointsLabel2.TextValue;
			repo.FormMe.PointsLabel2.Click();
			modelProperties = repo.FormMe.PropertyLabel2.TextValue;
			if (model.Equals(modelProperties))
			{
				Report.Log(ReportLevel.Success, "Properties window displayed for device 2.");
			}
			else
			{
				Report.Log(ReportLevel.Error, "Properties window not displayed for device 2.");
			}*/
		}
		
		
		/********************************************************************
		 * Function Name: AssignDeviceBaseForMultipleDevices
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Poonam Kadam
		 * Last Update : 22/03/2019 - Alpesh Dhakad - Updated code and also updated Xpath
		 ********************************************************************/
		[UserCodeMethod]
		public static void AssignDeviceBaseForMultipleDevices(string DeviceLabel, string sBaseofDevice, string sBasePropertyRowIndex)
		{
			int iRowIndex;
			string sExistingBase;
			sBase = sBaseofDevice;
			sRowIndex = sBasePropertyRowIndex;
			sLabelName = DeviceLabel;
			repo.ProfileConsys1.PanelInvetoryGrid.LabelofDevice.Click();
			repo.ProfileConsys1.BaseofDeviceRow.Click();
			repo.ProfileConsys1.BaseofDeviceRow.PressKeys("{Right}");
			int.TryParse(sRowIndex, out iRowIndex);
			iRowIndex = iRowIndex+1;
			sRowIndex = iRowIndex.ToString();
			repo.ProfileConsys1.Cell_BaseofDevice.Click();
			sExistingBase = repo.ProfileConsys1.SomeText.TextValue;
			////sExistingBase = sExistingBase.Replace(@"\""",string.Empty);
			if(!sExistingBase.Equals(sBase))
			{
				repo.ProfileConsys1.BaseofDeviceRow.MoveTo("760;19");
				repo.ProfileConsys1.BaseofDeviceRow.Click("760;19");
				int.TryParse(sRowIndex, out iRowIndex);
				iRowIndex = iRowIndex-1;
				sRowIndex = iRowIndex.ToString();
				repo.ProfileConsys1.BaseofDeviceRow.MoveTo("760;19");
				repo.ProfileConsys1.BaseofDeviceRow.Click("760;19");
				
				if(repo.ContextMenu.btn_BaseSelectionInfo.Exists())
				{
					repo.ContextMenu.btn_BaseSelection.Click();
				}
				else
				{
					repo.ContextMenu.btn_Base_Selection_Multiple.Click();
				}
			}
		}


		
		/********************************************************************
		 * Function Name: SelectPointsGridRow
		 * Function Details: To select points grid row
		 * Parameter/Arguments: sRowNumber
		 * Output:
		 * Function Owner: Poonam Kadam
		 * Last Update : 31/12/2018
		 ********************************************************************/
		// Change cable resistance method
		[UserCodeMethod]
		public static void SelectPointsGridRow(string sRowNumber)
		{
			sRowIndex=sRowNumber;
			//Click on row from points grid
			repo.FormMe.PointsGridRow.Click();
		}

		/********************************************************************
		 * Function Name: AddDevicesfromMainProcessorGallery
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 03/01/2019  Updated on 28/01/2019 - Added Report Log
		 ********************************************************************/
		[UserCodeMethod]
		public static void AddDevicesfromMainProcessorGallery(string sDeviceName,string sType, string PanelType)
		{
			sMainProcessorGalleryIndex = SelectMainProcessorGalleryType(sType, PanelType);
			ModelNumber=sDeviceName;
			repo.FormMe.btn_MainProcessorGalleryDropDown.Click();
			repo.ContextMenu.txt_SelectDevice.Click();
			Report.Log(ReportLevel.Info, "Device "+sDeviceName+" added successfully");
		}
		
		
		/********************************************************************
		 * Function Name: SelectMainProcessorGalleryType
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 03/01/2018    Updated:23/01/2019 -  Shweta Bhosale-Added panel type check
		 ********************************************************************/
		[UserCodeMethod]
		public static string SelectMainProcessorGalleryType(string sType,string PanelType)
		{
			switch (sType)
			{
				case "Repeaters":
					sMainProcessorGalleryIndex="3";
					break;
				case "Loops":
					sMainProcessorGalleryIndex="4";
					break;
				case "Slot Cards":
					sMainProcessorGalleryIndex="5";
					break;
				case "Miscellaneous":
					if(PanelType.Equals("FIM"))
					{
						sMainProcessorGalleryIndex="5";
					}
					else
					{
						sMainProcessorGalleryIndex="6";
					}
					break;
				case "Printers":

					if(PanelType.Equals("FIM"))
					{
						sMainProcessorGalleryIndex="6";
					}
					else
					{
						sMainProcessorGalleryIndex="7";
					}
					break;
				case "Attached Functionality":
					if(PanelType.Equals("FIM"))
					{
						sMainProcessorGalleryIndex="7";
					}
					else
					{
						sMainProcessorGalleryIndex="8";
					}
					break;
				default:
					Console.WriteLine("Please specify correct gallery name");
					break;
			}
			return sMainProcessorGalleryIndex;
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
		 * Function Name: VerifyProgressBarIndicatorForISUnits
		 * Function Details: Verify progress bar indicator for Intrinsically-safe Units
		 * Parameter/Arguments: sFileName, sAddDevicesSheet, sAddIsDevicesSheet
		 * Output:
		 * Function Owner: Poonam Kadam
		 * Last Update : 12/31/2018
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
		 * Function Name: SelectInventoryGridRow
		 * Function Details: To select inventory grid row
		 * Parameter/Arguments: sRowNumber, sSkuNumber
		 * Output:
		 * Function Owner: Poonam Kadam
		 * Last Update : 2/1/2018
		 ********************************************************************/
		// Change cable resistance method
		[UserCodeMethod]
		public static void SelectInventoryGridRow(string sRowNumber)
		{
			sRow=sRowNumber;
			repo.FormMe.InventoryGridRow.Click();
			Report.Log(ReportLevel.Success, "Inventory grid row selected");
		}
		
		/********************************************************************
		 * Function Name: AddDevicefromPanelAccessoriesGallery
		 * Function Details: Add devices from panel accessories
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Poonam Kadam
		 * Last Update : 04/01/2019, 18/01/2019 - Alpesh Dhakad - Updated Report log
		 ********************************************************************/
		[UserCodeMethod]
		public static void AddDevicefromPanelAccessoriesGallery(string sDeviceName,string sType)
		{
			sAccessoriesGalleryIndex= SelectPanelAccessoriesGalleryType(sType);
			ModelNumber=sDeviceName;
			repo.FormMe.btn_PanelAccessoriesDropDown.Click();
			repo.ContextMenu.txt_SelectDevice.Click();
			Report.Log(ReportLevel.Success,"Panel Accessories device " + sDeviceName + " added successfully " );
		}
		
		/********************************************************************
		 * Function Name: SelectPanelAccessoriesGalleryType
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update : 23/01/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static string SelectPanelAccessoriesGalleryType(string sType)
		{
			switch (sType)
			{
				case "Accessories":
					sAccessoriesGalleryIndex="2";
					break;
				default:
					Console.WriteLine("Please specify correct gallery name");
					break;
			}
			return sAccessoriesGalleryIndex;
		}
		
		
		/*****************************************************************************************************************
		 * Function Name: additionOfDevicesOnEthernet
		 * Function Details: To add devices on main processor ethernet connection using excel test data
		 * Parameter/Arguments: filename and add devices sheet name
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 03/01/2019
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void additionOfDevicesOnEthernet(string sFileName,string sAddDevicesSheet,string PanelType)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared string type
			string sType,sLabelName;
			
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				
				AddDevicesfromMainProcessorGallery(ModelNumber,sType,PanelType);
				Report.Log(ReportLevel.Info, "Device "+ModelNumber+" added successfully");
			}
			//Close opened excel sheet
			Excel_Utilities.CloseExcel();
		}
		
		/*****************************************************************************************************************
		 * Function Name: additionOfDevicesOnRBus
		 * Function Details: To add devices on main processor RBus connection using excel test data
		 * Parameter/Arguments: Filename and add devices sheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 04/01/2019
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void additionOfDevicesOnRBus(string sFileName,string sAddDevicesSheet,string PanelType)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared string type
			string sType,sLabelName;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				
				AddDevicesfromMainProcessorGallery(ModelNumber,sType,PanelType);
				Report.Log(ReportLevel.Info, "Device "+ModelNumber+" added successfully");
			}
			//Close opened excel sheet
			Excel_Utilities.CloseExcel();
		}
		
		
		/*****************************************************************************************************************
		 * Function Name: additionOfDevicesOnXBus
		 * Function Details: To add XBus devices using excel test data using RBus devices
		 * Parameter/Arguments: File name and add devices sheet name
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 07/01/2019
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void additionOfDevicesForXBus(string sFileName,string sAddDevicesSheet,string PanelType)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared string type
			string sType,sLabelName;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				
				AddDevicesfromMainProcessorGallery(ModelNumber,sType,PanelType);
				Report.Log(ReportLevel.Info, "Device "+ModelNumber+" added successfully");
			}
			
			//Close opened excel sheet
			Excel_Utilities.CloseExcel();
		}
		
		
		/********************************************************************
		 * Function Name: VerifyAlarmLoad
		 * Function Details: To verify alarm load of sounder
		 * Parameter/Arguments:expected alarm load
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :19/2/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyAlarmLoad(string sAlarmLoad)
		{
			//Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Search Alarm Load property
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("Alarm Load" +"{ENTER}" );
			
			// Click on Alarm Load cell
			repo.ProfileConsys1.PARTItemsPresenter.cell_DeviceMode.Click();
			
			// Retrieve value alarm load
			string actualAlarmLoad = repo.ProfileConsys1.PARTItemsPresenter.txt_DayMatchesNight.TextValue;
			
			// Comparing DayMode and sDayMode values
			if(actualAlarmLoad.Equals(sAlarmLoad))
			{
				Report.Log(ReportLevel.Success,"Alarm Load " +sAlarmLoad+ " is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Alarm Load is not displayed as "+actualAlarmLoad+ " instead of "+ sAlarmLoad);
			}
			//Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
		}
		
		/********************************************************************
		 * Function Name: SelectRowUsingLabelName
		 * Function Details: To select item from grid using label
		 * Parameter/Arguments: sLabelName
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update : 21/2/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void SelectRowUsingLabelName(string sLabel)
		{
			sLabelName = sLabel;
			repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1.Click();
			Report.Log(ReportLevel.Success, "Device with Label name " + sLabel+" selected");
		}
		
		/********************************************************************
		 * Function Name: RightClickOnSelectedRow
		 * Function Details: To open the context menu options using right click
		 * Parameter/Arguments: RowNumber
		 * Output:
		 * Function Owner: Purvi
		 * Last Update : 11/4/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void RightClickOnSelectedRow(string RowNumber)
		{
			sRowIndex = RowNumber;
			repo.FormMe.PointsGridRow.Click(System.Windows.Forms.MouseButtons.Right);
			
		}
		
		/********************************************************************
		 * Function Name: DeleteDevicesPresentInCustomGallery
		 * Function Details: To delete devices from custom gallery
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta
		 * Last Update : 11/4/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void DeleteDevicesPresentInCustomGallery()
		{
			if(repo.FormMe.CustomGalleryInfo.Exists())
			{
				repo.FormMe.btn_CustomDevices.Click();
				int numberOfCustomDevices= repo.CustomDevices.ListBox.Children.Count;
				for(int i=0;i<=numberOfCustomDevices;i++)
				{
					sListIndex=i.ToString();
					repo.CustomDevices.CustomGalleyListItem.Click();
					repo.CustomDevices.CustomGalleyListItem.MoveTo("510;278");
					repo.CustomDevices.DeleteButtonforCustom.Click();
				}
			}
			else
			{
				Report.Log(ReportLevel.Info, "Custom devices are not present in gallery");
			}
		}
		/********************************************************************
		 * Function Name: VerifyDeviceUsingLabelName
		 * Function Details: To verify item with label name
		 * Parameter/Arguments: sLabelName
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update : 12/3/2019
		 ********************************************************************/
		// Change cable resistance method
		[UserCodeMethod]
		public static void VerifyDeviceUsingLabelName(string sLabel)
		{
			sLabelName = sLabel;
			if(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info.Exists())
			{
				Report.Log(ReportLevel.Success,"Device with label name " +sLabel+ " is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Device with label name " +sLabel+ " not is displayed correctly");
			}
		}
		/********************************************************************
		 * Function Name: changeAndVerifyAlarmLoad
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void changeAndVerifyAlarmLoad(int AlarmLoad, string rangeState, int expectedResult)
		{
			
			int Value,actualValue,revertTo;
			string actualAlarmLoad;
			//Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Search Alarm Load property
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("Alarm Load" +"{ENTER}" );
			
			// Click on Alarm Load cell
			repo.ProfileConsys1.PARTItemsPresenter.cell_DeviceMode.Click();
			
			
			if((rangeState.Equals("Valid")))
			{
				Keyboard.Press("{LControlKey down}{Akey}{LControlKey up}"+AlarmLoad +"{ENTER}");
				
				// Click on Alarm Load cell
				repo.ProfileConsys1.PARTItemsPresenter.cell_DeviceMode.Click();
				
				// Retrieve value alarm load
				actualAlarmLoad = repo.ProfileConsys1.PARTItemsPresenter.txt_DayMatchesNight.TextValue;
				int.TryParse(actualAlarmLoad, out Value);
				if(Value==AlarmLoad)
				{
					Report.Log(ReportLevel.Success,"Number of Alarm Load is changed to: "+AlarmLoad);
				}
				else
				{
					Report.Log(ReportLevel.Failure,"Number of Alarm Load is not changed to: "+AlarmLoad);
				}
			}
			else if((rangeState.Equals("InvalidRange")))
			{
				string initialValue = repo.ProfileConsys1.PARTItemsPresenter.txt_DayMatchesNight.TextValue;
				int.TryParse(initialValue,out revertTo);
				Keyboard.Press("{LControlKey down}{Akey}{LControlKey up}"+AlarmLoad +"{ENTER}");
				// Click on Alarm Load cell
				repo.ProfileConsys1.PARTItemsPresenter.cell_DeviceMode.Click();
				
				string revertedValue = repo.ProfileConsys1.PARTItemsPresenter.txt_DayMatchesNight.TextValue;
				int.TryParse(revertedValue, out actualValue);
				if(actualValue==revertTo)
				{
					Report.Log(ReportLevel.Success,"Number of Alarm Load is reverted to: "+revertTo);
				}
				else
				{
					Report.Log(ReportLevel.Failure,"Number of Alarm Load is not reverted to: "+revertTo);
				}
			}
			else
			{
				
				Keyboard.Press("{LControlKey down}{Akey}{LControlKey up}"+AlarmLoad +"{ENTER}");
				// Click on Alarm Load cell
				repo.ProfileConsys1.PARTItemsPresenter.cell_DeviceMode.Click();
				
				string revertedValue = repo.ProfileConsys1.PARTItemsPresenter.txt_DayMatchesNight.TextValue;
				int.TryParse(revertedValue,out actualValue);
				if(actualValue==expectedResult)
				{
					Report.Log(ReportLevel.Success,"Number of Alarm Load is reverted to: "+expectedResult);
				}
				else
				{
					Report.Log(ReportLevel.Failure,"Number of Alarm Load is not reverted to: "+expectedResult);
				}
				
			}
			
		}
		
		/********************************************************************
		 * Function Name: verifyMinMaxThroughSpinControlForAlarmLoad
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void verifyMinMaxThroughSpinControlForAlarmLoad(string minLimit,string maxLimit)
		{
			// Click on Alarm Load cell
			repo.ProfileConsys1.PARTItemsPresenter.cell_DeviceMode.Click();
			
			Keyboard.Press("{LControlKey down}{Akey}{LControlKey up}"+ maxLimit +"{ENTER}");
			
			// Click on Alarm Load cell
			repo.ProfileConsys1.PARTItemsPresenter.cell_DeviceMode.Click();
			
			repo.FormMe.AlarmLoadSpinUpButton.Click();
			string actualValue = repo.ProfileConsys1.PARTItemsPresenter.txt_DayMatchesNight.TextValue;
			if(actualValue.Equals(maxLimit))
			{
				Report.Log(ReportLevel.Success,"Spin control accepts values within specified max limit");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Spin control doesnot accepts values within specified max limit");
			}
			Keyboard.Press("{ENTER}");
			
			Keyboard.Press("{LControlKey down}{Akey}{LControlKey up}"+ minLimit +"{ENTER}");
			
			// Click on Alarm Load cell
			repo.ProfileConsys1.PARTItemsPresenter.cell_DeviceMode.Click();
			
			repo.FormMe.AlarmLoadSpinDownButton.Click();
			actualValue = repo.ProfileConsys1.PARTItemsPresenter.txt_DayMatchesNight.TextValue;
			if(actualValue.Equals(minLimit))
			{
				Report.Log(ReportLevel.Success,"Spin control accepts values within specified min limit");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Spin control does not accepts values within specified min limit");
			}
		}
		
		
		/********************************************************************
		 * Function Name: AddDevicesfromMultiplePointWizard
		 * Function Details: To add multiple devices using multiple point wizard
		 * Parameter/Arguments: Device name and its quantity
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 29/03/2019 Alpesh Dhakad- Updated btn_MultiplePointWizard xpath and change script accordingly
		 ********************************************************************/
		[UserCodeMethod]
		public static void AddDevicesfromMultiplePointWizard(string sDeviceName,int DeviceQty )
		{
			repo.FormMe.btn_MultiplePointWizard.Click();
			//repo.ProfileConsys1.btn_MultiplePointWizard_DoNotUse.Click();
			repo.AddDevices.txt_AllDevices.Click();
			
			repo.AddDevices.txt_SearchDevices.Click();
			Keyboard.Press("{LControlKey down}{Akey}{LControlKey up}"+sDeviceName);
			ModelNumber = sDeviceName;
			repo.AddDevices.txt_ModelNumber.Click();
			repo.AddDevices.txt_Quantity.Click();
			Keyboard.Press("{LControlKey down}{Akey}{LControlKey up}"+DeviceQty.ToString());
			
			repo.AddDevices.btn_AddDevices.Click();
			Report.Log(ReportLevel.Success,+DeviceQty+" \""+sDeviceName+ "\" Device Added successfully");
			Delay.Milliseconds(200);

		}
		
		/***********************************************************************************************************
		 * Function Name: CreateProject
		 * Function Details: To create project with different market with List Index Mentioned in the argument
		 * Parameter/Arguments: sMarket, iListIndex
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 05/03/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void CreateProject(string sMarket,int iListIndex )
		{
			repo.ProfileConsys1.File.Click();
			
			repo.ProfileConsys1.TextNew.Click();
			
			repo.ProfileConsys1.PARTRight.txt_MarketNew.Click();
			Keyboard.Press(sMarket);
			
			sListIndex=iListIndex.ToString();
			repo.ProfileConsys1.lst_Market.Click();
			
			repo.ProfileConsys1.PARTRight.btn_CreateNewProject.Click();
			
			// Enter Project name
			repo.CreateNewProject.CreateNewProjectContainer.txt_ProjectName.Click();
			Delay.Duration(1000, false);
			Keyboard.Press("Verify");
			
			// Enter text in Client Name field
			repo.CreateNewProject.CreateNewProjectContainer.txt_ClientName.Click();
			Keyboard.Press("JCI");
			
			// Enter text Client Address field
			repo.CreateNewProject.CreateNewProjectContainer.txt_ClientAddress.Click();
			Keyboard.Press("JCI");
			
			// Enter text in Installer Name field
			repo.CreateNewProject.CreateNewProjectContainer.txt_InstallerName.Click();
			Keyboard.Press("JCI");
			
			// Enter text in Installer Address field
			repo.CreateNewProject.CreateNewProjectContainer.txt_InstallerAddress.Click();
			Keyboard.Press("JCI");

			// Click on Ok Button
			repo.CreateNewProject.CreateNewProjectContainer.btn_OK.Click();


		}
		
		/***********************************************************************************************************
		 * Function Name: verifyShoppingListDevices
		 * Function Details: To verify shopping list devices via clicking on its row
		 * Parameter/Arguments: sFileName,sDeviceSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 07/03/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyShoppingListDevices(string sFileName, string sDeviceSheet )
		{
			Excel_Utilities.OpenExcelFile(sFileName,sDeviceSheet);
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			string sDeviceName,sType,sLabelName;
			
			for(int i=8; i<=rows; i++)
			{
				sDeviceName=  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				sRow=  ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				
				repo.FormMe.ShoppingListDevices.Click();
				
				repo.FormMe.ShoppingListDevices.EnsureVisible();
				Report.Log(ReportLevel.Info,"Device is successfully displayed in shopping list as " +sDeviceName);
				
				
			}
			Excel_Utilities.CloseExcel();
			
		}
		
		
		/***********************************************************************************************************
		 * Function Name: verifyShoppingList
		 * Function Details: To verify Shopping List row count
		 * Parameter/Arguments: ShoppingListDeviceCount
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 07/03/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyShoppingList(int ShoppingListDeviceCount)
		{
			int ActualShoppingListDeviceCount = repo.FormMe.ShoppingListContainer.Children.Count();

			if(ActualShoppingListDeviceCount.Equals(ShoppingListDeviceCount))
			{
				Report.Log(ReportLevel.Success,"Devices is displayed correctly in shopping list");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Devices not displayed correctly in shopping list");
			}
			
		}
		
		/***********************************************************************************************************
		 * Function Name: AddDeviceOrderColumn
		 * Function Details: To add device order column
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 07/03/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void AddDeviceOrderColumn()
		{
			
			// Select Point grid and right click on it
			Mouse.Click(repo.FormMe.HeadersPanel, System.Windows.Forms.MouseButtons.Right);
			
			// Click Show column chooser to select column
			repo.ContextMenu.ShowColumnChooser.Click();
			
			// Click on Device order checkbox to add the column in points grid
			repo.ProfileConsys.chkBox_DeviceOrder.Click();
			Report.Log(ReportLevel.Info," Device order column added successfully ");
			
			// Close column choose window
			repo.ProfileConsys.btn_CloseColumnChooser.Click();
			
		}
		
		/***********************************************************************************************************
		 * Function Name: VerifyDeviceOrder
		 * Function Details: To add device order column
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 04/04/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyDeviceOrder(string DeviceOrderRow, string DeviceOrder, bool Present)
		{
			sDeviceOrderRow = DeviceOrderRow;
			repo.FormMe.DeviceOrder_txt.Click();
			string ActualDeviceOrder = repo.FormMe.DeviceOrder_txt.TextValue;
			if(Present)
			{
				if(ActualDeviceOrder.Equals(DeviceOrder))
				{
					Report.Log(ReportLevel.Success,"Device Order is displayed correctly");
				}
				else
				{
					Report.Log(ReportLevel.Failure,"Device Order not displayed correctly");
				}
			}
			else
			{
				if(ActualDeviceOrder.Equals(DeviceOrder))
				{
					Report.Log(ReportLevel.Failure,"Wrong Device Order is present");
				}
				else
				{
					Report.Log(ReportLevel.Success,"Device Order not displayed");
				}
			}
			
			
		}
		
		
		/***********************************************************************************************************
		 * Function Name: EditAlarmLoad
		 * Function Details: EditAlarmLoad
		 * Parameter/Arguments: Alarm Load to be entered
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update : 11/03/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void EditAlarmLoad(string AlarmLoad)
		{
			
			//Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Search Alarm Load property
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("Alarm Load" +"{ENTER}" );
			
			// Click on Alarm Load cell
			repo.ProfileConsys1.PARTItemsPresenter.cell_DeviceMode.Click();
			
			repo.ProfileConsys1.PARTItemsPresenter.cell_DeviceMode.PressKeys(AlarmLoad);
			
		}
		
		/***********************************************************************************************************
		 * Function Name: VerifyAlarmCurrentLoadProperty
		 * Function Details: VerifyAlarmCurrentLoad property
		 * Parameter/Arguments: IsReadOnly
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update : 12/03/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyAlarmCurrentLoadProperty(bool isReadOnly)
		{
			
			bool result = repo.FormMe.AlarmCurrent.Enabled;
			if(result==isReadOnly)
			{
				Report.Log(ReportLevel.Success,"AlarmCurrent Property is as expected");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"AlarmCurrent property is "+ result + " instead of" +isReadOnly);
			}
			
		}
		

		/***********************************************************************************************************
		 * Function Name: verifyPointsGridColumn
		 * Function Details: To verify points grid columns text
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 15/03/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyPointsGridColumn(string expectedColumnText, string ColumnNumber)
		{
			// Click on points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			// Set sColumn value
			sColumn = ColumnNumber;
			
			// Retrieve Column text value
			string actualColumnText = repo.FormMe.PointsGridHeaderColumn.TextValue;
			
			// Compare actual and expected column value
			if(actualColumnText.Equals(expectedColumnText))
			{
				Report.Log(ReportLevel.Success,"Column is "+expectedColumnText+ " which is displayed correctly in points grid");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Column is "+actualColumnText+ " which is displayed incorrectly instead of " +expectedColumnText+ " in points grid ");
			}
			
		}
		/***********************************************************************************************************
		 * Function Name: verifyDescription
		 * Function Details: To verify description from properties
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 13/03/2019  24/05/2019 - Updated script, added if statement for tab_Points and cell_properties
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyDescription(string sDescription)
		{
			if(repo.ProfileConsys1.tab_PointsInfo.Exists())
			{
				//Click on Points tab
				repo.ProfileConsys1.tab_Points.Click();
			}
			
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Search Alarm Load property
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("Description" +"{ENTER}" );
			
			
			
			if(repo.ProfileConsys1.PARTItemsPresenter.cell_DeviceModeInfo.Exists())
			{
				// Click on Alarm Load cell
				repo.ProfileConsys1.PARTItemsPresenter.cell_DeviceMode.Click();
			}
			else
			{
				// Click on Properties cell
				repo.FormMe.cell_Properties.Click();
				
			}

			if(repo.ProfileConsys1.PARTItemsPresenter.txt_DayMatchesNightInfo.Exists())
			{
				// Retrieve value alarm load
				string actualDescription = repo.ProfileConsys1.PARTItemsPresenter.txt_DayMatchesNight.TextValue;
				
				// Comparing DayMode and sDayMode values
				if(actualDescription.Equals(sDescription))
				{
					Report.Log(ReportLevel.Success,"Description " +sDescription+ " is displayed correctly");
				}
				else
				{
					Report.Log(ReportLevel.Failure,"Description is displayed as "+actualDescription+ " instead of "+ sDescription);
				}
				
			}
			else
			{
				string actualDescription =	repo.FormMe.txt_PropertiesTextValue.TextValue;
				
				// Comparing DayMode and sDayMode values
				if(actualDescription.Equals(sDescription))
				{
					Report.Log(ReportLevel.Success,"Description " +sDescription+ " is displayed correctly");
				}
				else
				{
					Report.Log(ReportLevel.Failure,"Description is displayed as "+actualDescription+ " instead of "+ sDescription);
				}
				
				
			}
			
			
			
			if(repo.ProfileConsys1.tab_PointsInfo.Exists())
			{
				//Click on Points tab
				repo.ProfileConsys1.tab_Points.Click();
			}
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
		}
		
		
		
		/***********************************************************************************************************
		 * Function Name: verifyAdditionOfSoundersInLPS800
		 * Function Details: To verify sounder properties i.e. Description and alarm load from Properties
		 * Parameter/Arguments: filename and device sheetname
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 14/03/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyAdditionOfSoundersInLPS800(string sFileName, string sDeviceSheet)
		{
			repo.ProfileConsys1.tab_Points.Click();
			
			Excel_Utilities.OpenExcelFile(sFileName,sDeviceSheet);
			
			// Count the number of rows in excel
			int rows = Excel_Utilities.ExcelRange.Rows.Count;
			
			string sType,expectedDescription,expectedAlarmLoad,sourceDeviceIndex,targetDeviceIndex;
			
			for(int i=9; i<=rows; i++)
			{
				string sDeviceName = ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				expectedDescription = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				expectedAlarmLoad = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				
				repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1.Click();
				
				verifyDescription(expectedDescription);
				
				VerifyAlarmLoad(expectedAlarmLoad);
				
			}
			
			//Click on Physical Layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			//Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			//Click on Physical Layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			// Retrieve value of sourceDeviceIndex & targetDeviceIndex from sheet
			sourceDeviceIndex= ((Range)Excel_Utilities.ExcelRange.Cells[3,6]).Value.ToString();
			targetDeviceIndex= ((Range)Excel_Utilities.ExcelRange.Cells[4,6]).Value.ToString();
			
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
			
			// Close Excel sheet
			Excel_Utilities.CloseExcel();
			
			
		}
		
		
		
		/***********************************************************************************************************
		 * Function Name: verifySavedOpenProjectOnLPS800Addition
		 * Function Details: To verify saved open project on LPS800 addition
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 13/03/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void verifySavedOpenProjectOnLPS800Addition(string sFileName, string sDeviceSheet)
		{
			// Select Point grid and right click on it
			Mouse.Click(repo.FormMe.HeadersPanel, System.Windows.Forms.MouseButtons.Right);
			
			// Click Show column chooser to select column
			repo.ContextMenu.ShowColumnChooser.Click();
			
			// Click on Device order checkbox to add the column in points grid
			repo.ProfileConsys.chkBox_DeviceOrder.Click();
			Report.Log(ReportLevel.Info," Device order column added successfully ");
			
			// Close column choose window
			repo.ProfileConsys.btn_CloseColumnChooser.Click();
			
			// Open Another excel to verify IS devices state
			Excel_Utilities.OpenExcelFile(sFileName,sDeviceSheet);
			
			// Count the number of rows in excel
			int rows = Excel_Utilities.ExcelRange.Rows.Count;
			
			// Set sDeviceOrderRow value
			sDeviceOrderRow = (1).ToString();
			
			// Retrieve LPS800DeviceOrder value from excel sheet
			string LPS800DeviceOrder = ((Range)Excel_Utilities.ExcelRange.Cells[8,7]).Value.ToString();
			
			// Click on Device order label
			repo.FormMe.DeviceOrder_txt.Click();
			
			// Retrieve value from Device order
			string actualLPSDeviceOrderValue = repo.FormMe.DeviceOrder_txt.TextValue;
			
			// Compare actualDeviceOrderValue and sDeviceOrderName values and then displaying result
			if(actualLPSDeviceOrderValue.Equals(LPS800DeviceOrder))
			{
				Report.Log(ReportLevel.Success, "Device " +sDeviceName+ " with " +LPS800DeviceOrder+ " added successfully and displaying correct device order");
			}
			else
			{
				Report.Log(ReportLevel.Failure, "Device " +sDeviceName+ " with " +LPS800DeviceOrder+ " not added or not displaying correct device order");
			}
			
			
			for(int i=9; i<=rows; i++)
			{
				sDeviceOrderRow= (i-6).ToString();
				string changedDeviceOrder = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				
				// Click on Device Order
				repo.FormMe.DeviceOrder_txt.Click();
				
				// Retrieve value from Device order
				string actualDeviceOrderValue = repo.FormMe.DeviceOrder_txt.TextValue;
				
				// Compare actualDeviceOrderValue and sDeviceOrderName values and then displaying result
				if(actualDeviceOrderValue.Equals(changedDeviceOrder))
				{
					Report.Log(ReportLevel.Success, "Device " +sDeviceName+ " with " +changedDeviceOrder+ " added successfully and displaying correct device order");
				}
				else
				{
					Report.Log(ReportLevel.Failure, "Device " +sDeviceName+ " with " +changedDeviceOrder+ " not added or not displaying correct device order");
				}
				
			}
			// Close Excel sheet
			Excel_Utilities.CloseExcel();
		}
		
		
		/***********************************************************************************************************
		 * Function Name: verifyDeletionOfSoundersInLPS800
		 * Function Details: To verify deletion of Sounders in LPS800
		 * Parameter/Arguments: filename and device sheetname
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 14/03/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyDeletionOfSoundersInLPS800(string sFileName, string sDeviceSheet, string sDeleteDeviceSheet)
		{
			// Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			// Open excel sheet
			Excel_Utilities.OpenExcelFile(sFileName,sDeviceSheet);
			
			// Count the number of rows in excel
			int rows = Excel_Utilities.ExcelRange.Rows.Count;
			
			string sType,expectedDescription,expectedAlarmLoad;
			
			for(int i=8; i<=rows; i++)
			{
				string sDeviceName = ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				expectedDescription = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				expectedAlarmLoad = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				
				repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1.Click();
				
				verifyDescription(expectedDescription);
				
				VerifyAlarmLoad(expectedAlarmLoad);
				
			}
			// Close Excel sheet
			Excel_Utilities.CloseExcel();
			
			// Open Excel sheet
			Excel_Utilities.OpenExcelFile(sFileName,sDeleteDeviceSheet);
			
			// Count the number of rows in excel
			rows = Excel_Utilities.ExcelRange.Rows.Count;
			
			for(int i=8; i<=rows; i++)
			{
				string sDeviceName = ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				
				repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1.Click();
				repo.ProfileConsys1.btn_Delete.Click();
				
				Validate.AttributeEqual(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info, "Text", sLabelName);
				Report.Log(ReportLevel.Success, "Device "+sLabelName+" deleted successfully");
				
				
			}
			// Close Excel sheet
			Excel_Utilities.CloseExcel();
		}
		
		
		/***********************************************************************************************************
		 * Function Name: verifyPointGridRowCount
		 * Function Details: To verify points grid row count
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 14/03/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyPointGridRowCount(int ExpectedPointGridRowCount)
		{
			// Retrieve Point grid rows count
			int ActualPointGridRowCount = repo.FormMe.PointsGridContainer.Children.Count();

			// Compare Actual and Expected Point Grid Row count
			if(ActualPointGridRowCount.Equals(ExpectedPointGridRowCount))
			{
				Report.Log(ReportLevel.Success,"Devices is displayed correctly in Points grid");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Devices not displayed correctly in Points grid");
			}
			
		}
		
		
		
		/***********************************************************************************************************
		 * Function Name: verifySoundersDeviceOrderInLPS800
		 * Function Details: To verify Sounders DeviceOrder In LPS800
		 * Parameter/Arguments: filename and device sheetname
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 15/03/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void verifySoundersDeviceOrderInLPS800(string sFileName, string sDeviceSheet)
		{
			Excel_Utilities.OpenExcelFile(sFileName,sDeviceSheet);
			
			// Count the number of rows in excel
			int rows = Excel_Utilities.ExcelRange.Rows.Count;
			
			string sType;
			
			for(int i=8; i<=rows; i++)
			{
				sDeviceOrderRow= (i-3).ToString();
				string sDeviceName = ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				string changedDeviceOrder = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				
				// Click on Device Order
				repo.FormMe.txt_DeviceOrderLabel.Click();
				
				// Retrieve Device order value
				string actualDeviceOrderValue = repo.FormMe.txt_DeviceOrderLabel.TextValue;
				
				// Compare actualDeviceOrderValue and sDeviceOrderName values and then displaying result
				if(actualDeviceOrderValue.Equals(changedDeviceOrder))
				{
					Report.Log(ReportLevel.Success, "Device " +sDeviceName+ " with " +changedDeviceOrder+ " added successfully and displaying correct device order");
				}
				else
				{
					Report.Log(ReportLevel.Failure, "Device " +sDeviceName+ " with " +changedDeviceOrder+ " not added or not displaying correct device order");
				}
				
				// Click on Label name in points grid
				repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1.Click();
				
				// Retrieve label name
				string actualLabelName = repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1.TextValue;
				
				// Compare actualDeviceOrderValue and sDeviceOrderName values and then displaying result
				if(actualLabelName.Equals(sLabelName))
				{
					Report.Log(ReportLevel.Success, "Device " +sDeviceName+ " with " +sLabelName+ " displaying correctly in correct device order");
				}
				else
				{
					Report.Log(ReportLevel.Failure, "Device " +sDeviceName+ " with " +sLabelName+ " not displaying correctly in correct device order");
				}
				
			}
			// Close excel sheet
			Excel_Utilities.CloseExcel();
		}
		
		
		/********************************************************************
		 * Function Name: AddISDevicesfromMultiplePointWizard
		 * Function Details: To verify IS devices present in Multiple Point Wizard
		 * Parameter/Arguments: Device name
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 15/03/2019 and 29/03/2019 Alpesh Dhakad- Updated btn_MultiplePointWizard xpath and change script accordingly
		 ********************************************************************/
		[UserCodeMethod]
		public static void AddISDevicesfromMultiplePointWizard(string sDeviceName )
		{
			repo.FormMe.btn_MultiplePointWizard.Click();
			//repo.ProfileConsys1.btn_MultiplePointWizard_DoNotUse.Click();
			repo.AddDevices.txt_AllDevices.Click();
			
			repo.AddDevices.txt_SearchDevices.Click();
			Keyboard.Press("{LControlKey down}{Akey}{LControlKey up}"+sDeviceName);
			ModelNumber = sDeviceName;
			if(repo.AddDevices.txt_ModelNumberInfo.Exists())
			{
				Report.Log(ReportLevel.Failure,"IS Devices are presnt in Multiple Point Wizard");
				
			}
			else
			{
				Report.Log(ReportLevel.Success,"IS Devices are absent in Multiple Point Wizard");
				
			}
			
			Delay.Milliseconds(200);

		}

		/********************************************************************
		 * Function Name: VerifyISDevicesGetPasted
		 * Function Details: To verify IS devices get pasted only on Exi800
		 * Parameter/Arguments: Device name
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 18/03/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyISDevicesGetPasted(bool sExi800Selected )
		{
			if(sExi800Selected)
			{
				if(repo.FormMe.PasteInfo.Exists())
				{
					
					Report.Log(ReportLevel.Success,"IS Devices can be pasted in Exi800");
					repo.FormMe.Paste.Click();
				}
				else
				{
					Report.Log(ReportLevel.Failure,"IS Devices cannot be pasted in Exi800");
					
				}
			}
			
			else
			{
				if(repo.FormMe.PasteInfo.Exists())
				{
					Report.Log(ReportLevel.Failure,"IS Devices can be pasted in Exi800");
					
				}
				else
					
				{
					Report.Log(ReportLevel.Success,"IS Devices cannot be pasted in non-Ex devices");
					repo.FormMe.Paste.Click();
					
				}
			}
		}
		
		/********************************************************************
		 * Function Name:
		 * Function Details:
		 * Parameter/Arguments: boolean value
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void verifyIsolatorCheckbox(bool ExpectedState)
		{
			//Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Enter the Day Matches night text in Search Properties fields to view day matches night related text;
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("Isolator" +"{ENTER}" );

			// Click on Isolator  cell
			repo.ProfileConsys1.PARTItemsPresenter.cell_DayMatchesNight.Click();
			
			// To retrieve the attribute value as boolean by its ischecked properties and store in actual state
			bool actualState =  repo.ProfileConsys1.chkbox_Isolator.GetAttributeValue<bool>("ischecked");
			
			//As per actual state and expected state values
			if(actualState.Equals(ExpectedState))
			{
				Report.Log(ReportLevel.Success, "Isolator device is displayed as expected ");
			}
			else
			{
				Report.Log(ReportLevel.Failure, "Isolator device is not displayed as expected ");
			}
			
			//Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
		}
		

		/********************************************************************
		 * Function Name: verifyTripCurrentOnBaseChange
		 * Function Details: To verify trip current details on base change and isolator properties check
		 * Parameter/Arguments: sfilename and sheetname
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void verifyTripCurrentOnBaseChange(string sFileName, string sDeviceSheet)
		{
			// Open Excel sheet
			Excel_Utilities.OpenExcelFile(sFileName,sDeviceSheet);
			
			// Count the number of rows in excel
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			string sType,expectedDCUnits,expectedDCUnitsAfterBaseChange,expectedIsolatorState;
			bool expectedIsolatorCheckboxState;
			
			for (int i=8; i<=rows; i++)
			{
				sDeviceName = ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				expectedDCUnits = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				expectedDCUnitsAfterBaseChange = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				
				expectedIsolatorState = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				
				
				sBase = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				sRowIndex= ((Range)Excel_Utilities.ExcelRange.Cells[i,10]).Value.ToString();
				
				bool.TryParse(expectedIsolatorState, out expectedIsolatorCheckboxState);
				
				
				Devices_Functions.AddDevicesfromGallery(sDeviceName,sType);
				verifyIsolatorCheckbox(expectedIsolatorCheckboxState);
				
				DC_Functions.verifyDCUnitsValue(expectedDCUnits);
				
				//AssignDeviceBase(sLabelName,sBase,sRowIndex);
				
				AssignDeviceBaseForMultipleDevices(sLabelName,sBase,sRowIndex);
				
				Report.Log(ReportLevel.Info, "Base " + sBase + " assigned to "+ sLabelName);
				
				DC_Functions.verifyDCUnitsValue(expectedDCUnitsAfterBaseChange);
			}
			//Close Excel sheet
			Excel_Utilities.CloseExcel();
		}
		
		/********************************************************************
		 * Function Name: VerifyBaseAfterReopening
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyBaseAfterReopening(string DeviceLabel, string sBaseName)
		{
			sLabelName = DeviceLabel;
			repo.ProfileConsys1.tab_Points.Click();
			repo.ProfileConsys1.PanelInvetoryGrid.LabelofDevice.Click();
			repo.ProfileConsys1.Cell_BaseofDevice.Click();
			string sActualBaseName = repo.ProfileConsys1.Cell_BaseofDevice.Text;
			if(sActualBaseName.Equals(sBaseName))
			{
				Report.Log(ReportLevel.Success,"Base value is retained");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Base value is not retained");
			}
			
		}
		
		
		/********************************************************************
		 * Function Name: VerifyBaseIsVisibleInList
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyBaseIsVisibleInList(string DeviceLabel, string sBaseofDevice, string sBasePropertyRowIndex,bool IsVisible)
		{
			int iRowIndex;
			sBase = sBaseofDevice;
			sRowIndex = sBasePropertyRowIndex;
			sLabelName = DeviceLabel;
			repo.ProfileConsys1.PanelInvetoryGrid.LabelofDevice.Click();
			repo.ProfileConsys1.BaseofDeviceRow.Click();
			repo.ProfileConsys1.BaseofDeviceRow.PressKeys("{Right}");
			int.TryParse(sRowIndex, out iRowIndex);
			iRowIndex = iRowIndex+1;
			sRowIndex = iRowIndex.ToString();
			repo.ProfileConsys1.Cell_BaseofDevice.Click();
			repo.ProfileConsys1.BaseofDeviceRow.MoveTo("760;19");
			repo.ProfileConsys1.BaseofDeviceRow.Click("760;19");
			int.TryParse(sRowIndex, out iRowIndex);
			iRowIndex = iRowIndex-1;
			sRowIndex = iRowIndex.ToString();
			repo.ProfileConsys1.BaseofDeviceRow.MoveTo("760;19");
			repo.ProfileConsys1.BaseofDeviceRow.Click("760;19");
			bool state = repo.ContextMenu.btn_BaseSelection.Visible;
			if(IsVisible)
			{
				if(state.Equals(IsVisible))
				{
					Report.Log(ReportLevel.Success,"Base "+ sBase +" exist in list");
				}
				else
				{
					Report.Log(ReportLevel.Failure,"Base "+ sBase +" doesn't exist in list");
				}
			}
			else
			{
				if(state.Equals(IsVisible))
				{
					Report.Log(ReportLevel.Failure,"Base "+ sBase +"  exist in list");
				}
				else
				{
					Report.Log(ReportLevel.Success,"Base "+ sBase +" doesn't exist in list");
				}
				
			}
		}
		
		/*****************************************************************************************************************
		 * Function Name:verifySwitchingAllowedPowerSource
		 * Function Details: To verify Switching Allowed PowerSource
		 * Parameter/Arguments: filename and sheetname
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 29/03/2019
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verifySwitchingAllowedPowerSource(string sFileName,string sAddDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared variables type
			string sDeviceName,sType,PanelType,sLabel,sPowerSupply,sChangePowerSupply;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				sDeviceName=  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sLabel=  ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				sPowerSupply = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sChangePowerSupply = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				
				PanelType= ((Range)Excel_Utilities.ExcelRange.Cells[5,5]).Value.ToString();
				
				AddDevicesfromMainProcessorGallery(sDeviceName,sType,PanelType);
				
				SelectRowUsingLabelName(sLabel);
				
				// To verify power supply value
				VerifyPowerSupply(sPowerSupply);
				
				// To change power supply value
				ChangePowerSupply(sChangePowerSupply);
			}
			Excel_Utilities.CloseExcel();
		}
		
		
		
		
		/********************************************************************
		 * Function Name: VerifyPowerSupply
		 * Function Details: To verify power supply value
		 * Parameter/Arguments: sPowerSupply
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 30/03/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyPowerSupply(string sPowerSupply)
		{
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Enter the Device text in Search Properties fields to view device related text
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("Power Supply" +"{ENTER}");
			
			// Click on Power Supply cell
			repo.FormMe.cell_PowerSupply.Click();
			
			
			// Get the text value of Device Sensitivity field
			string PowerSupply = repo.FormMe.txt_PowerSupply.Text;
			
			//Comparing expected and actual Device Sensitivity value
			if(PowerSupply.Equals(sPowerSupply))
			{
				Report.Log(ReportLevel.Success,"Power supply value " +PowerSupply + " is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Power supply value " +PowerSupply+ " is not displayed correctly");
			}

			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
			
		}

		

		
		/********************************************************************
		 * Function Name: VerifyPowerSupplyAfterReopenProject
		 * Function Details: To verify power supply value  after Reopening project
		 * Parameter/Arguments: sPowerSupply
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 30/03/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyPowerSupplyAfterReopenProject(string sPowerSupply)
		{
			// Click on SearchProperties text field
			repo.FormMe.txt_SearchProperties_AfterReopen.Click();
			
			// Enter the Device text in Search Properties fields to view device related text
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("Power Supply" +"{ENTER}");
			
			// Click on Power Supply cell
			repo.FormMe.cell_PowerSupply.Click();
			
			
			// Get the text value of Device Sensitivity field
			string PowerSupply = repo.FormMe.txt_PowerSupply.Text;
			
			//Comparing expected and actual Device Sensitivity value
			if(PowerSupply.Equals(sPowerSupply))
			{
				Report.Log(ReportLevel.Success,"Power supply value " +PowerSupply + " is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Power supply value " +PowerSupply+ " is not displayed correctly");
			}

			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
			
		}
		
		/********************************************************************
		 * Function Name: ChangePowerSupply
		 * Function Details: To change power supply value
		 * Parameter/Arguments: ChangePowerSupply value
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 30/03/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void ChangePowerSupply(string sChangePowerSupply)
		{
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Enter the Device text in Search Properties fields to view Power supply related text
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("Power Supply" +"{ENTER}" );
			
			// Click on Power Supply cell
			repo.FormMe.cell_PowerSupply.Click();
			
			// Enter the value to change Power Supply value
			repo.FormMe.txt_PowerSupply.PressKeys((sChangePowerSupply) +"{ENTER}" + "{ENTER}");
			
			// Click on Power Supply cell
			repo.FormMe.cell_PowerSupply.Click();
			
			// Get the text value of Device Sensitivity field
			string PowerSupply = repo.FormMe.txt_PowerSupply.Text;
			
			//Comparing expected and actual Device Sensitivity value
			if(PowerSupply.Equals(sChangePowerSupply))
			{
				Report.Log(ReportLevel.Success,"Power supply value  changed successfully to " +PowerSupply + " is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Power supply value is not changed to " +sChangePowerSupply+ " and displayed incorrectly");
			}

			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");

		}
		
		
		/********************************************************************
		 * Function Name: VerifyErrorMessageFor5V
		 * Function Details: To verify power supply value
		 * Parameter/Arguments: sPowerSupply
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 1/04/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyErrorMessageFor5V(string sPowerSupply, bool sWarningSign)
		{
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Enter the Device text in Search Properties fields to view device related text
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("Power Supply" +"{ENTER}");
			
			// Click on Power Supply cell
			repo.FormMe.cell_PowerSupply.Click();
			
			
			if(sWarningSign)
			{
				if(repo.FormMe.Error_Symbol_PowerSupplyInfo.Exists())
				{
					Report.Log(ReportLevel.Success,"Error message is displayed ");
				}
				else
				{
					Report.Log(ReportLevel.Failure,"Error message is not displayed ");
				}
			}
			else
			{
				if(repo.FormMe.Error_Symbol_PowerSupplyInfo.Exists())
				{
					Report.Log(ReportLevel.Failure,"Error message is displayed ");
				}
				else
				{
					Report.Log(ReportLevel.Success,"Error message is not displayed ");
				}
			}
			
			
		}
		
		/********************************************************************
		 * Function Name: ChangeLabelName
		 * Function Details: To change Label Name
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void ChangeLabelName(string changeLabelName)
		{
			//Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Enter the Device text in Search Properties fields to view device related text
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("Label" +"{ENTER}" );
			
			// Click on Label Name cell
			repo.ProfileConsys1.cell_LabelName.Click();
			
			// Enter the value to change Label Name
			repo.ProfileConsys1.PARTItemsPresenter.txt_changeDeviceSensitivity.PressKeys("{LControlKey down}{Akey}{Delete}{LControlKey up}"+(changeLabelName)+"{ENTER}" + "{ENTER}");
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
			
			
			
		}
		
		
		/***********************************************************************************************************
		 * Function Name: verifyPointGridProperties
		 * Function Details: To verify points grid properties for a device
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Poonam Kadam
		 * Last Update : 03/04/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyPointGridProperties(string ExpectedPointGridRowCount, string ExpectedPointGridColumn, string ExpectedDeviceProperty)
		{
			// Retrieve Point grid rows and column count
			sColumn = getColumnNumberForPointsGrid(ExpectedPointGridColumn);
			sRow = ExpectedPointGridRowCount;
			
			Report.Log(ReportLevel.Success,"Row an column values are set as"+sColumn+sRow);
			
			string ActualPointGridProperty = repo.FormMe.txt_PointGridDeviceProperty.TextValue;
			// Compare Actual and Expected Point Grid Row count
			Report.Log(ReportLevel.Success,"Actual"+ActualPointGridProperty+"  Expected"+ExpectedDeviceProperty);
			
			if(ActualPointGridProperty.Equals(ExpectedDeviceProperty))
			{
				Report.Log(ReportLevel.Success,"Device property is verified in Points grid");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Device property is not verified not in Points grid");
			}
			
		}
		
		/***********************************************************************************************************
		 * Function Name: verifyBlankDeviceAddress
		 * Function Details: To verify points grid for blank device address
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Poonam Kadam
		 * Last Update : 03/04/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyBlankDeviceAddress(string ExpectedPointGridRowCount, string ExpectedPointGridColumn)
		{
			// Retrieve Point grid rows and column count
			sColumn = getColumnNumberForPointsGrid(ExpectedPointGridColumn);
			sRow = ExpectedPointGridRowCount;
			
			Report.Log(ReportLevel.Success,"Row an column values are set as"+sColumn+sRow);
			
			string ActualPointGridProperty = repo.FormMe.txt_PointGridDeviceProperty.TextValue;
			// Compare Actual and Expected Point Grid Row count
			Report.Log(ReportLevel.Success,"Actual"+ActualPointGridProperty);
			
			if(ActualPointGridProperty==null)
			{
				Report.Log(ReportLevel.Success,"Device address is verified in Points grid as blank");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Device address is not blank");
			}
			
		}
		
		
		/********************************************************************
		 * Function Name: getColumnNumberForPointsGrid
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Poonam Kadam
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static string getColumnNumberForPointsGrid(string columnName)
		{
			string columnNumber="";
			switch (columnName)
			{
				case "SKU":
					columnNumber="1";
					Report.Log(ReportLevel.Success,"Column number is set as"+columnNumber);
					break;
					
				case "Model":
					columnNumber="2";
					Report.Log(ReportLevel.Success,"Column number is set as"+columnNumber);
					break;
					
				case "Label":
					columnNumber="3";
					Report.Log(ReportLevel.Success,"Column number is set as"+columnNumber);
					break;
					
				case "Address":
					columnNumber="4";
					Report.Log(ReportLevel.Success,"Column number is set as"+columnNumber);
					break;
					
			}
			return columnNumber;
			
		}
		
		/***********************************************************************************************************
		 * Function Name: verifyDeviceProperties
		 * Function Details: To verify properties from properties section
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Poonam Kadam
		 * Last Update : 05/04/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyDeviceProperties(string sPropertyLabel, string sPropertyValue)
		{
			//Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Search for the expcted property
			repo.ProfileConsys1.txt_SearchProperties.PressKeys(sPropertyLabel +"{ENTER}" );
			
			// Click on property cell
			repo.ProfileConsys1.PARTItemsPresenter.cell_DeviceMode.Click();
			
			// Retrieve value for the property
			string actualPropertyValue = repo.ProfileConsys1.PARTItemsPresenter.txt_DayMatchesNight.TextValue;
			
			// Comparing actual and expected property value
			if(actualPropertyValue.Equals(sPropertyValue))
			{
				Report.Log(ReportLevel.Success,"Property value " +sPropertyValue+ " is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Property value is displayed as "+actualPropertyValue+ " instead of "+ sPropertyValue);
			}
			//Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
		}
		
		/***********************************************************************************************************
		 * Function Name: editDeviceLabel
		 * Function Details: To edit device label from properties section
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Poonam Kadam
		 * Last Update : 05/04/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void editDeviceLabel(string sPropertyLabel, string sNewLabel)
		{
			if(repo.ProfileConsys1.tab_PointsInfo.Exists())
			{
				//Click on Points tab
				repo.ProfileConsys1.tab_Points.Click();
			}
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Search for the Label property
			repo.ProfileConsys1.txt_SearchProperties.PressKeys(sPropertyLabel +"{ENTER}" );
			
			if(repo.ProfileConsys1.PARTItemsPresenter.cell_DeviceModeInfo.Exists())
			{
				// Click on Label property cell
				repo.ProfileConsys1.PARTItemsPresenter.cell_DeviceMode.Click();
				
				//Modifying the label
				repo.ProfileConsys1.PARTItemsPresenter.cell_DeviceMode.PressKeys("{LControlKey down}{Akey}{Delete}{LControlKey up}");
				repo.ProfileConsys1.PARTItemsPresenter.cell_DeviceMode.PressKeys(sNewLabel +"{ENTER}" );
				Report.Log(ReportLevel.Success,"Label is editied to " +sNewLabel);
				
			}
			else
			{
				// Click on label cell
				repo.FormMe.cell_Properties.Click();
				
				//Modifying the label
				repo.FormMe.cell_Properties.PressKeys("{LControlKey down}{Akey}{Delete}{LControlKey up}");
				repo.FormMe.cell_Properties.PressKeys(sNewLabel +"{ENTER}" );
				Report.Log(ReportLevel.Success,"Label is edited to " +sNewLabel);
				
			}
			
			if(repo.ProfileConsys1.tab_PointsInfo.Exists())
			{
				//Click on Points tab
				repo.ProfileConsys1.tab_Points.Click();
			}
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
		}
		
		/********************************************************************
		 * Function Name: DeleteAccessoryFromPanelAccessoriesTab
		 * Function Details: To delete Accessory from panel acccessories tab
		 * Parameter/Arguments: sRowNumber, sSkuNumber
		 * Output:
		 * Function Owner: Poonam Kadam
		 * Last Update : 8/4/2019
		 ********************************************************************/
		
		[UserCodeMethod]
		public static void DeleteAccessoryFromPanelAccessoriesTab()
		{
			repo.FormMe.PanelAccessoriesLabel.Click();
			Report.Log(ReportLevel.Success, "Accessory is Selected");
			repo.ProfileConsys1.btn_Delete.Click();
			Report.Log(ReportLevel.Success, "Accessory is deleted");
		}
		
		/********************************************************************
		 * Function Name: PressControlButton
		 * Function Details: To press control button from keyboard
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Poonam kadam
		 * Last Update : 09/4/19
		 ********************************************************************/
		[UserCodeMethod]
		public static void PressControlButton()
		{
			Keyboard.Press("{LControlKey down}");
			Report.Log(ReportLevel.Success, "Control button pressed");
		}
		
		/********************************************************************
		 * Function Name: ReleaseControlButton
		 * Function Details: To release control button from keyboard
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Poonam kadam
		 * Last Update : 09/4/19
		 ********************************************************************/
		[UserCodeMethod]
		public static void ReleaseControlButton()
		{
			Keyboard.Press("{LControlKey up}");
			Report.Log(ReportLevel.Success, "Control button released");
		}
		
		/********************************************************************
		 * Function Name: SelectMultipleDevices
		 * Function Details: To Select multiple devices
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Poonam kadam
		 * Last Update : 10/4/19
		 ********************************************************************/
		[UserCodeMethod]
		public static void SelectMultipleDevices(string strRow)
		{
			PressControlButton();
			SelectPointsGridRow(strRow);
			Keyboard.Press("{LControlKey up}");
			Keyboard.Press("{LControlKey up}");
			Report.Log(ReportLevel.Success, "Control button released");
		}
		
		/********************************************************************
		 * Function Name: VerifyDeviceDisplayedInPhysicalLayout
		 * Function Details: Verify if device is displayed in physical layout at expected position
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Poonam kadam
		 * Last Update : 11/4/19
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyDeviceDisplayedInPhysicalLayout(string DeviceIndex, string ExpectedDeviceAddress, string ExpectedDeviceName)
		{
			sPhysicalLayoutDeviceIndex = DeviceIndex;
			sDeviceAddress = ExpectedDeviceAddress;
			string ActualDeviceName = repo.FormMe.lst_PhysicalLayoutDevice.TextValue;
			string ActualDeviceAddress=repo.FormMe.txt_PhysicalLayoutDeviceAddress.TextValue;
			Report.Log(ReportLevel.Success,"Expected"+ActualDeviceName + ActualDeviceAddress);
			// Compare actualIndex and sPhysicalLayoutDeviceIndex values and then displaying result
			if(ActualDeviceName.Equals(ExpectedDeviceName)&&(ActualDeviceAddress.Equals(ExpectedDeviceAddress)))
			{
				Report.Log(ReportLevel.Success, "Device " +ExpectedDeviceName+ " added successfully and displaying correctly in Physical Layout");
			}
			else
			{
				Report.Log(ReportLevel.Failure, "Device " +ExpectedDeviceName+ " not added or not displaying correctly in Physical Layout");
			}
		}
		
		
		/********************************************************************
		 * Function Name: ChangeCableCapacitance
		 * Function Details: To change cable capacitance
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 19/04/2019
		 ********************************************************************/
		// Change cable capacitance method
		[UserCodeMethod]
		public static void ChangeCableCapacitance(int fchangeCableCapacitance, string sLabelName)
		{
			//Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			//Click on Panel Node
			repo.ProfileConsys1.PanelNode.Click();
			
			//Click on Loop A in Navigation tree tab
			repo.ProfileConsys1.NavigationTree.Loop_A.Click();
			
			SelectRowUsingLabelName(sLabelName);
			
			//Click on cable capacitance cell
			repo.ProfileConsys1.cell_CableCapacitance.Click();
			
			//Change the value of cable length
			Keyboard.Press("{LControlKey down}{Akey}{LControlKey up}"+fchangeCableCapacitance + "{Enter}");
			
			//Click on Panel Node
			repo.ProfileConsys1.PanelNode.Click();
			Delay.Duration(1000, false);
		}
		
		/********************************************************************
		 * Function Name: VerifyCableLength
		 * Function Details: To verify cable length
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 24/04/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyCableLength(string sCableLength)
		{
			//Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			//Click on Panel Node
			repo.ProfileConsys1.PanelNode.Click();
			
			//Click on Loop A in Navigation tree tab
			repo.ProfileConsys1.NavigationTree.Loop_A.Click();
			
			//Click on cable length cell
			repo.ProfileConsys1.cell_CableLength.Click();
			
			string actualCableLength = repo.ProfileConsys1.txt_CableLength.TextValue;
			
			
			// Comparing actual and expected value
			if(actualCableLength.Equals(sCableLength))
			{
				Report.Log(ReportLevel.Success,"Cable length value " +sCableLength+ " is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Cable length value is displayed as "+actualCableLength+ " instead of "+ sCableLength);
			}


			
		}
		
		/********************************************************************
		 * Function Name: VerifyCableLengthInNodeGalleryItems
		 * Function Details: To verify cable length
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 24/04/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyCableLengthInNodeGalleryItems(string sCableLength)
		{
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Enter the Day Matches night text in Search Properties fields to view cable length;
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("{LControlKey down}{Akey}{LControlKey up}Length" +"{ENTER}" );
			
			repo.FormMe.cell_CableLength.Click();
			
			string actualCableLength = repo.FormMe.txt_InventoryProperty.TextValue;
			
			// Comparing actual and expected value
			if(actualCableLength.Equals(sCableLength))
			{
				Report.Log(ReportLevel.Success,"Cable length value " +sCableLength+ " is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Cable length value is displayed as "+actualCableLength+ " instead of "+ sCableLength);
			}

		}
		
		
		/********************************************************************
		 * Function Name: verifyMinMaxThroughSpinControlForCableLength
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 24/04/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void verifyMinMaxThroughSpinControlForCableLength(string minLimit,string maxLimit)
		{
			// Click on Cable Length cell
			repo.FormMe.cell_CableLength.Click();
			
			Keyboard.Press("{LControlKey down}{Akey}{LControlKey up}"+ maxLimit +"{ENTER}");
			
			// Click on Cable Length cell
			repo.FormMe.cell_CableLength.Click();
			
			repo.FormMe.cableLengthSpinUpButton.Click();
			
			string actualCableLengthValue = repo.FormMe.txt_InventoryProperty.TextValue;
			
			if(actualCableLengthValue.Equals(maxLimit))
			{
				Report.Log(ReportLevel.Success,"Spin control accepts values within specified max limit");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Spin control does not accepts values within specified max limit");
			}
			Keyboard.Press("{ENTER}");
			
			Keyboard.Press("{LControlKey down}{Akey}{LControlKey up}"+ minLimit +"{ENTER}");
			
			// Click on Cable Length cell
			repo.FormMe.cell_CableLength.Click();
			
			repo.FormMe.cableLengthSpinDownButton.Click();
			
			actualCableLengthValue =repo.FormMe.txt_InventoryProperty.TextValue;
			
			if(actualCableLengthValue.Equals(minLimit))
			{
				Report.Log(ReportLevel.Success,"Spin control accepts values within specified min limit");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Spin control does not accepts values within specified min limit");
			}
		}
		
		/********************************************************************
		 * Function Name: SelectNodeInventoryLabel
		 * Function Details: To select points grid row
		 * Parameter/Arguments: sRowNumber
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 25/04/2018
		 ********************************************************************/
		[UserCodeMethod]
		public static void SelectNodeInventoryLabel(string sRowNumber)
		{
			sRow=sRowNumber;
			//Click on row from points grid
			repo.FormMe.NodeInventoryLabel.Click();
		}
		
		/***********************************************************************************************************
		 * Function Name: verifyShoppingListDevicesTextForPxD
		 * Function Details: To verify shopping list devices via clicking on its row
		 * Parameter/Arguments: sFileName,sDeviceSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :  26/04/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyShoppingListDevicesTextForPxD(string sExpectedText)
		{
			
			string actualText = repo.ShoppingListCompatibilityModeE.Cell18.Text;
			
			if(actualText.Equals(sExpectedText))
			{
				Report.Log(ReportLevel.Success,"Model name " +actualText+ " is displayed successfully");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Model name" +sExpectedText+ " is not displayed correctly instead " +actualText+  " is displayed " );
			}
		}
		
		
		/***********************************************************************************************************
		 * Function Name: verifyShoppingListDevicesTextForPSC
		 * Function Details: To verify shopping list devices via clicking on its row
		 * Parameter/Arguments: sFileName,sDeviceSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 26/04/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyShoppingListDevicesTextForPSC(string sExpectedText)
		{
			
			string actualText = repo.ShoppingListCompatibilityModeE.CellF22.Text;
			
			if(actualText.Equals(sExpectedText))
			{
				Report.Log(ReportLevel.Success,"Model name " +actualText+ " is displayed successfully");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Model name" +sExpectedText+ " is not displayed correctly instead " +actualText+  "is displayed " );
			}
		}
		
		
		/***********************************************************************************************************
		 * Function Name: VerifyDeleteButton
		 * Function Details: VerifyDeleteButton state
		 * Parameter/Arguments: IsReadOnly
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 30/04/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyDeleteButton(bool isReadOnly)
		{
			
			bool result = repo.ProfileConsys1.btn_Delete.Enabled;
			if(result==isReadOnly)
			{
				Report.Log(ReportLevel.Success,"Delete button displaying is as expected");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Delete button displaying is not as expected");
			}
			
		}
		
		
		/*****************************************************************************************************************
		 * Function Name:  VerifyCurrentDCUnitscalculation
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Purvi Bhasim
		 * Last Update : 08/01/2019
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyCurrentDCUnitscalculation(string sFileName,string sAddPanelSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared variables
			string ModelNumber,sType,sLabelName,sAssignedBase,expectedDCUnits,DefaultDCUnits,ChangedDCUnit,sPanelLEDCount;
			int PanelLED;
			
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
				
				int.TryParse(sPanelLEDCount,out PanelLED);
				
				// Click on Expander node
				repo.ProfileConsys1.NavigationTree.Expander.Click();
				
				// Click on Loop Card node
				repo.ProfileConsys1.NavigationTree.Expand_LoopCard.Click();
				
				// Click on Loop A node
				repo.ProfileConsys1.NavigationTree.Loop_A.Click();
				
				Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
				
				//Assign Base to the device
				Devices_Functions.AssignDeviceBase(sLabelName,sAssignedBase,sRowIndex);
				

				//Verify Default DC Units
				verifyDCUnitsValue(expectedDCUnits);
				
				repo.ProfileConsys1.SiteNode.Click();
				
			}
			//Go to Loop A
			repo.ProfileConsys1.NavigationTree.Loop_A.Click();
			
			//go to points grid
			repo.ProfileConsys1.tab_Points.Click();
			
			Keyboard.Press("{LControlKey down}{Akey}{LControlKey up}");
			
			//Copy Devices
			repo.FormMe.btn_Copy.Click();
			
			//Go to Loop C
			repo.ProfileConsys1.NavigationTree.Loop_C.Click();
			
			//Paste the devices
			repo.FormMe.Paste.Click();
			
			//Verify DC Units
			verifyDCUnitsValue(expectedDCUnits);
			
			repo.ProfileConsys1.SiteNode.Click();
			
			//Go to Loop C
			repo.ProfileConsys1.NavigationTree.Loop_C.Click();
			
			//go to points grid
			repo.ProfileConsys1.tab_Points.Click();
			
			Keyboard.Press("{LControlKey down}{Akey}{LControlKey up}");
			
			//Copy Devices
			repo.FormMe.ButtonCut.Click();
			
			//Verify Default DC Units
			verifyDCUnitsValue(DefaultDCUnits);
			
			repo.ProfileConsys1.SiteNode.Click();
			
			// Click on Expander node
			repo.ProfileConsys1.NavigationTree.Expander.Click();
			
			Panel_Functions.changePanelLED(PanelLED);
			
			// Click on Loop Card node
			repo.ProfileConsys1.NavigationTree.Expand_LoopCard.Click();
			
			// Click on Loop A node
			repo.ProfileConsys1.NavigationTree.Loop_A.Click();
			
			//Verify Default DC Units
			verifyDCUnitsValue(ChangedDCUnit);

		}
		
		
		/********************************************************************
		 * Function Name: AddMaxNumberOfPanelAccessories(
		 * Function Details: Add devices from panel accessories
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Date: 26/4/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void AddAndVerifyMaxNumberOfPanelAccessories(string sFileName, string sDeviceSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sDeviceSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;

			// Declared string type
			string ModelNumber,sType,Maxnumber;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				ModelNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				Maxnumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				
				int MaxNumber = Convert.ToInt32(Maxnumber);
				
				
				for(int j=1; j<=MaxNumber; j++)
				{
					AddDevicefromPanelAccessoriesGallery(ModelNumber,sType);
				}
				
				bool EnabledStatus = false;
				//Verify gallery disabled
				VerifyDeviceIsDisabledOrEnabled(ModelNumber,sType,EnabledStatus);
				
				//Verify slot cards
				if(MaxNumber>6)
				{
					int remainingSlots = MaxNumber-6;
					string sremainingSlots = remainingSlots.ToString();
					string expectedText = "Other Slot Cards ("+MaxNumber+" of 6)";
					
					//Verify Other slot cards 1
					string actualSlotText = repo.FormMe.OtherSlotCards_Text.TextValue;
					if(actualSlotText.Equals(expectedText))
					{
						Report.Log(ReportLevel.Success,"Other slot cards are dispayed correctly ");
					}
					else
					{
						Report.Log(ReportLevel.Success,"Other slot cards are not dispayed correctly ");
					}
					
					string expectedText2 = "Other Slot Cards ("+remainingSlots+" of 6)";
					repo.FormMe.Backplane2_Expander.Click();
					
					//Verify Other slot cards 2
					string actualSlotText2 = repo.FormMe.OtherSlotCards2_Text.TextValue;
					
					if(actualSlotText2.Equals(expectedText2))
					{
						Report.Log(ReportLevel.Success,"Other slot cards are dispayed correctly ");
					}
					else
					{
						Report.Log(ReportLevel.Success,"Other slot cards are not dispayed correctly ");
					}
				}
				else
				{
					//Verify other slot cards 1
					string expectedText = "Other Slot Cards ("+MaxNumber+" of 6)";
				}
				
				//Delete slot cards from Panel Accessories
				for(int j=1; j<=MaxNumber; j++)
				{
					string labelName= ModelNumber+"-"+j;
					
					//Delete device using label name
					DeleteAccessoryFromPanelAccessoriesTab();
					
					
					if(j==1)
					{
						//Verify Gallery
						EnabledStatus = true;
						//Verify gallery disabled
						VerifyDeviceIsDisabledOrEnabled(ModelNumber,sType,EnabledStatus);
					}
				}
				
				
				
				
				
			}
			
		}
		
		
		/********************************************************************
		 * Function Name: VerifyCustomDeviceDisplayedInCustomGallery
		 * Function Details: Verify if custom device is displayed in custom Gallery
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 16/4/19
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifySounderCustomDeviceDisplayedInCustomGallery(string DeviceName, bool isEnabled)
		{
			//DeleteDevicesPresentInCustomGallery();
			if(isEnabled)
			{
				if(repo.FormMe.Custom_Item_In_Gallery_For_Sounders.Enabled)
				{
					repo.FormMe.Custom_Item_In_Gallery_For_Sounders.Click();
					Report.Log(ReportLevel.Success, "Device " +DeviceName+ " is enabled in gallery");
				}
				else
				{
					Report.Log(ReportLevel.Failure, "Device " +DeviceName+ " is disabled in gallery");
				}
				
			}
			else
			{
				if(repo.FormMe.Custom_Item_In_Gallery_For_Others.Enabled)
				{
					Report.Log(ReportLevel.Failure, "Device " +DeviceName+ " is enabled in gallery");
				}
				else
				{
					Report.Log(ReportLevel.Success, "Device " +DeviceName+ " is disabled in gallery");
				}
			}
		}
		
		/********************************************************************
		 * Function Name: VerifyDeviceIsDisabledOrEnabled
		 * Function Details:Verify device is enabled in panel accessories
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 26/04/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyDeviceIsDisabledOrEnabled(string sDeviceName,string sType,bool sEnabled)
		{
			sAccessoriesGalleryIndex= SelectPanelAccessoriesGalleryType(sType);
			ModelNumber=sDeviceName;
			sDeviceIndex = SelectDeviceFromPanelAccessories(ModelNumber);
			repo.FormMe.btn_PanelAccessoriesDropDown.Click();
			

			
			if(sEnabled)
			{
				if(repo.ContextMenu.PanelAccessories_Device.Enabled)
				{
					Report.Log(ReportLevel.Success,"Panel Accessories device " + sDeviceName + " is enabled" );
				}
				else
				{
					Report.Log(ReportLevel.Failure,"Panel Accessories device " + sDeviceName + " is not enabled" );
				}
			}
			else
			{
				if(repo.ContextMenu.PanelAccessories_Device.Enabled)
				{
					Report.Log(ReportLevel.Failure,"Panel Accessories device " + sDeviceName + " is not disabled even after reaching max limit " );
				}
				else
				{
					Report.Log(ReportLevel.Success,"Panel Accessories device " + sDeviceName + " is disabled even after reaching max limit " );
				}
			}
			
			
			
		}

		
		
		/********************************************************************
		 * Function Name: VerifyPanelNodePanelAccessoriesGallery
		 * Function Details: To verify day mode field
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 06/05/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyPanelNodePanelAccessoriesGallery(string sDeviceName,string sType,string state)
		{
			if(state.Equals("Enabled"))
			{
				sAccessoriesGalleryIndex= SelectPanelAccessoriesGalleryType(sType);
				ModelNumber=sDeviceName;
				repo.FormMe.btn_PanelAccessoriesDropDown.Click();
				
				if (repo.ContextMenu.txt_SelectDevice.Enabled)
				{
					Report.Log(ReportLevel.Success, "Gallery Item: " + sDeviceName+ " Enabled in gallery");
				}
				else
				{
					Report.Log(ReportLevel.Failure, "Gallery Item: " + sDeviceName+ " Disabled in gallery");
				}
			}
			else
			{
				sAccessoriesGalleryIndex= SelectPanelAccessoriesGalleryType(sType);
				ModelNumber=sDeviceName;
				repo.FormMe.btn_PanelAccessoriesDropDown.Click();
				
				if (repo.ContextMenu.txt_SelectDevice.Enabled)
				{
					Report.Log(ReportLevel.Failure, "Gallery Item: " + sDeviceName+ " enabled in gallery");
				}
				else
				{
					Report.Log(ReportLevel.Success, "Gallery Item: " + sDeviceName+ " disabled in gallery");
				}
			}
			
		}
		
		/********************************************************************
		 * Function Name: SelectPanelNodePanelAccessoriesRow
		 * Function Details: To select item from grid using label
		 * Parameter/Arguments: sLabelName
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 06/05/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void SelectPanelNodePanelAccessoriesRow(string RowNumber)
		{
			sRow = RowNumber;
			repo.FormMe.SelectPanelAccessoriesLabel.Click();
			Report.Log(ReportLevel.Success, "Device with Label name selected");
		}
		

		/********************************************************************
		 * Function Name: AddDevicesfromPanelNodeGallery
		 * Function Details: Add devices from panel node gallery
		 * Parameter/Arguments: devicename , device type and panel type
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 10/05/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void AddDevicesfromPanelNodeGallery(string sDeviceName,string sType, string PanelType)
		{
			sMainProcessorGalleryIndex = SelectMainProcessorGalleryType(sType, PanelType);
			ModelNumber=sDeviceName;
			repo.FormMe.btn_PanelNodelGalleryDropDown.Click();
			repo.ContextMenu.txt_SelectDevice.Click();
			Report.Log(ReportLevel.Info, "Device "+sDeviceName+" added successfully");
		}
		
		/********************************************************************
		 * Function Name: VerifyNodeGallery
		 * Function Details: To verify node gallery
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 14/05/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyNodeGallery(string sDeviceName,string sType,string state,string PanelType)
		{
			if(state.Equals("Enabled"))
			{
				sAccessoriesGalleryIndex= SelectMainProcessorGalleryType(sType,PanelType);
				ModelNumber=sDeviceName;
				repo.FormMe.btn_PanelNodelGalleryDropDown.Click();
				
				
				if (repo.ContextMenu.txt_SelectDevice.Enabled)
				{
					Report.Log(ReportLevel.Success, "Gallery Item: " + sDeviceName+ " Enabled in gallery");
				}
				else
				{
					Report.Log(ReportLevel.Failure, "Gallery Item: " + sDeviceName+ " Disabled in gallery");
				}
			}
			else
			{
				sAccessoriesGalleryIndex= SelectMainProcessorGalleryType(sType,PanelType);
				ModelNumber=sDeviceName;
				repo.FormMe.btn_PanelNodelGalleryDropDown.Click();
				
				if (repo.ContextMenu.txt_SelectDevice.Enabled)
				{
					Report.Log(ReportLevel.Failure, "Gallery Item: " + sDeviceName+ " enabled in gallery");
				}
				else
				{
					Report.Log(ReportLevel.Success, "Gallery Item: " + sDeviceName+ " disabled in gallery");
				}
			}
			
		}
		/********************************************************************
		 * Function Name: RightClickOnSelectedGridRow
		 * Function Details: To open the context menu options using right click in grid
		 * Parameter/Arguments: RowNumber
		 * Output:
		 * Function Owner: Poonam
		 * Last Update : 15/5/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void RightClickOnSelectedInventoryGridRow(string RowNumber)
		{
			sRow = RowNumber;
			repo.FormMe.InventoryGridRow.Click(System.Windows.Forms.MouseButtons.Right);
			
		}
		
		/********************************************************************
		 * Function Name: VerifyPanelType
		 * Function Details: To Verify panel type
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 15/05/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyPanelType(string sFileName,string sAddDevicesSheet, string sPanelName)
		{
			//Click on Panel Node
			repo.ProfileConsys1.PanelNodeText.Click();
			
			repo.ProfileConsys1.SiteNode.Click();
			
			repo.ProfileConsys1.PanelNodeText.Click();
			
			repo.ProfileConsys1.SiteNode.Click();
			
			repo.ProfileConsys1.PanelNodeText.Click();
			
			repo.FormMe.tab_PanelAccessories.Click();
			
			repo.ProfileConsys1.PanelNodeText.Click();
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Enter the Panel type ext in Search Properties fields to view Panel type text;
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("Panel Type" +"{ENTER}" );
			
			// Click on Panel type cell
			repo.ProfileConsys1.cell_NumberOfAlarmLeds.Click();
			
			//Retrieve value of Panel type text and store in PanelType
			string actualPanelName = repo.FormMe.PanelType.TextValue;

			if(actualPanelName.Equals(sPanelName))
			{
				Report.Log(ReportLevel.Success, "Panel name " +actualPanelName+ " is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure, "Panel name is not displayed correctly");
				
			}
			
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			int columns = Excel_Utilities.ExcelRange.Columns.Count;

			
			string PanelTypeName = ((Range)Excel_Utilities.ExcelRange.Cells[10,17]).Value.ToString();
			
			
			VerifyPanelTypeNames(PanelTypeName);
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");

		}
		/********************************************************************
		 * Function Name: clickContextMenuOptionOnRightClick
		 * Function Details: To verify if paste button is enabled
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Poonam
		 * Last Update : 20/5/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void clickContextMenuOptionOnRightClick(string sContectMenuOption)
		{
			sListIndex=sContectMenuOption;
			repo.ContextMenu.ContextMenuOption.Click();
			Report.Log(ReportLevel.Success, sContectMenuOption+" button is clicked");
		}
		
		/********************************************************************
		 * Function Name:verifyContextMenuOptionOnRightClickEnabledOrDisabled
		 * Function Details: To verify if context menu option is enabled or disabled when we right click on grid row
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Poonam
		 * Last Update : 20/5/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void verifyContextMenuOptionOnRightClickEnabledOrDisabled(string sContectMenuOption)
		{
			sListIndex=sContectMenuOption;
			if(repo.ContextMenu.ContextMenuOption.Enabled)
			{
				Report.Log(ReportLevel.Success, sContectMenuOption+" button is enabled");
			}
			else{
				Report.Log(ReportLevel.Success, sContectMenuOption+" button is disabled");
			}
		}
		/********************************************************************
		 * Function Name: verifyPasteButtonEnabled
		 * Function Details: To verify if paste button is enabled in ribbon
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Poonam
		 * Last Update : 20/5/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void verifyPasteButtonEnabled()
		{
			if (repo.FormMe.Paste.Enabled)
			{
				Report.Log(ReportLevel.Success, "Paste button is enabled");
			}
			else
			{
				Report.Log(ReportLevel.Failure, "Paste button is disabled");
			}
		}
		
		/********************************************************************
		 * Function Name: VerifyPanelTypeNames
		 * Function Details: To Verify panel type
		 * Parameter/Arguments:PanelTypeNameList
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 16/05/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyPanelTypeNames(string PanelTypeNameList)
		{
			// Split Paneltype name and then add panels in the collection list
			List<string>  splitPanelTypesNames  = PanelTypeNameList.Split(',').ToList();
			
			foreach(string item in splitPanelTypesNames)
			{
				bool found=false;
				foreach(ListItem listitem in repo.ContextMenu.PanelTypeList.Items)
				{
					if(item == listitem.Text)
					{
						found = true;
						Report.Log(ReportLevel.Success, "Panel name " +listitem.Text+ " is displayed correctly in panel type dropdown list");
						break;
					}
					
				}
				
				if(found == false)
				{
					Report.Log(ReportLevel.Info,"Panel " +item+ " not  found in the list");
					

				}
				
				
			}
			
			
			
		}
		
		
		/********************************************************************
		 * Function Name: verifyPasteButtonDisabled
		 * Function Details: To verify if paste button is disabled in ribbon
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Poonam
		 * Last Update : 20/5/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void verifyPasteButtonDisabled()
		{
			if (repo.FormMe.Paste.Enabled)
			{
				Report.Log(ReportLevel.Failure, "Paste button is enabled");
			}
			else
			{
				Report.Log(ReportLevel.Success, "Paste button is disabled");
			}
		}

		
		/********************************************************************
		 * Function Name: VerifyPanelType
		 * Function Details: To Verify panel type
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 15/05/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyPanelTypeInDropdown(string PanelName, string PanelTypeNameList,string PanelTypeNameListNotAvailable)
		{
			//Click on Panel Node
			repo.ProfileConsys1.PanelNodeText.Click();
			
			repo.ProfileConsys1.SiteNode.Click();
			
			repo.ProfileConsys1.PanelNodeText.Click();
			
			repo.ProfileConsys1.SiteNode.Click();
			
			repo.ProfileConsys1.PanelNodeText.Click();
			
			repo.FormMe.tab_PanelAccessories.Click();
			
			repo.ProfileConsys1.PanelNodeText.Click();
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Enter the Panel type ext in Search Properties fields to view Panel type text;
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("Panel Type" +"{ENTER}" );
			
			// Click on Panel type cell
			repo.ProfileConsys1.cell_NumberOfAlarmLeds.Click();
			
			//Retrieve value of Panel type text and store in PanelType
			string actualPanelName = repo.FormMe.PanelType.TextValue;

			if(actualPanelName.Equals(PanelName))
			{
				Report.Log(ReportLevel.Success, "Panel name " +actualPanelName+ " is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure, "Panel name is not displayed correctly");
				
			}
			
			VerifyPanelTypeNames(PanelTypeNameList);
			VerifyPanelTypeNamesNotAvailable(PanelTypeNameListNotAvailable);
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");

		}
		
		/********************************************************************
		 * Function Name: VerifyPanelTypeNames
		 * Function Details: To Verify panel type
		 * Parameter/Arguments:PanelTypeNameList
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 16/05/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyPanelTypeNamesNotAvailable(string PanelTypeNameListNotAvailable)
		{
			// Split Paneltype name and then add panels in the collection list
			List<string>  splitPanelTypesNamesNotAvailable  = PanelTypeNameListNotAvailable.Split(',').ToList();
			
			foreach(string item in splitPanelTypesNamesNotAvailable)
			{
				bool found=true;
				foreach(ListItem listitem in repo.ContextMenu.PanelTypeList.Items)
				{
					if(item == listitem.Text)
					{
						found = false;
						Report.Log(ReportLevel.Failure, "Panel name " +listitem.Text+ " is displayed incorrectly in panel type dropdown list");
						break;
					}
					
				}
				
				if(found == true)
				{
					Report.Log(ReportLevel.Success, "Panel " +item+ " not  found in the list as expected");
					

				}
				
				
			}
			
			
			
		}
		
		/********************************************************************
		 * Function Name: VerifyEnableDisablePanelAccessoriesGallery
		 * Function Details:
		 * Parameter/Arguments: sType,deviceName,state
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 23/05/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyEnableDisablePanelAccessoriesGallery(string sType,string deviceName, string state)
		{
			if(state.Equals("Enabled"))
			{
				sAccessoriesGalleryIndex= SelectPanelAccessoriesGalleryType(sType);
				ModelNumber=deviceName;
				repo.FormMe.btn_PanelAccessoriesDropDown.Click();
				if (repo.ContextMenu.txt_SelectDevice.Enabled)
				{
					Report.Log(ReportLevel.Success, "Accessories : " + deviceName+ " Enabled in gallery");
				}
				else
				{
					Report.Log(ReportLevel.Failure, "Accessories : " + deviceName+ " Disabled in gallery");
				}
			}
			else
			{
				sAccessoriesGalleryIndex= SelectPanelAccessoriesGalleryType(sType);
				ModelNumber=deviceName;
				repo.FormMe.btn_PanelAccessoriesDropDown.Click();
				if (repo.ContextMenu.txt_SelectDevice.Enabled)
				{
					Report.Log(ReportLevel.Failure, "Accessories : " + deviceName+ " enabled in gallery");
				}
				else
				{
					Report.Log(ReportLevel.Success, "Accessories: " + deviceName+ " disabled in gallery");
				}
			}
			
		}


		
		/********************************************************************
		 * Function Name: VerifyLabelInSearchProperties
		 * Function Details: To verify label in search properties
		 * Parameter/Arguments:expected Label text
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 24/05/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyLabelInSearchProperties(string sLabel)
		{
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Search Label properties
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("Label" +"{ENTER}" );
			
			// Click on label cell
			repo.FormMe.cell_Properties.Click();
			
			// Retrieve value of label
			string actualLabel = repo.FormMe.txt_PropertiesTextValue.TextValue;
			
			// Comparing actualLabel and sLabel values
			if(actualLabel.Equals(sLabel))
			{
				Report.Log(ReportLevel.Success,"Label text " +sLabel+ " is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Label text is not displayed as "+actualLabel+ " instead of "+ sLabel);
			}
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
		}
		
		/********************************************************************
		 * Function Name: VerifySKUInSearchProperties
		 * Function Details: To verify SKU in search properties
		 * Parameter/Arguments:expected Label text
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 24/05/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifySKUInSearchProperties(string sSKU)
		{
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Search Label properties
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("SKU" +"{ENTER}" );
			
			// Click on label cell
			repo.FormMe.cell_Properties.Click();
			
			// Retrieve value of label
			string actualSKUValue = repo.FormMe.txt_PropertiesTextValue.TextValue;
			
			// Comparing actualLabel and sLabel values
			if(actualSKUValue.Equals(sSKU))
			{
				Report.Log(ReportLevel.Success,"Actual SKU value " +actualSKUValue+ " is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Actual SKU Value is displayed as "+actualSKUValue+ " instead of "+ sSKU);
			}
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
		}
		
		/********************************************************************
		 * Function Name: VerifyModelInSearchProperties
		 * Function Details: To verify Model in search properties
		 * Parameter/Arguments:expected Model text
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 24/05/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyModelInSearchProperties(string sModel)
		{
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Search Label properties
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("Model" +"{ENTER}" );
			
			// Click on label cell
			repo.FormMe.cell_Properties.Click();
			
			// Retrieve value of label
			string actualModelText = repo.FormMe.txt_PropertiesTextValue.TextValue;
			
			// Comparing actualModelText and sModel values
			if(actualModelText.Equals(sModel))
			{
				Report.Log(ReportLevel.Success,"Model text " +actualModelText+ " is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Model text is displayed as "+actualModelText+ " instead of "+ sModel);
			}
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
		}
		
		
		/********************************************************************
		 * Function Name: VerifyFOMInSearchProperties
		 * Function Details: To verify FOM in search properties
		 * Parameter/Arguments:expected FOM text
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 24/05/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyFOMInSearchProperties(string sFOM)
		{
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Search Label properties
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("FOM" +"{ENTER}" );
			
			// Click on label cell
			repo.FormMe.cell_Properties.Click();
			
			// Retrieve value of label
			string actualFOM = repo.FormMe.txt_PropertiesTextValue.TextValue;
			
			// Comparing actualModelText and sModel values
			if(actualFOM.Equals(sFOM))
			{
				Report.Log(ReportLevel.Success,"FOM text " +actualFOM+ " is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"FOM text is displayed as "+actualFOM+ " instead of "+ sFOM);
			}
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
		}
		
		/********************************************************************
		 * Function Name: VerifyFOMInSearchProperties
		 * Function Details: To verify FOM in search properties
		 * Parameter/Arguments:expected FOM text
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 24/05/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyMPMInSearchProperties()
		{
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
			
			// Search Label properties
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("MPM" +"{ENTER}" );
			
			// Click on label cell
			repo.FormMe.cell_Properties.Click();
			
			
			// To retrieve the attribute value as boolean by its ischecked properties and store in actual state
			bool actualState =  repo.FormMe.chkbox_MPM800.GetAttributeValue<bool>("ischecked");
			
			//As per actual state and expected state values verfiying day mode and day sensitivity field state and action performed on checkbox
			if(actualState)
			{
				Report.Log(ReportLevel.Info,"MPM checkbox is available and is checked");
			}
			else
			{
				Report.Log(ReportLevel.Info,"MPM checkbox is available and is not checked");
			}

			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
		}
		
		
		/********************************************************************
		 * Function Name: VerifyDescriptionTextRowInSearchProperties
		 * Function Details: To verify Verify Description Text Row in Search Properties
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 25/05/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyDescriptionTextRowInSearchProperties()
		{
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Search Label properties
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("Description" +"{ENTER}" );
			
			// Click on
			repo.FormMe.txt_PropertiesTextRow.Click();
			
			if(repo.FormMe.txt_PropertiesTextRowInfo.Exists())
			{
				Report.Log(ReportLevel.Success,"Description text row available");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Description text row is not available");
			}
			
			
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
		}

		
		/********************************************************************
		 * Function Name: ChangeFOMInSearchProperties
		 * Function Details: To selected FOM in search properties
		 * Parameter/Arguments:expected FOM text
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 24/05/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void ChangeFOMInSearchProperties(string changeFOMValue)
		{
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
			
			// Search Label properties
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("FOM" +"{ENTER}" );
			
			// Click on label cell
			repo.FormMe.cell_Properties.Click();
			
			// Enter the value to change FOM value
			repo.FormMe.txt_PropertiesTextValue.PressKeys((changeFOMValue) +"{ENTER}" + "{ENTER}");
			
			// Click on label cell
			repo.FormMe.cell_Properties.Click();
			
			// Retrieve value of label
			string actualFOM = repo.FormMe.txt_PropertiesTextValue.TextValue;
			
			//Comparing expected and actual changed values for FOM
			if(actualFOM.Equals(changeFOMValue))
			{
				Report.Log(ReportLevel.Success,"FOM text changed successfully to  " +actualFOM+ " is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"FOM text is not changed to "+changeFOMValue+ " and displayed incorrectly");
			}
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
		}
		
		
		/********************************************************************
		 * Function Name: CheckUncheckMPMCheckboxInSearchProperties
		 * Function Details: To Check Uncheck MPMCheckbox In SearchProperties
		 * Parameter/Arguments:expected FOM text
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 27/05/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void CheckUncheckMPMCheckboxInSearchProperties(bool ExpectedState)
		{
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
			
			// Search Label properties
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("MPM" +"{ENTER}" );
			
			// Click on label cell
			repo.FormMe.cell_Properties.Click();

			// To retrieve the attribute value as boolean by its ischecked properties and store in actual state
			bool actualState =  repo.FormMe.chkbox_MPM800.GetAttributeValue<bool>("ischecked");
			
			//As per actual state and expected state values verfiying day mode and day sensitivity field state and action performed on checkbox
			if(actualState.Equals(ExpectedState))
			{
				Report.Log(ReportLevel.Success,"MPM checkbox is displayed as expected");
			}
			else
			{
				// Click on MPM checkbox
				repo.FormMe.chkbox_MPM800.Click();
				Report.Log(ReportLevel.Success,"Action performed on MPM checkbox");
			}
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
		}
		
		
		
		/***********************************************************************************************************
		 * Function Name: verifyShoppingListDevicesTextForThirdDevice
		 * Function Details: To verify shopping list devices via clicking on its row
		 * Parameter/Arguments: sFileName,sDeviceSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 27/05/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyShoppingListDevicesTextForThirdDevice(string sExpectedText)
		{
			
			string actualText = repo.ShoppingListCompatibilityModeE.CellF26.Text;
			
			if(actualText.Equals(sExpectedText))
			{
				Report.Log(ReportLevel.Success,"Model name " +actualText+ " is displayed successfully");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Model name" +sExpectedText+ " is not displayed correctly instead " +actualText+  "is displayed " );
			}
		}
		
		/***********************************************************************************************************
		 * Function Name: verifyShoppingListDevicesTextForCell3And14
		 * 		 * Function Details: To verify shopping list devices via clicking on its row
		 * Parameter/Arguments: sFileName,sDeviceSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 01/06/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyShoppingListDevicesTextForCell3And14(string sExpectedPanelText, string sExpectedDeviceText)
		{
			repo.ShoppingListCompatibilityModeE.CellF3.Click();
			
			string actualPanelText = repo.ShoppingListCompatibilityModeE.CellF3.Text;
			
			if(actualPanelText.Equals(sExpectedPanelText))
			{
				Report.Log(ReportLevel.Success,"Model name " +actualPanelText+ " is displayed successfully");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Model name" +sExpectedPanelText+ " is not displayed correctly instead " +actualPanelText+  "is displayed " );
			}
			
			repo.ShoppingListCompatibilityModeE.CellF14.Click();
			
			string actualDeviceText = repo.ShoppingListCompatibilityModeE.CellF14.Text;
			
			if(actualDeviceText.Equals(sExpectedDeviceText))
			{
				Report.Log(ReportLevel.Success,"Model name " +actualDeviceText+ " is displayed successfully");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Model name" +sExpectedDeviceText+ " is not displayed correctly instead " +actualDeviceText+  "is displayed " );
			}
		}
		
		
		/***********************************************************************************************************
		 * Function Name: verifyShoppingListDevicesTextForCell17And21
		 * 		 * Function Details: To verify shopping list devices via clicking on its row
		 * Parameter/Arguments: sFileName,sDeviceSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 01/06/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyShoppingListDevicesTextForCell17And21(string sExpectedPanelText, string sExpectedDeviceText)
		{
			repo.ShoppingListCompatibilityModeE.CellF17.Click();
			
			string actualPanelText = repo.ShoppingListCompatibilityModeE.CellF17.Text;
			
			if(actualPanelText.Equals(sExpectedPanelText))
			{
				Report.Log(ReportLevel.Success,"Model name " +actualPanelText+ " is displayed successfully");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Model name" +sExpectedPanelText+ " is not displayed correctly instead " +actualPanelText+  "is displayed " );
			}
			
			repo.ShoppingListCompatibilityModeE.CellF21.Click();
			
			string actualDeviceText = repo.ShoppingListCompatibilityModeE.CellF21.Text;
			
			if(actualDeviceText.Equals(sExpectedDeviceText))
			{
				Report.Log(ReportLevel.Success,"Model name " +actualDeviceText+ " is displayed successfully");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Model name" +sExpectedDeviceText+ " is not displayed correctly instead " +actualDeviceText+  "is displayed " );
			}
			
		}
		
		/********************************************************************
		 * Function Name: AddDevicesfromMultiplePointWizardWithRegion
		 * Function Details: To add multiple devices using multiple point wizard with different loop
		 * Parameter/Arguments: Device name, Region and its quantity
		 * Output:
		 * Function Owner: Poonam Kadam
		 * Last Update : 23/05/19
		 ********************************************************************/
		[UserCodeMethod]
		public static void AddDevicesfromMultiplePointWizardWithRegion(string sDeviceName,int DeviceQty, string sRegion )
		{
			repo.FormMe.btn_MultiplePointWizard.Click();
			//repo.ProfileConsys1.btn_MultiplePointWizard_DoNotUse.Click();
			repo.AddDevices.txt_AllDevices.Click();
			
			repo.AddDevices.txt_SearchDevices.Click();
			Keyboard.Press("{LControlKey down}{Akey}{LControlKey up}"+sDeviceName);
			ModelNumber = sDeviceName;
			repo.AddDevices.txt_ModelNumber.Click();
			repo.AddDevices.txt_Quantity.Click();
			Keyboard.Press("{LControlKey down}{Akey}{LControlKey up}"+DeviceQty.ToString());
			repo.AddDevices.MPWRegionBox.Click();
			repo.AddDevices.MultiPointWizardRegionDropDownBtn.Click();
			sRowIndex=sRegion;
			Report.Log(ReportLevel.Success,"sRowIndex="+sRowIndex);
			repo.ContextMenu.MultiPointWizardRegionComboBox.Click();
			repo.AddDevices.btn_AddDevices.Click();
			Report.Log(ReportLevel.Success,+DeviceQty+" \""+sDeviceName+ "\" Device Added successfully");
			Delay.Milliseconds(200);

		}
		
		
		/********************************************************************
		 * Function Name: MoveScrollBarDownInPointsGrid
		 * Function Details: To move scroll bar down vertically in points grid
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Poonam Kadam
		 * Last Update : 23/05/19
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifySelectedNode()
		{
			string stext=repo.FormMe.NavigationTree.SelectedItems.ToString();
			Report.Log(ReportLevel.Success," Device Added successfully"+stext);
		}

		/********************************************************************
		 * Function Name: MoveScrollBarDownInPointsGrid
		 * Function Details: To verify selected node
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Poonam Kadam
		 * Last Update : 27/05/19
		 ********************************************************************/
		[UserCodeMethod]
		public static void MoveScrollBarDownInPointsGrid()
		{
			// Create a adapter and stored in source adapter element
			//Adapter sourceE = repo.FormMe.UpArrowScrollButtonPointsGrid;//HorizontalScrollBarPointsGrid;
			
			// Create a adapter and stored in targer adapter element
			//Adapter targetE = repo.FormMe.DownArrowScrollButtonPointsGrid;

			// Drag scroll bar from First position to its defined position
			//Ranorex.AutomationHelpers.UserCodeCollections.DragNDropLibrary.DragAndDrop(sourceE,targetE);
			//Mouse.ButtonDown(System.Windows.Forms.MouseButtons.Left);
			//repo.FormMe.HorizontalScrollBarPointsGrid.
			SelectPointsGridRow("1");
			Keyboard.Press("{PageDown}");
			Keyboard.Press("{PageDown}");
			Keyboard.Press("{PageDown}");
			Keyboard.Press("{PageDown}");
			Keyboard.Press("{PageDown}");
			Keyboard.Press("{PageDown}");
		}
		
		/********************************************************************
		 * Function Name:verifyExportButtonInGalleryEnabledOrDisabled
		 * Function Details: To verify if Export button is enabled or disabled in Ribbon
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Poonam Kadam
		 * Last Update : 03/6/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void verifyExportButtonInGalleryEnabledOrDisabled()
		{
			if(repo.FormMe.Export2ndTime.Enabled)
			{
				Report.Log(ReportLevel.Success, "Export button is enabled");
			}
			else
			{
				Report.Log(ReportLevel.Success,"Export button is disabled");
			}
		}
		
		/********************************************************************
		 * Function Name: VerifyLabelInPanelAccessories
		 * Function Details: To verify label text  in Panel Accessories
		 * Parameter/Arguments:expected Label text
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 04/06/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyLabelInPanelAccessories(string sLabel)
		{
			repo.FormMe.PanelAccessoriesLabel.Click();
			
			// Retrieve value of label
			string actualLabel = repo.FormMe.PanelAccessoriesLabel.TextValue;
			
			// Comparing actualLabel and sLabel values
			if(actualLabel.Equals(sLabel))
			{
				Report.Log(ReportLevel.Success,"Label text " +actualLabel+ " is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Label text is displayed as "+actualLabel+ " instead of "+ sLabel);
			}
		}
		
		/********************************************************************
		 * Function Name: ChangeCableLengthFromInventory
		 * Function Details: To change cable length from inventory properties section
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Poonam Kadam
		 * Last Update :07/Jun/19
		 ********************************************************************/
		// Change cable length method
		[UserCodeMethod]
		public static void ChangeCableLengthFromInventory(int fchangeCableLength)
		{
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Enter the Day Matches night text in Search Properties fields to view cable length;
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("Leng" +"{ENTER}" );
			
			//Click on cable length cell
			repo.FormMe.txt_InventoryProperty.Click();
			
			//Change the value of cable length
			repo.FormMe.txt_InventoryProperty.PressKeys(fchangeCableLength + "{ENTER}");
			repo.ProfileConsys1.txt_SearchProperties.Click();
			Keyboard.Press("{LControlKey down}{Akey}{LControlKey up}{ENTER}");
		}
		
		/********************************************************************
		 * Function Name:verifyContextMenuOptionTextOnRightClickInPointsGrid
		 * Function Details: To verify if context menu option text when we right click on points grid row
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Poonam
		 * Last Update : 14/6/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void verifyContextMenuOptionTextOnRightClickInPointsGrid(string sContectMenuOption)
		{
			sListIndex=sContectMenuOption;
			if(repo.ContextMenu.ColumnChooserListText.Visible)
			{
				Report.Log(ReportLevel.Success, sContectMenuOption+" option is displayed");
			}
			else{
				Report.Log(ReportLevel.Success, sContectMenuOption+" option is not displayed");
			}
		}
		
	}
}