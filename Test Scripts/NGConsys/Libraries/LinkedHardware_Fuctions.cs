/*
 * Created by Ranorex
 * User: jbhasip
 * Date: 8/9/2019
 * Time: 3:12 PM
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
using System.Windows;

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace TestProject.Libraries
{
	/// <summary>
	/// Ranorex user code collection. A collection is used to publish user code methods to the user code library.
	/// </summary>
	[UserCodeCollection]
	public class LinkedHardware_Fuctions
	{
		// You can use the "Insert New User Code Method" functionality from the context menu,
		// to add a new method with the attribute [UserCodeMethod].
		
		//Create instance of repository to access repository items
		static NGConsysRepository repo = NGConsysRepository.Instance;
		
		/**************************************************************************************************************************************
		 * Function Name: VerifyLinkedDevicesGetAddedInLoop
		 * Function Details: Add a device and its child till Max Limit and check if linked devices get added in the Loop
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 9/08/2019  Alpesh Dhakad - 23/08/2019 - Updated with new navigation tree method, xpath and panel type column number
		 * Alpesh Dhakad - 14/05/2020 - Updated test script as per new implementation
		 **************************************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyLinkedDevicesGetAddedInLoop(string sFileName,string sSheetName)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sSheetName);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared string type
			string ParentDeviceName,ParentDeviceType,ChildDeviceName,ChildDeviceType,ParentLabel,expectedLabel1,expectedLabel2,PanelType;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=10; i++)
			{
				ParentDeviceName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				ParentDeviceType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				ParentLabel = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				ChildDeviceName = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				ChildDeviceType = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				
				//Click on  node Expander
				Common_Functions.ClickOnNavigationTreeExpander("Node");
				
				// Click on Loop Card node
				//Common_Functions.ClickOnNavigationTreeExpander(PanelType);
				
				// Click on Panel node
				Common_Functions.ClickOnNavigationTreeItem("Node");
				
				//Add parent Device
				Devices_Functions.AddDevicesfromPanelNodeGallery(ParentDeviceName,ParentDeviceType,PanelType);
				
				//Select row of parent Device
				Devices_Functions.SelectRowUsingLabelNameFromInventoryTab(ParentLabel);
				
				
				if(!ChildDeviceName.IsEmpty())
				{
					//Select Child Device
					Devices_Functions.AddDevicesfromPanelNodeGallery(ChildDeviceName,ChildDeviceType,PanelType);
				}
				
			}
			//Verify Linked Devices are added in Loop A
			
			// Click on Loop A node
			Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
			
			
			//Click Points Tab
			
			for(int i=8; i<=10; i++)
			{
				expectedLabel1 = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				expectedLabel2 = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				
				Devices_Functions.VerifyDeviceUsingLabelName(expectedLabel1);
				Devices_Functions.VerifyDeviceUsingLabelName(expectedLabel1);
			}
			
		}
		
		/********************************************************************
		 * Function Name: VerifyCheckboxOfProperties
		 * Function Details: Add a device and its child till Max Limit and check if linked devices get added in the Loop
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 22/08/2019
		 *  Alpesh Dhakad - 13/05/2020 - Updated test script as per new implementation
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyCheckboxOfProperties(string sFileName,string sSheetName)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sSheetName);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count,LabelNo;
			
			// Declared string type
			string DeviceName,sType,LabelName,CheckboxName,sDefaultStatus,SKUNo,sStatus,sSecondChannelCheckbox,sStatusInLoopB;
			
			LabelNo = 126; //For Loop B
			
			//Expand Panel Node
			Common_Functions.ClickOnNavigationTreeExpander("Node");
			
			//Expand FIM/PFI Loop card
			//Common_Functions.ClickOnNavigationTreeExpander("PFI");
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				DeviceName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				LabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				CheckboxName = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sDefaultStatus = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				SKUNo = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				sStatus = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				sSecondChannelCheckbox = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				sStatusInLoopB = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				
				bool DefaultStatus = Convert.ToBoolean(sDefaultStatus);
				bool Status = Convert.ToBoolean(sStatus);
				bool StatusInLoopB = Convert.ToBoolean(sStatusInLoopB);
				string sLabelNo = LabelNo.ToString();
				
				//Click on Loop A
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				//Add Devices
				Devices_Functions.AddDevicesfromGallery(DeviceName,sType);
				
				//Click on Points Tab
				Common_Functions.clickOnPointsTab();
				
				//Click on device in grid
				Devices_Functions.SelectRowUsingLabelName(LabelName);
				
				//Verify checkbox
				Devices_Functions.verifyCheckboxInSearchProperties(CheckboxName,DefaultStatus);
				
				
				
				//Click on Points Tab
				Common_Functions.clickOnPointsTab();
				
				//Click on device in grid
				Devices_Functions.SelectRowUsingLabelName(LabelName);
				
				//Click on copy
				repo.FormMe.btn_Copy.Click();
				
				//Click on Loop A
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-B");
				
				//Click on Points Tab
				Common_Functions.clickOnPointsTab();
				
				//Click on Paste
				Common_Functions.clickOnPasteButton();
				
				string NewLabelName = DeviceName+" - "+sLabelNo;
				
				//Verify device is pasted
				Devices_Functions.VerifyDeviceUsingLabelName(NewLabelName);
				
				//Verify checkbox
				//Devices_Functions.VerifyCheckboxExists(CheckboxName,DefaultStatus);
				
				
				Devices_Functions.verifyCheckboxInSearchProperties(CheckboxName,StatusInLoopB);
				
				//Click on checkbox
				Devices_Functions.ClickCheckboxInSearchProperties(CheckboxName);
				
				
				
				//Slect 2nd channel
				if(!sSecondChannelCheckbox.IsEmpty())
				{
					LabelNo = LabelNo+1;
					
					sLabelNo = LabelNo.ToString();

					string SecondChannelLabelName = DeviceName+" - "+sLabelNo;
					bool SecondChannelCheckbox = Convert.ToBoolean(sSecondChannelCheckbox);
					Devices_Functions.SelectRowUsingLabelName(SecondChannelLabelName);
					
					Devices_Functions.VerifyCheckboxExists(CheckboxName,SecondChannelCheckbox);
				}
				
				
				//Click on Site Node and Shopping List
				// Click on Site node
				Common_Functions.ClickOnNavigationTreeItem("Site");
				
				Common_Functions.clickOnShoppingListTab();
				
				//Verify in shopping list
				Export_Functions.SearchDeviceInExportUsingSKUOrDescription(SKUNo,true);
				LabelNo = LabelNo+1;
			}
		}
		
		/********************************************************************
		 * Function Name: VerifySurfaceBoxParameters
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 23/08/2019
		 * Alpesh Dhakad - 14/05/2020 Updated script as per new structure
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifySurfaceBoxParameters(string sFileName,string sSheetName)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sSheetName);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared string type
			string DeviceName,sType,LabelName,sPropertyName,sParameter,SKUNo;
			
			//Expand Panel Node
				Common_Functions.ClickOnNavigationTreeExpander("Node");
				
				//Expand FIM/PFI Loop card
				//Common_Functions.ClickOnNavigationTreeExpander("PFI");
				
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				DeviceName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				LabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				sPropertyName = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sParameter = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				SKUNo = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				
				
				//Click on Loop A
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				//Add Devices
				Devices_Functions.AddDevicesfromGallery(DeviceName,sType);
				
				//Click on Points Tab
				Common_Functions.clickOnPointsTab();
				
				//Click on device in grid
				Devices_Functions.SelectRowUsingLabelName(LabelName);
				
				//Select parameter
				Devices_Functions.ChangeParameterInSearchProperties(sPropertyName,sParameter);
				
				//Click on Site Node and Shopping List
				// Click on Site node
				Common_Functions.ClickOnNavigationTreeItem("Site");
				
				Common_Functions.clickOnShoppingListTab();
				
				
				
				//Verify in shopping list
				//Export_Functions.SearchDeviceInExportUsingSKUOrDescription(SKUNo,true);
				//Export_Functions.SearchDeviceInExportUsingSKUOrDescription(sParameter,true);
			}
		}
	}
}
