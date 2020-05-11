/*
 * Created by Ranorex
 * User: jdhakaa
 * Date: 5/15/2019
 * Time: 10:59 AM
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
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;


using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace TestProject.Libraries
{
	/// <summary>
	/// Ranorex user code collection. A collection is used to publish user code methods to the user code library.
	/// </summary>
	[UserCodeCollection]
	public class SlotCards_Functions
	{
		// You can use the "Insert New User Code Method" functionality from the context menu,
		// to add a new method with the attribute [UserCodeMethod].
		
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
		
		static string sRowIndex
		{
			get { return repo.sRowIndex; }
			set { repo.sRowIndex = value; }
		}
		
		static string sMainProcessorGalleryIndex
		{
			get { return repo.sMainProcessorGalleryIndex; }
			set { repo.sMainProcessorGalleryIndex = value; }
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
		
		static string sOtherSlotCardName
		{
			get { return repo.sOtherSlotCardName; }
			set { repo.sOtherSlotCardName = value; }
		}
		
		
		
		/********************************************************************************************************
		 * Function Name: VerifySlotCardsAndBackplanesDistribution
		 * Function Details: To verify slot cards and backplane distribution
		 * Parameter/Arguments: sFileName, sAddDevicesSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 11/05/2019 Alpesh Dhakad - 30/07/2019 & 20/08/2019 - Updated test script as per new build xpath
		 ********************************************************************************************************/
		[UserCodeMethod]
		public static void VerifySlotCardsAndBackplanesDistribution(string sFileName,string sAddDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			int columns = Excel_Utilities.ExcelRange.Columns.Count;
			int deviceCount;
			// Declared string type
			string sType,sDeviceCount,sDeviceName,PanelType,ExpectedBackplane1,ExpectedBackplane2,ExpectedBackplane3,PanelName,PanelNode,CPUType;
			string sBackplane1SlotCardName,sBackplane2SlotCardName,sBackplane3SlotCardName;

			for(int i=10; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				ExpectedBackplane1 = ((Range)Excel_Utilities.ExcelRange.Cells[i,11]).Value.ToString();
				ExpectedBackplane2 = ((Range)Excel_Utilities.ExcelRange.Cells[i,12]).Value.ToString();
				ExpectedBackplane3 = ((Range)Excel_Utilities.ExcelRange.Cells[i,13]).Value.ToString();
				sBackplane1SlotCardName = ((Range)Excel_Utilities.ExcelRange.Cells[i,14]).Value.ToString();
				sBackplane2SlotCardName = ((Range)Excel_Utilities.ExcelRange.Cells[i,15]).Value.ToString();
				sBackplane3SlotCardName = ((Range)Excel_Utilities.ExcelRange.Cells[i,16]).Value.ToString();
				
				
				// Add panels
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				
				for(int j=4; j<=10; j++){
					
					
					sDeviceName =  ((Range)Excel_Utilities.ExcelRange.Cells[8,j]).Value.ToString();
					sType = ((Range)Excel_Utilities.ExcelRange.Cells[9,j]).Value.ToString();
					sDeviceCount = ((Range)Excel_Utilities.ExcelRange.Cells[i,j]).Value.ToString();
					
					
					PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[4,7]).Value.ToString();
					
					int.TryParse(sDeviceCount, out deviceCount);
					
					// Verify device count and then add devices from panel accessories gallery or panel node gallery
					if(deviceCount>0)
					{
						if (sType.Equals("Accessories"))
						{
							// Click on Panel node
						Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
							Common_Functions.clickOnPanelAccessoriesTab();
							for(int k=1; k<=deviceCount;k++)
							{
								Devices_Functions.AddDevicefromPanelAccessoriesGallery(sDeviceName,sType);
							}
						}
						else
						{
							// Click on Panel node
						Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
						Common_Functions.clickOnInventoryTab();
							
							for(int k=1; k<=deviceCount;k++)
							{
								Devices_Functions.AddDevicesfromPanelNodeGallery(sDeviceName,sType,PanelType);
							}
						}
					}
					
				}
				
				// Click on Expander node
				Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
				
				
				// Verify expected backplane1
				if(ExpectedBackplane1.Equals("Yes"))
				{
					//Common_Functions.ClickOnNavigationTreeExpander("Backplane  1/3");
					//Common_Functions.VerifyNavigationTreeItemText("Backplane  1/3");
					Common_Functions.VerifyAndClickNavigationTreeItemText(sBackplane1SlotCardName);
					//Common_Functions.ClickOnNavigationTreeExpander("Backplane  1/3");
					
				}
				else
				{
					Common_Functions.VerifyNavigationTreeItem("Backplane  1/3", ExpectedBackplane1);
				}
				
				// Verify expected backplane2
				if(ExpectedBackplane2.Equals("Yes"))
				{
					Common_Functions.ClickOnNavigationTreeExpander("Backplane  2/3");
					Common_Functions.VerifyNavigationTreeItemText("Backplane  2/3");
					Common_Functions.VerifyAndClickNavigationTreeItemText(sBackplane2SlotCardName);
					Common_Functions.ClickOnNavigationTreeExpander("Backplane  2/3");
				}
				else
				{
					Common_Functions.VerifyNavigationTreeItem("Backplane  2/3", ExpectedBackplane2);
					
					
				}
				
				// Verify expected backplane3
				if(ExpectedBackplane3.Equals("Yes"))
				{
					Common_Functions.ClickOnNavigationTreeExpander("Backplane  3/3");
					Common_Functions.VerifyNavigationTreeItemText("Backplane  3/3");
					Common_Functions.VerifyAndClickNavigationTreeItemText(sBackplane3SlotCardName);
					Common_Functions.ClickOnNavigationTreeExpander("Backplane  3/3");
						
				}
				else
				{
					Common_Functions.VerifyNavigationTreeItem("Backplane  3/3", ExpectedBackplane3);
					
				}
				
				if(rows!=10)
				{
					// Delete Panel
					Panel_Functions.DeletePanel(1,PanelNode,1);
				}
				
			}
			
			// Close Excel
			Excel_Utilities.CloseExcel();
		}
		
		
		
		/********************************************************************************************************************
		 * Function Name: VerifySlotCardsAndBackplanesDistributionWithOnePanel
		 * Function Details: To verify slot cards and backplane distribution
		 * Parameter/Arguments: sFileName, sAddDevicesSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 13/05/2019  Alpesh Dhakad - 30/07/2019 - Updated test script as per new build xpaths
		 ********************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifySlotCardsAndBackplanesDistributionWithOnePanel(string sFileName,string sAddDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			int columns = Excel_Utilities.ExcelRange.Columns.Count;
			int deviceCount;
			// Declared string type
			string sType,sDeviceCount,sDeviceName,PanelType,ExpectedBackplane1,ExpectedBackplane2,ExpectedBackplane3,PanelName,PanelNode,CPUType;
			string sBackplane1SlotCardName,sBackplane2SlotCardName,sBackplane3SlotCardName;

			for(int i=10; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				ExpectedBackplane1 = ((Range)Excel_Utilities.ExcelRange.Cells[i,11]).Value.ToString();
				ExpectedBackplane2 = ((Range)Excel_Utilities.ExcelRange.Cells[i,12]).Value.ToString();
				ExpectedBackplane3 = ((Range)Excel_Utilities.ExcelRange.Cells[i,13]).Value.ToString();
				sBackplane1SlotCardName = ((Range)Excel_Utilities.ExcelRange.Cells[i,14]).Value.ToString();
				sBackplane2SlotCardName = ((Range)Excel_Utilities.ExcelRange.Cells[i,15]).Value.ToString();
				sBackplane3SlotCardName = ((Range)Excel_Utilities.ExcelRange.Cells[i,16]).Value.ToString();
				
				// Add panels
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				
				Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
				
				for(int j=4; j<=10; j++){
					
					
					sDeviceName =  ((Range)Excel_Utilities.ExcelRange.Cells[8,j]).Value.ToString();
					sType = ((Range)Excel_Utilities.ExcelRange.Cells[9,j]).Value.ToString();
					sDeviceCount = ((Range)Excel_Utilities.ExcelRange.Cells[i,j]).Value.ToString();
					
					
					PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[4,7]).Value.ToString();
					
					int.TryParse(sDeviceCount, out deviceCount);
					
					// Verify device count and then add devices from panel accessories gallery or panel node gallery
					if(deviceCount>0)
					{
						if (sType.Equals("Accessories"))
						{
							Common_Functions.ClickOnNavigationTreeItem(PanelNode);
							//repo.FormMe.PanelNode1.Click();
							//repo.ProfileConsys1.NavigationTree.Expander.Click();
							Common_Functions.clickOnPanelAccessoriesTab();
							for(int k=1; k<=deviceCount;k++)
							{
								//Devices_Functions.AddDevicefromPanelAccessoriesGallery(sDeviceName,sType);
								Devices_Functions.AddDevicesfromGallery(sDeviceName,sType);
							}
						}
						else
						{
							Common_Functions.ClickOnNavigationTreeItem(PanelNode);
							//repo.FormMe.PanelNode1.Click();
							//repo.ProfileConsys1.NavigationTree.Expander.Click();
							Common_Functions.clickOnInventoryTab();
							
							for(int k=1; k<=deviceCount;k++)
							{
								//Devices_Functions.AddDevicesfromPanelNodeGallery(sDeviceName,sType,PanelType);
								Devices_Functions.AddDevicesfromGallery(sDeviceName,sType);
							}
						}
					}
					
				}
				
				// Verify expected backplane1
				if(ExpectedBackplane1.Equals("Yes"))
				{
					//Common_Functions.ClickOnNavigationTreeExpander("Backplane  1/3");
					//Common_Functions.VerifyNavigationTreeItemText("Backplane  1/3");
					Common_Functions.VerifyAndClickNavigationTreeItemText(sBackplane1SlotCardName);
					//Common_Functions.ClickOnNavigationTreeExpander("Backplane  1/3");
					
				}
				else
				{
					Common_Functions.VerifyNavigationTreeItem("Backplane  1/3", ExpectedBackplane1);
				}
				
				// Verify expected backplane2
				if(ExpectedBackplane2.Equals("Yes"))
				{
					Common_Functions.ClickOnNavigationTreeExpander("Backplane  2/3");
					//Common_Functions.VerifyNavigationTreeItemText("Backplane  2/3");
					Common_Functions.VerifyAndClickNavigationTreeItemText(sBackplane2SlotCardName);
					//Common_Functions.ClickOnNavigationTreeExpander("Backplane  2/3");
				}
				else
				{
					Common_Functions.VerifyNavigationTreeItem("Backplane  2/3", ExpectedBackplane2);
					
					
				}
				
				// Verify expected backplane3
				if(ExpectedBackplane3.Equals("Yes"))
				{
					//Common_Functions.ClickOnNavigationTreeExpander("Backplane  3/3");
					//Common_Functions.VerifyNavigationTreeItemText("Backplane  3/3");
					Common_Functions.VerifyAndClickNavigationTreeItemText(sBackplane3SlotCardName);
					//Common_Functions.ClickOnNavigationTreeExpander("Backplane  3/3");
						
				}
				else
				{
					Common_Functions.VerifyNavigationTreeItem("Backplane  3/3", ExpectedBackplane3);
					
				}
				
			}
			
			// Close Excel
			Excel_Utilities.CloseExcel();
		}
		
		/***********************************************************************************************************
		 * Function Name: VerifyPanelTypeDropdownOnSlotCardsPosition
		 * Function Details: string sFileName,string sAddDevicesSheet, string sPanelName
		 * Parameter/Arguments: expectedText
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 15/05/2019  Alpesh Dhakad - 23/08/2019 - Updated with new navigation tree method,
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyPanelTypeDropdownOnSlotCardsPosition(string sFileName,string sAddDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			int columns = Excel_Utilities.ExcelRange.Columns.Count;
			int deviceCount;
			// Declared string type
			string sType,sDeviceCount,sDeviceName,PanelType,ExpectedBackplane1,ExpectedBackplane2,ExpectedBackplane3,PanelName,PanelNode,CPUType;
			string sBackplane1SlotCardName,sBackplane2SlotCardName,sBackplane3SlotCardName,PanelTypeNameList,PanelTypeNameListNotAvailable;

			for(int i=10; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				ExpectedBackplane1 = ((Range)Excel_Utilities.ExcelRange.Cells[i,11]).Value.ToString();
				ExpectedBackplane2 = ((Range)Excel_Utilities.ExcelRange.Cells[i,12]).Value.ToString();
				ExpectedBackplane3 = ((Range)Excel_Utilities.ExcelRange.Cells[i,13]).Value.ToString();
				sBackplane1SlotCardName = ((Range)Excel_Utilities.ExcelRange.Cells[i,14]).Value.ToString();
				sBackplane2SlotCardName = ((Range)Excel_Utilities.ExcelRange.Cells[i,15]).Value.ToString();
				sBackplane3SlotCardName = ((Range)Excel_Utilities.ExcelRange.Cells[i,16]).Value.ToString();
				PanelTypeNameList = ((Range)Excel_Utilities.ExcelRange.Cells[i,17]).Value.ToString();
				PanelTypeNameListNotAvailable = ((Range)Excel_Utilities.ExcelRange.Cells[i,18]).Value.ToString();
				// Add panels
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				
				for(int j=4; j<=10; j++){
					
					
					sDeviceName =  ((Range)Excel_Utilities.ExcelRange.Cells[8,j]).Value.ToString();
					sType = ((Range)Excel_Utilities.ExcelRange.Cells[9,j]).Value.ToString();
					sDeviceCount = ((Range)Excel_Utilities.ExcelRange.Cells[i,j]).Value.ToString();
					
					
					PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[4,7]).Value.ToString();
					
					int.TryParse(sDeviceCount, out deviceCount);
					
					// Verify device count and then add devices from panel accessories gallery or panel node gallery
					if(deviceCount>0)
					{
						if (sType.Equals("Accessories"))
						{
							// Click on Panel node
							Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
							Common_Functions.clickOnPanelAccessoriesTab();
							for(int k=1; k<=deviceCount;k++)
							{
								Devices_Functions.AddDevicefromPanelAccessoriesGallery(sDeviceName,sType);
							}
						}
						else
						{
							// Click on Panel node
							Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
							Common_Functions.clickOnInventoryTab();
							
							for(int k=1; k<=deviceCount;k++)
							{
								Devices_Functions.AddDevicesfromPanelNodeGallery(sDeviceName,sType,PanelType);
							}
						}
					}
					
				}
				
				//Expand Node
				Common_Functions.ClickOnNavigationTreeExpander("Node");
				
				// Verify expected backplane1
				if(ExpectedBackplane1.Equals("Yes"))
				{
					//Common_Functions.ClickOnNavigationTreeExpander("Backplane  1/3");
					Common_Functions.VerifyNavigationTreeItemText("Backplane  1/3");
					Common_Functions.VerifyAndClickNavigationTreeItemText(sBackplane1SlotCardName);
					//Common_Functions.ClickOnNavigationTreeExpander("Backplane  1/3");
					
				}
				else
				{
					Common_Functions.VerifyNavigationTreeItem("Backplane  1/3", ExpectedBackplane1);
				}
				
				// Verify expected backplane2
				if(ExpectedBackplane2.Equals("Yes"))
				{
					Common_Functions.ClickOnNavigationTreeExpander("Backplane  2/3");
					Common_Functions.VerifyNavigationTreeItemText("Backplane  2/3");
					Common_Functions.VerifyAndClickNavigationTreeItemText(sBackplane2SlotCardName);
					Common_Functions.ClickOnNavigationTreeExpander("Backplane  2/3");
				}
				else
				{
					Common_Functions.VerifyNavigationTreeItem("Backplane  2/3", ExpectedBackplane2);
					
					
				}
				
				// Verify expected backplane3
				if(ExpectedBackplane3.Equals("Yes"))
				{
					Common_Functions.ClickOnNavigationTreeExpander("Backplane  3/3");
					Common_Functions.VerifyNavigationTreeItemText("Backplane  3/3");
					Common_Functions.VerifyAndClickNavigationTreeItemText(sBackplane3SlotCardName);
					Common_Functions.ClickOnNavigationTreeExpander("Backplane  3/3");
						
				}
				else
				{
					Common_Functions.VerifyNavigationTreeItem("Backplane  3/3", ExpectedBackplane3);
					
				}
				
				Devices_Functions.VerifyPanelTypeInDropdown(PanelName,PanelTypeNameList,PanelTypeNameListNotAvailable);
				
				
				if(rows!=10)
				{
					// Delete Panel
					Panel_Functions.DeletePanel(1,PanelNode,1);
				}
			}
			
			// Close Excel
			Excel_Utilities.CloseExcel();

		}
		
		/***********************************************************************************************************
		 * Function Name: VerifyAddingRemovingOfTLI800SlotCards
		 * Function Details:
		 * Parameter/Arguments: string sFileName,string sAddDevicesSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 24/05/2019 Alpesh Dhakad - 23/08/2019 - Updated with new navigation tree method, xpath 
		 ***********************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyAddingRemovingOfTLI800SlotCards(string sFileName,string sAddDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared string type
			string sType,sDeviceName,PanelName,PanelNode,CPUType,initialState,afterDeletechangedState,newDeviceName,lastChangedState;
			
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				sType =  ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sDeviceName = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				initialState = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				newDeviceName = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				afterDeletechangedState = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				lastChangedState = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				
				// Add panels
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				
				// Click on Expander node
					Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
					
					// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
				
				Common_Functions.clickOnPanelAccessoriesTab();
				
				Devices_Functions.AddDevicefromPanelAccessoriesGallery(sDeviceName,sType);
				
				sDeviceName = newDeviceName;
				
				Devices_Functions.VerifyEnableDisablePanelAccessoriesGallery(sType,sDeviceName,initialState);
				
				Devices_Functions.DeleteAccessoryFromPanelAccessoriesTab();
				
				Devices_Functions.VerifyEnableDisablePanelAccessoriesGallery(sType,sDeviceName,afterDeletechangedState);
				
				Devices_Functions.AddDevicefromPanelAccessoriesGallery(sDeviceName,sType);
				
				Devices_Functions.VerifyEnableDisablePanelAccessoriesGallery(sType,sDeviceName,lastChangedState);
				
				if(rows!=8)
				{
					// Delete Panel
					Panel_Functions.DeletePanel(1,PanelNode,1);
				}
			}
			
			// Close Excel
			Excel_Utilities.CloseExcel();
			
		}
		
		/***********************************************************************************************************
		 * Function Name: VerifyTLI800Properties
		 * Function Details:
		 * Parameter/Arguments: string sFileName,string sAddDevicesSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 24/05/2019 Alpesh Dhakad - 23/08/2019 - Updated with new navigation tree method, xpath
		 ***********************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyTLI800Properties(string sFileName,string sAddDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared string type
			string sType,sDeviceName,PanelName,PanelNode,CPUType,sSKU,sModel,sLabel,sFOM;
			
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				sType =  ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sDeviceName = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				sLabel = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				sSKU = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				sModel = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				sFOM = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				
				
				// Add panels
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				
				// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
				// Click on panel accessories tab
				Common_Functions.clickOnPanelAccessoriesTab();
				
				// Add devices from panel accessories gallery
				Devices_Functions.AddDevicefromPanelAccessoriesGallery(sDeviceName,sType);
				
				// Verify label in search properties
				Devices_Functions.VerifyLabelInSearchProperties(sLabel);

				// Verify SKU in search properties
				Devices_Functions.VerifySKUInSearchProperties(sSKU);
				
				// Verify Model in search properties
				Devices_Functions.VerifyModelInSearchProperties(sModel);
				
				// Verify Description row in search properties
				Devices_Functions.VerifyDescriptionTextRowInSearchProperties();

				// Verify FOM in search properties
				Devices_Functions.VerifyFOMInSearchProperties(sFOM);
				
				// Verify MPM in search properties
				Devices_Functions.VerifyMPMInSearchProperties();
				
			}
			
			// Close Excel
			Excel_Utilities.CloseExcel();
		}
		
		/***********************************************************************************************************
		 * Function Name: VerifyShoppingListOnSelectingFOMandMPM
		 * Function Details:
		 * Parameter/Arguments: string sFileName,string sAddDevicesSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 30/05/2019  17/07/2019 - Alpesh Dhakad - Updated code
		 * Alpesh Dhakad - 23/08/2019 - Updated with new navigation tree method, xpath
		 ***********************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyShoppingListOnSelectingFOMandMPM(string sFileName,string sAddDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared string type
			string sType,sDeviceName,PanelName,PanelNode,CPUType,sSKU,sModel,sLabel,sFOMChange,sMPMExpectedState,sFOMExpectedText,sShoppingListCount;
			string sMPMExpectedText,sANNExpectedText,sShoppingListCountAfterUncheck,sMPMExpectedStateAgain,sMPMExpectedTextAfterUncheck,sANNExpectedTextAfterUncheck;
			bool changeCheckboxStateTo,changeCheckboxStateToAgain;
			int shoppingListCount,shoppingListCountAfterUncheck;
			
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				sType =  ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sDeviceName = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				sLabel = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				sSKU = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				sModel = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				sFOMChange = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				sMPMExpectedState = ((Range)Excel_Utilities.ExcelRange.Cells[i,10]).Value.ToString();
				sShoppingListCount = ((Range)Excel_Utilities.ExcelRange.Cells[i,11]).Value.ToString();
				sFOMExpectedText = ((Range)Excel_Utilities.ExcelRange.Cells[i,12]).Value.ToString();
				sMPMExpectedText = ((Range)Excel_Utilities.ExcelRange.Cells[i,13]).Value.ToString();
				sANNExpectedText = ((Range)Excel_Utilities.ExcelRange.Cells[i,14]).Value.ToString();
				sMPMExpectedStateAgain = ((Range)Excel_Utilities.ExcelRange.Cells[i,15]).Value.ToString();
				sShoppingListCountAfterUncheck = ((Range)Excel_Utilities.ExcelRange.Cells[i,16]).Value.ToString();
				sMPMExpectedTextAfterUncheck = ((Range)Excel_Utilities.ExcelRange.Cells[i,17]).Value.ToString();
				sANNExpectedTextAfterUncheck = ((Range)Excel_Utilities.ExcelRange.Cells[i,18]).Value.ToString();
				
				
				
				int.TryParse(sShoppingListCount, out shoppingListCount);
				int.TryParse(sShoppingListCountAfterUncheck, out shoppingListCountAfterUncheck);
				bool.TryParse(sMPMExpectedState, out changeCheckboxStateTo);
				bool.TryParse(sMPMExpectedStateAgain, out changeCheckboxStateToAgain);
				
				
				// Add panels
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				
				
				// Click on Panel node
				Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
				
				// Click on panel accessories tab
				Common_Functions.clickOnPanelAccessoriesTab();
				
				// Add devices from panel accessories gallery
				Devices_Functions.AddDevicefromPanelAccessoriesGallery(sDeviceName,sType);
				
				// Change FOM Value in Search properties
				Devices_Functions.ChangeFOMInSearchProperties(sFOMChange);
				
				// Verify and perform check or uncheck MPM checkbox in search properties
				Devices_Functions.CheckUncheckMPMCheckboxInSearchProperties(changeCheckboxStateTo);
				
				// Click on Site node
				Common_Functions.ClickOnNavigationTreeItem("Site");
			
				
				// Click on Shopping list tab
				Common_Functions.clickOnShoppingListTab();
					
				/******************************************************************************* 
                    Updated code - Purvi Bhasin (3/09/2019) Added the method for shopping list
			    ********************************************************************************/
				Export_Functions.SearchDeviceInExportUsingSKUOrDescription(sFOMExpectedText,true);
				
				Export_Functions.SearchDeviceInExportUsingSKUOrDescription(sMPMExpectedText,true);
				
				Export_Functions.SearchDeviceInExportUsingSKUOrDescription(sANNExpectedText,true);
				
				// Click on Panel node
				Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
				// Click on panel accessories tab
				Common_Functions.clickOnPanelAccessoriesTab();
				
				// Verify and perform check or uncheck MPM checkbox in search properties
				Devices_Functions.CheckUncheckMPMCheckboxInSearchProperties(changeCheckboxStateToAgain);
				
				// Click on Site node
				Common_Functions.ClickOnNavigationTreeItem("Site");
			
				// Click on shopping list tab
				Common_Functions.clickOnShoppingListTab();
				
				Export_Functions.SearchDeviceInExportUsingSKUOrDescription(sFOMExpectedText,true);
				
				Export_Functions.SearchDeviceInExportUsingSKUOrDescription(sMPMExpectedText,false);
				
				Export_Functions.SearchDeviceInExportUsingSKUOrDescription(sANNExpectedText,false);
								
				// Click on Panel Node
				Common_Functions.ClickOnNavigationTreeItem("Node");
				// Click on Site node
				Common_Functions.ClickOnNavigationTreeItem("Site");
			
				// Verify if row count is more than 8 then delete the panel
				if(rows!=8)
				{
					// Delete Panel
					Panel_Functions.DeletePanel(1,PanelNode,1);
				}
				
			}
			
			// Close Excel
			Excel_Utilities.CloseExcel();
		}
		
		
		/***********************************************************************************************************
		 * Function Name: VerifyShoppingListOnAddingTLI800AndTLI800EN
		 * Function Details:
		 * Parameter/Arguments: string sFileName,string sAddDevicesSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 27/05/2019  17/07/2019 - Alpesh Dhakad - Updated code
		 * Alpesh Dhakad - 23/08/2019 - Updated with new navigation tree method, xpath 
		 ***********************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyShoppingListOnAddingTLI800AndTLI800EN(string sFileName,string sAddDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared string type
			string sType,sDeviceName,PanelName,PanelNode,CPUType,initialState,sShoppingListCount,newDeviceName,secondPanelName;
			string sSecondCPUType,sSecondDeviceName,sSecondDeviceType,sSecondShoppingListCount,secondDisableDevice,secondDeviceState,secondPanelNode;
			
			int shoppingListCount,secondShoppingListCount;
			
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				sType =  ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sDeviceName = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				initialState = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				newDeviceName = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				sShoppingListCount = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				secondPanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				sSecondCPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,10]).Value.ToString();
				sSecondDeviceName = ((Range)Excel_Utilities.ExcelRange.Cells[i,11]).Value.ToString();
				sSecondDeviceType =  ((Range)Excel_Utilities.ExcelRange.Cells[i,12]).Value.ToString();
				secondDeviceState =  ((Range)Excel_Utilities.ExcelRange.Cells[i,13]).Value.ToString();
				sSecondShoppingListCount =  ((Range)Excel_Utilities.ExcelRange.Cells[i,14]).Value.ToString();
				secondDisableDevice =  ((Range)Excel_Utilities.ExcelRange.Cells[i,15]).Value.ToString();
				secondPanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,16]).Value.ToString();
				
				
				int.TryParse(sShoppingListCount, out shoppingListCount);
				int.TryParse(sSecondShoppingListCount, out secondShoppingListCount);
				
				
				// Add panels
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				
				// Click on Panel node
				Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
				// Click on panel accessories tab
				Common_Functions.clickOnPanelAccessoriesTab();
				
				// Add devices from panel accessories gallery
				Devices_Functions.AddDevicefromPanelAccessoriesGallery(sDeviceName,sType);
				
				// Set newDevice name value in sDeviceName
				sDeviceName = newDeviceName;
				
				// Verify Enable or disable of devices in panel accessories gallery
				Devices_Functions.VerifyEnableDisablePanelAccessoriesGallery(sType,sDeviceName,initialState);
				
				// Click on Site node
				Common_Functions.ClickOnNavigationTreeItem("Site");
			
				// Click on shopping list tab
				//Common_Functions.clickOnShoppingListTab();
				
				// Click on panel accessories tab
				//repo.FormMe.tab_Panel_Network.Click();
				
				// Click on Site node
				Common_Functions.ClickOnNavigationTreeItem("Site");
			
				// Add one panel after adding 1 one panel
				Panel_Functions.AddOnePanel(2,secondPanelName,sSecondCPUType);
				
				// Click on Panel node
				Common_Functions.ClickOnNavigationTreeItem(secondPanelNode);
			
				// Click on panel accessories tab
				Common_Functions.clickOnPanelAccessoriesTab();
				
				// Set newDevice name value in sDeviceName
				sDeviceName = sSecondDeviceName;
				
				// Set sType value
				sType = sSecondDeviceType;
				
				// Add devices from panel accessories gallery
				Devices_Functions.AddDevicefromPanelAccessoriesGallery(sDeviceName,sType);
				
				// set sDeviceName
				sDeviceName = secondDisableDevice;
				
				// Verify Enable or disable of devices in panel accessories gallery
				Devices_Functions.VerifyEnableDisablePanelAccessoriesGallery(sType,sDeviceName,secondDeviceState);
				
				// Click on Site node
				Common_Functions.ClickOnNavigationTreeItem("Site");
			
				
				// Click on Site node
				Common_Functions.ClickOnNavigationTreeItem("Site");
			
				
				// Click on shopping list tab
				Common_Functions.clickOnShoppingListTab();
				
				/*********************************************************************
				 Updated - Purvi Bhasin (3/09/2019) - Shopping list method is added
				 *********************************************************************/
				
				//Verify device is present in shopping list 
				Export_Functions.SearchDeviceInExportUsingSKUOrDescription(PanelName,true);
				
				Export_Functions.SearchDeviceInExportUsingSKUOrDescription(sDeviceName,true);
				
				Export_Functions.SearchDeviceInExportUsingSKUOrDescription(secondPanelName,true);
				
				Export_Functions.SearchDeviceInExportUsingSKUOrDescription(sSecondDeviceName,true);
					
				// Click on Panel Node
				Common_Functions.ClickOnNavigationTreeItem("Node");
				// Click on Site node
				Common_Functions.ClickOnNavigationTreeItem("Site");
			
				// Verify if row count is more than 8 then delete the panel
				if(rows!=8)
				{
					// Delete Panel
					Panel_Functions.DeletePanel(2,PanelNode,1);
				}
				
			}
			// Close Excel sheet
			Excel_Utilities.CloseExcel();
		}
		
		/***********************************************************************************************************
		 * Function Name: VerifyIOBInAccessoriesGallery
		 * Function Details: To Verify IOB state In Accessories Gallery
		 * Parameter/Arguments: string sFileName,string sAddDevicesSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 03/06/2019 Alpesh Dhakad - 30/07/2019 - Updated script as per new build updated xpath
		 * Alpesh Dhakad - 06/08/2019 - Added Site node click event before delete panel step
		 * Alpesh Dhakad - 20/08/2019 - Updated script as per new build updated xpath
		 ***********************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyIOBInAccessoriesGallery(string sFileName,string sAddDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared string type
			string sType,sDeviceName,PanelName,PanelNode,CPUType,InitialState;
			
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				sType =  ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sDeviceName = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				InitialState = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				
				// Add panels
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				
				// Click on Panel node
				Common_Functions.ClickOnNavigationTreeItem(PanelNode);

				// Click on panel accessories tab
				Common_Functions.clickOnPanelAccessoriesTab();
				
				// Verify enable or disable state of devices in panel accessories gallery
				Devices_Functions.VerifyEnableDisablePanelAccessoriesGallery(sType, sDeviceName, InitialState);

				// Click on Site node
				Common_Functions.ClickOnNavigationTreeItem("Site");

				
				// Verify if row count is more than 8 then delete the panel
				if(rows!=8)
				{
					// Delete Panel
					Panel_Functions.DeletePanel(1,PanelNode,1);
				}
			}
			// Close open excel file
			Excel_Utilities.CloseExcel();
		}
		
		/***********************************************************************************************************
		 * Function Name: VerifyIOBProperties
		 * Function Details: To Verify IOB properties
		 * Parameter/Arguments: string sFileName,string sAddDevicesSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 03/06/2019 Alpesh Dhakad - 30/07/2019 & 20/08/2019 - Updated script as per new build updated xpath
		 * 05/05/2020 - Alpesh Dhakad - Updated script as per new implementation
		 ***********************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyIOBProperties(string sFileName,string sAddDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared string type
			string sType,sDeviceName,PanelName,PanelNode,CPUType,sLabel,sSKU,sModel,sNewLabel;
			
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				sType =  ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sDeviceName = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				sLabel = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				sSKU = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				sModel = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				sOtherSlotCardName = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				sNewLabel = ((Range)Excel_Utilities.ExcelRange.Cells[i,10]).Value.ToString();
				
				// Add panels
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				
				/// Click on Panel node
				Common_Functions.ClickOnNavigationTreeItem(PanelNode);

				// Click on panel accessories tab
				Common_Functions.clickOnPanelAccessoriesTab();
				
				// Add devices from panel accessories gallery
				Devices_Functions.AddDevicefromPanelAccessoriesGallery(sDeviceName,sType);
				
				//  Verify label in Search properties
				Devices_Functions.VerifyLabelInSearchProperties(sLabel);
				
				// Verify SKU in Search properties
				Devices_Functions.VerifySKUInSearchProperties(sSKU);
				
				// Verify Model in Search properties
				Devices_Functions.VerifyModelInSearchProperties(sModel);
				
				// Verify Description field in Search properties
				Devices_Functions.VerifyDescriptionTextRowInSearchProperties();
				
				// Click on Expander node
				Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
				
				// // Verify and click backplane1 expander
				//Common_Functions.ClickOnNavigationTreeExpander("Backplane");
				
				// Click on Panel Accessories label
				repo.FormMe.PanelAccessoriesLabel.Click();
				
				//  Verify label in Search properties
				Devices_Functions.VerifyLabelInSearchProperties(sLabel);
				
				// Verify SKU in Search properties
				Devices_Functions.VerifySKUInSearchProperties(sSKU);
				
				// Verify Model in Search properties
				Devices_Functions.VerifyModelInSearchProperties(sModel);
				
				// Verify Description field in Search properties
				Devices_Functions.VerifyDescriptionTextRowInSearchProperties();
				
				//  Verify label in Search properties
				Devices_Functions.editDeviceLabel("Label",sNewLabel);
				
				//Devices_Functions.VerifyLabelInPanelAccessories(sNewLabel);

			}
			// Close open excel file
			Excel_Utilities.CloseExcel();
		}
		
		/**************************************************************************************************************************************************
		 * Function Name: VerifyAdditionOfDevicesInBackplaneWithOnePanel
		 * Function Details: To verify slot cards and backplane distribution
		 * Parameter/Arguments: sFileName, sAddDevicesSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 06/06/2019 Alpesh Dhakad - 30/07/2019 - Updated script as per new build updated xpath
		 * Alpesh Dhakad - 16/08/2019,19/08/2019,28/08/2019,07/09/2019 & 16/10/2019 - Updated with new navigation tree method, xpath and devices gallery 
		 * 05/05/2020 - Alpesh Dhakad - Updated script as per new implementation
		 **************************************************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyAdditionOfDevicesInBackplaneWithOnePanel(string sFileName,string sAddDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			int columns = Excel_Utilities.ExcelRange.Columns.Count;
			int deviceCount;
			// Declared string type
			string sType,sDeviceCount,sDeviceName,PanelType,ExpectedBackplane1,ExpectedBackplane2,ExpectedBackplane3,PanelName,PanelNode,CPUType;
			string sBackplane1SlotCardName,sBackplane2SlotCardName,sBackplane3SlotCardName;

			for(int i=10; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				ExpectedBackplane1 = ((Range)Excel_Utilities.ExcelRange.Cells[i,11]).Value.ToString();
				ExpectedBackplane2 = ((Range)Excel_Utilities.ExcelRange.Cells[i,12]).Value.ToString();
				ExpectedBackplane3 = ((Range)Excel_Utilities.ExcelRange.Cells[i,13]).Value.ToString();
				sBackplane1SlotCardName = ((Range)Excel_Utilities.ExcelRange.Cells[i,14]).Value.ToString();
				sBackplane2SlotCardName = ((Range)Excel_Utilities.ExcelRange.Cells[i,15]).Value.ToString();
				sBackplane3SlotCardName = ((Range)Excel_Utilities.ExcelRange.Cells[i,16]).Value.ToString();
				
				for(int j=4; j<=10; j++){
					
					
					sDeviceName =  ((Range)Excel_Utilities.ExcelRange.Cells[8,j]).Value.ToString();
					sType = ((Range)Excel_Utilities.ExcelRange.Cells[9,j]).Value.ToString();
					sDeviceCount = ((Range)Excel_Utilities.ExcelRange.Cells[i,j]).Value.ToString();
					
					
					PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[4,7]).Value.ToString();
					
					int.TryParse(sDeviceCount, out deviceCount);
					
					// Verify device count and then add devices from panel accessories gallery or panel node gallery
					if(deviceCount>0)
					{
						if (sType.Equals("Accessories"))
						{
							// Click on Panel node
							Common_Functions.ClickOnNavigationTreeItem("Node");
							
							//repo.ProfileConsys1.NavigationTree.Expander.Click();
							Common_Functions.clickOnPanelAccessoriesTab();
							for(int k=1; k<=deviceCount;k++)
							{
								//Devices_Functions.AddDevicefromPanelAccessoriesGallery(sDeviceName,sType);
								Devices_Functions.AddDevicesfromGallery(sDeviceName,sType);
							}
						}
						else
						{
							// Click on Panel node
							Common_Functions.ClickOnNavigationTreeItem("Node");
							
							Common_Functions.clickOnInventoryTab();
							
							for(int k=1; k<=deviceCount;k++)
							{
								//Devices_Functions.AddDevicesfromPanelNodeGallery(sDeviceName,sType,PanelType);
								Devices_Functions.AddDevicesfromGallery(sDeviceName,sType);
								
								Thread.Sleep(3000);
							}
						}
					}
					
				}
				
				// Verify expected backplane1
				if(ExpectedBackplane1.Equals("Yes"))
				{
					//Common_Functions.ClickOnNavigationTreeExpander("Backplane");
					//Common_Functions.VerifyNavigationTreeItemText("Backplane");
					Common_Functions.VerifyAndClickNavigationTreeItemText(sBackplane1SlotCardName);
					//Common_Functions.ClickOnNavigationTreeExpander("Backplane");
					
				}
				else
				{
					Common_Functions.VerifyNavigationTreeItem("Backplane  1/3",ExpectedBackplane1);
					
				}
				
				// Verify expected backplane2
				if(ExpectedBackplane2.Equals("Yes"))
				{
					Common_Functions.ClickOnNavigationTreeExpander("Backplane  2/3");
					Common_Functions.VerifyNavigationTreeItemText("Backplane  2/3");
					Common_Functions.VerifyAndClickNavigationTreeItemText(sBackplane2SlotCardName);
					Common_Functions.ClickOnNavigationTreeExpander("Backplane  2/3");
						
				
					
				}
				else
				{
					Common_Functions.VerifyNavigationTreeItem("Backplane  2/3",ExpectedBackplane2);
						
			
					
				}
				
//				// Verify expected backplane3
//				if(ExpectedBackplane3.Equals("Yes"))
//				{
//					Common_Functions.ClickOnNavigationTreeExpander("Backplane  3/3");
//					Common_Functions.VerifyNavigationTreeItemText("Backplane  3/3");
//					Common_Functions.VerifyAndClickNavigationTreeItemText(sBackplane3SlotCardName);
//					Common_Functions.ClickOnNavigationTreeExpander("Backplane  3/3");
//						
//				}
//				else
//				{
//					Common_Functions.VerifyNavigationTreeItem("Backplane  3/3", ExpectedBackplane3);
//					
//					
//				}
				
			}
			
			// Close Excel
			Excel_Utilities.CloseExcel();
		}
		
		/************************************************************************************************************
		 * Function Name: VerifyShoppingListDevicesAfterAdditionOfDevices
		 * Function Details: To verify slot cards and backplane distribution
		 * Parameter/Arguments: sFileName, sAddDevicesSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 06/06/2019 Alpesh Dhakad - 30/07/2019 - Updated script as per new build updated xpath
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyShoppingListDevicesAfterAdditionOfDevices(string sFileName,string sAddDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			string FirstDevice,SecondDevice,ThirdDevice,FourthDevice,sShoppingListCount;
			int shoppingListCount;
			
			
			sShoppingListCount =  ((Range)Excel_Utilities.ExcelRange.Cells[4,8]).Value.ToString();
			FirstDevice =  ((Range)Excel_Utilities.ExcelRange.Cells[4,9]).Value.ToString();
			SecondDevice = ((Range)Excel_Utilities.ExcelRange.Cells[4,10]).Value.ToString();
			ThirdDevice = ((Range)Excel_Utilities.ExcelRange.Cells[4,11]).Value.ToString();
			FourthDevice = ((Range)Excel_Utilities.ExcelRange.Cells[4,12]).Value.ToString();
			
			
			int.TryParse(sShoppingListCount, out shoppingListCount);
			
			// Click on Site node
			Common_Functions.ClickOnNavigationTreeItem("Site");
			
			// Click on shopping list tab
			Common_Functions.clickOnShoppingListTab();
			
			
			Export_Functions.SearchDeviceInExportUsingSKUOrDescription(FirstDevice,true);
			Delay.Milliseconds(100);
			
			Export_Functions.SearchDeviceInExportUsingSKUOrDescription(SecondDevice,true);
			Delay.Milliseconds(100);
			
			Export_Functions.SearchDeviceInExportUsingSKUOrDescription(ThirdDevice,true);
			Delay.Milliseconds(100);
			
			//Export_Functions.SearchDeviceInExportUsingSKUOrDescription(FourthDevice,true);
			//Delay.Milliseconds(100);
			
		}
		
		
		/*************************************************************************************************************************
		 * Function Name: VerifyDeletionOfDevicesInBackplane
		 * Function Details: To Verify
		 * Parameter/Arguments: string sFileName,string sAddDevicesSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 06/06/2019  Alpesh Dhakad - 30/07/2019, 07/09/2019 - Updated script as per new build updated xpath
		 * 05/05/2020 - Alpesh Dhakad - Updated script as per new implementation
		 *************************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyDeletionOfDevicesInBackplane(string sFileName,string sAddDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			int columns = Excel_Utilities.ExcelRange.Columns.Count;
			
			// Declared string type
			string ThirdDeviceLabel,FirstDevice,SecondDevice,FourthDevice,sShoppingListCount,sBackplane1SlotCardName,ExpectedBackplane1;
			int shoppingListCount;
			
			sShoppingListCount =  ((Range)Excel_Utilities.ExcelRange.Cells[4,8]).Value.ToString();
			ThirdDeviceLabel =  ((Range)Excel_Utilities.ExcelRange.Cells[4,9]).Value.ToString();
			
			
			FirstDevice = ((Range)Excel_Utilities.ExcelRange.Cells[4,10]).Value.ToString();
			SecondDevice = ((Range)Excel_Utilities.ExcelRange.Cells[4,11]).Value.ToString();
			FourthDevice = ((Range)Excel_Utilities.ExcelRange.Cells[4,12]).Value.ToString();
			sBackplane1SlotCardName = ((Range)Excel_Utilities.ExcelRange.Cells[10,14]).Value.ToString();
			ExpectedBackplane1 = ((Range)Excel_Utilities.ExcelRange.Cells[10,11]).Value.ToString();
			
			int.TryParse(sShoppingListCount, out shoppingListCount);
			
			//Devices_Functions.SelectRowUsingLabelName(ThirdDeviceLabel);
			Devices_Functions.SelectRowUsingLabelNameFromInventoryTab(ThirdDeviceLabel);
			
			Common_Functions.clickOnDeleteButton();
			
			// Verify expected backplane1
				if(ExpectedBackplane1.Equals("Yes"))
				{
					//Common_Functions.ClickOnNavigationTreeExpander("Backplane");
					Common_Functions.VerifyAndClickNavigationTreeItemText(sBackplane1SlotCardName);
				//	Common_Functions.ClickOnNavigationTreeExpander("Backplane");
					
				}
				else
				{
					Common_Functions.VerifyNavigationTreeItem("Backplane",ExpectedBackplane1);
					
				}
			
			
			// Click on Site node
			Common_Functions.ClickOnNavigationTreeItem("Site");
			
			
			// Click on shopping list tab
			Common_Functions.clickOnShoppingListTab();
			
			// Verify shopping list count
			Devices_Functions.verifyShoppingList(shoppingListCount);
			Delay.Milliseconds(500);
			
			Export_Functions.SearchDeviceInExportUsingSKUOrDescription(FirstDevice,true);
			
			Export_Functions.SearchDeviceInExportUsingSKUOrDescription(SecondDevice,true);
			
			//Export_Functions.SearchDeviceInExportUsingSKUOrDescription(FourthDevice,true);
			
			
			// Close Excel
			Excel_Utilities.CloseExcel();
			
			
		}
		
		/****************************************************************************************************************************************
		 * Function Name: VerifyAccessoriesGalleryUpdateOnMaxLimitSupportedByPanel
		 * Function Details: To verify slot cards and backplane distribution
		 * Parameter/Arguments: sFileName, sAddDevicesSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 07/06/2019 Alpesh Dhakad - 30/07/2019,20/08/2019 & 07/09/2019 - Updated script as per new build updated xpath
		 * 08/05/2020 - Alpesh Dhakad - Updated script as per new implementation
		 ****************************************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyAccessoriesGalleryUpdateOnMaxLimitSupportedByPanel(string sFileName,string sAddDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			int columns = Excel_Utilities.ExcelRange.Columns.Count;
			int deviceCount;
			// Declared string type
			string sType,sDeviceCount,sDeviceName,PanelType,ExpectedBackplane1,ExpectedBackplane2,ExpectedBackplane3,PanelName,PanelNode,CPUType;
			string sBackplane1SlotCardName,sBackplane2SlotCardName,sBackplane3SlotCardName,sDeviceState;

			for(int i=10; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				ExpectedBackplane1 = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				ExpectedBackplane2 = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				ExpectedBackplane3 = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				sBackplane1SlotCardName = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				sBackplane2SlotCardName = ((Range)Excel_Utilities.ExcelRange.Cells[i,10]).Value.ToString();
				sBackplane3SlotCardName = ((Range)Excel_Utilities.ExcelRange.Cells[i,11]).Value.ToString();
				sDeviceState = ((Range)Excel_Utilities.ExcelRange.Cells[i,12]).Value.ToString();
				
				
				for(int j=4; j<=5; j++){
					
					
					sDeviceName =  ((Range)Excel_Utilities.ExcelRange.Cells[8,j]).Value.ToString();
					sType = ((Range)Excel_Utilities.ExcelRange.Cells[9,j]).Value.ToString();
					sDeviceCount = ((Range)Excel_Utilities.ExcelRange.Cells[i,j]).Value.ToString();
					
					
					PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[4,7]).Value.ToString();
					
					int.TryParse(sDeviceCount, out deviceCount);
					
					// Verify device count and then add devices from panel accessories gallery or panel node gallery
					if(deviceCount>0)
					{
						if (sType.Equals("Accessories"))
						{
							// Click on Panel node
							Common_Functions.ClickOnNavigationTreeItem("Node");
							
						
							Common_Functions.clickOnPanelAccessoriesTab();
							for(int k=1; k<=deviceCount;k++)
							{
								Devices_Functions.AddDevicefromPanelAccessoriesGallery(sDeviceName,sType);
							}
						}
						else
						{
							// Click on Panel node
							Common_Functions.ClickOnNavigationTreeItem("Node");
							
							Common_Functions.clickOnInventoryTab();
							
							for(int k=1; k<=deviceCount;k++)
							{
								Devices_Functions.AddDevicesfromPanelNodeGallery(sDeviceName,sType,PanelType);
							}
						}
					}
					Devices_Functions.VerifyEnableDisablePanelAccessoriesGallery(sType,sDeviceName,sDeviceState);
					
					//Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
				}
				
				// Verify expected backplane1
				if(ExpectedBackplane1.Equals("Yes"))
				{
					//Common_Functions.ClickOnNavigationTreeExpander("Backplane");
					//Common_Functions.VerifyNavigationTreeItemText("Backplane");
					Common_Functions.VerifyAndClickNavigationTreeItemText(sBackplane1SlotCardName);
					//Common_Functions.ClickOnNavigationTreeExpander("Backplane");
					
				}
				else
				{
					Common_Functions.VerifyNavigationTreeItem("Backplane  1/3",ExpectedBackplane1);
					
				}
				
				// Verify expected backplane2
				if(ExpectedBackplane2.Equals("Yes"))
				{
					Common_Functions.ClickOnNavigationTreeExpander("Backplane  2/3");
					//Common_Functions.VerifyNavigationTreeItemText("Backplane  2/3");
					Common_Functions.VerifyAndClickNavigationTreeItemText(sBackplane2SlotCardName);
					//Common_Functions.ClickOnNavigationTreeExpander("Backplane  2/3");
						
				
					
				}
				else
				{
					Common_Functions.VerifyNavigationTreeItem("Backplane  2/3",ExpectedBackplane2);
						
			
					
				}
				
//				// Verify expected backplane3
//				if(ExpectedBackplane3.Equals("Yes"))
//				{
//					Common_Functions.ClickOnNavigationTreeExpander("Backplane  3/3");
//					Common_Functions.VerifyNavigationTreeItemText("Backplane  3/3");
//					Common_Functions.VerifyAndClickNavigationTreeItemText(sBackplane3SlotCardName);
//					Common_Functions.ClickOnNavigationTreeExpander("Backplane  3/3");
//						
//				}
//				else
//				{
//					Common_Functions.VerifyNavigationTreeItem("Backplane  3/3", ExpectedBackplane3);
//					
//					
//				}
				
			}

			// Close Excel
			Excel_Utilities.CloseExcel();
		}
		
		/*************************************************************************************************************************************************
		 * Function Name: VerifyMaxLimitForRepeatersSupportedByPanelOnEthernet
		 * Function Details: To verify max limit for repeaters supported by panel on Ethernet
		 * Parameter/Arguments: sFileName, sAddDevicesSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 10/06/2019 Alpesh Dhakad- 29/07/2019 - Updated script and add AddDevicesfromEthernetGallery as per new build xpath updates
		 * Alpesh Dhakad - 16/08/2019 - Updated with new navigation tree method, xpath and devices gallery 
		 ***************************************************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyMaxLimitForRepeatersSupportedByPanelOnEthernet(string sFileName,string sAddDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			int columns = Excel_Utilities.ExcelRange.Columns.Count;
			int deviceCount;
			// Declared string type
			string PanelName,PanelNode,CPUType,sType,sDeviceCount,sDeviceName,PanelType,sFirstDeviceName,sFirstDeviceState,sSecondDeviceName,sSecondDeviceState;
			
			for(int i=10; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[9,4]).Value.ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[4,7]).Value.ToString();
				
				sFirstDeviceName = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				sFirstDeviceState = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				sSecondDeviceName = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				sSecondDeviceState = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				
				// Add panels
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				
				// Click on Expander node
				Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
				
				// Click on Main processor expander
				//Common_Functions.ClickOnNavigationTreeExpander("Main");
				
				// Click on Ethernet node
				Common_Functions.ClickOnNavigationTreeItem("Ethernet");
				
				for(int j=4; j<=5; j++)
				{
					
					sDeviceName =  ((Range)Excel_Utilities.ExcelRange.Cells[8,j]).Value.ToString();
					sType = ((Range)Excel_Utilities.ExcelRange.Cells[9,j]).Value.ToString();
					sDeviceCount = ((Range)Excel_Utilities.ExcelRange.Cells[i,j]).Value.ToString();
					
					int.TryParse(sDeviceCount, out deviceCount);
					
					// Verify device count and then add devices from panel accessories gallery or panel node gallery
					if(deviceCount>0)
					{
						
						// Click on Ethernet node
						Common_Functions.ClickOnNavigationTreeItem("Ethernet");
						
						// Click on Inventory tab
						Common_Functions.clickOnInventoryTab();
						
						for(int k=1; k<=deviceCount;k++)
						{
							// Add Devices from main processor gallery
							//Devices_Functions.AddDevicesfromMainProcessorGallery(sDeviceName,sType,PanelType);
							Devices_Functions.AddDevicesfromEthernetGallery(sDeviceName,sType,PanelType);
							
						}
					}
					
				}
				// Verify Node gallery devices state for first device
				Devices_Functions.VerifyNodeGallery(sFirstDeviceName,sType,sFirstDeviceState,PanelType);
				
				// Verify Node gallery devices state for second device
				Devices_Functions.VerifyNodeGallery(sSecondDeviceName,sType,sSecondDeviceState,PanelType);
				
				// Click on Panel Node
				Common_Functions.ClickOnNavigationTreeItem(PanelNode);
				
				// Verify if row count is more than 8 then delete the panel
				if(rows!=10)
				{
					// Delete Panel
					Panel_Functions.DeletePanel(1,PanelNode,1);
				}
				
			}
			
			// Close Excel
			Excel_Utilities.CloseExcel();
		}
		
		
		/*********************************************************************************************************
		 * Function Name: VerifyMaxLimitForRepeatersSupportedByPanelOnRBus
		 * Function Details: To verify max limit for repeaters supported by panel on RBUS
		 * Parameter/Arguments: sFileName, sAddDevicesSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 11/06/2019 Alpesh Dhakad- 29/07/2019 - Updated script as per new build xpath updates
		 * Alpesh Dhakad - 16/08/2019 - Updated with new navigation tree method, xpath and devices gallery 
		 ********************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyMaxLimitForRepeatersSupportedByPanelOnRBus(string sFileName,string sAddDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			int columns = Excel_Utilities.ExcelRange.Columns.Count;
			int deviceCount;
			// Declared string type
			string PanelName,PanelNode,CPUType,sType,sDeviceCount,sDeviceName,PanelType,sFirstDeviceName,sFirstDeviceState,sSecondDeviceName,sSecondDeviceState;
			
			for(int i=10; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[9,4]).Value.ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[4,7]).Value.ToString();
				
				sFirstDeviceName = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				sFirstDeviceState = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				sSecondDeviceName = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				sSecondDeviceState = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				
				// Add panels
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				
				// Click on Expander node
				Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
				
				// Click on Main processor expander
				//Common_Functions.ClickOnNavigationTreeExpander("Main");
				
				// Click on R-BUS node
				Common_Functions.ClickOnNavigationTreeItem("R-BUS");
				
				for(int j=4; j<=5; j++)
				{
					
					sDeviceName =  ((Range)Excel_Utilities.ExcelRange.Cells[8,j]).Value.ToString();
					sType = ((Range)Excel_Utilities.ExcelRange.Cells[9,j]).Value.ToString();
					sDeviceCount = ((Range)Excel_Utilities.ExcelRange.Cells[i,j]).Value.ToString();
					
					int.TryParse(sDeviceCount, out deviceCount);
					
					// Verify device count and then add devices from panel accessories gallery or panel node gallery
					if(deviceCount>0)
					{
						
						// Click on R-BUS node
						Common_Functions.ClickOnNavigationTreeItem("R-BUS");
				
						// Click on Inventory tab
						Common_Functions.clickOnInventoryTab();
						
						for(int k=1; k<=deviceCount;k++)
						{
							// Add Devices from main processor gallery
							Devices_Functions.AddDevicesfromMainProcessorGallery(sDeviceName,sType,PanelType);
						}
					}
					
				}
				// Verify Node gallery devices state for first device
				Devices_Functions.VerifyNodeGallery(sFirstDeviceName,sType,sFirstDeviceState,PanelType);
				
				// Verify Node gallery devices state for second device
				Devices_Functions.VerifyNodeGallery(sSecondDeviceName,sType,sSecondDeviceState,PanelType);
				
					// Click on Panel Node
				Common_Functions.ClickOnNavigationTreeItem(PanelNode);
				
				// Verify if row count is more than 8 then delete the panel
				if(rows!=10)
				{
					// Delete Panel
					Panel_Functions.DeletePanel(1,PanelNode,1);
				}
				
			}
			
			// Close Excel
			Excel_Utilities.CloseExcel();
		}
		
		/***********************************************************************************************************************
		 * Function Name: VerifyAdditionOfDevicesFromOtherSlotCards
		 * Function Details: To verify slot cards
		 * Parameter/Arguments: sFileName, sAddDevicesSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 13/06/2019 Alpesh Dhakad- 29/07/2019 & 20-21/08/2019- Updated script as per new build xpath updates
		 * 04/05/2020 - Updated script as per new UI framework changes
		 ***********************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyAdditionOfDevicesFromOtherSlotCards(string sFileName,string sAddDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			int columns = Excel_Utilities.ExcelRange.Columns.Count;
			int deviceCount;
			// Declared string type
			string sType,sDeviceCount,sDeviceName,PanelType,ExpectedBackplane1,ExpectedBackplane2,ExpectedBackplane3,PanelName,PanelNode,CPUType;
			string sBackplane1SlotCardName,sBackplane2SlotCardName,sBackplane3SlotCardName;

			for(int i=10; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				sOtherSlotCardName = ((Range)Excel_Utilities.ExcelRange.Cells[i,10]).Value.ToString();
				ExpectedBackplane1 = ((Range)Excel_Utilities.ExcelRange.Cells[i,11]).Value.ToString();
				ExpectedBackplane2 = ((Range)Excel_Utilities.ExcelRange.Cells[i,12]).Value.ToString();
				ExpectedBackplane3 = ((Range)Excel_Utilities.ExcelRange.Cells[i,13]).Value.ToString();
				sBackplane1SlotCardName = ((Range)Excel_Utilities.ExcelRange.Cells[i,14]).Value.ToString();
				sBackplane2SlotCardName = ((Range)Excel_Utilities.ExcelRange.Cells[i,15]).Value.ToString();
				sBackplane3SlotCardName = ((Range)Excel_Utilities.ExcelRange.Cells[i,16]).Value.ToString();
				
				for(int j=4; j<=9; j++){
					
					
					sDeviceName =  ((Range)Excel_Utilities.ExcelRange.Cells[8,j]).Value.ToString();
					sType = ((Range)Excel_Utilities.ExcelRange.Cells[9,j]).Value.ToString();
					sDeviceCount = ((Range)Excel_Utilities.ExcelRange.Cells[i,j]).Value.ToString();
					
					
					PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[4,7]).Value.ToString();
					
					int.TryParse(sDeviceCount, out deviceCount);
					
					// Verify device count and then add devices from panel accessories gallery or panel node gallery
					if(deviceCount>0)
					{
						if (sType.Equals("Accessories"))
						{
							// Click on Panel node
							Common_Functions.ClickOnNavigationTreeItem(PanelNode);
				
							Common_Functions.clickOnPanelAccessoriesTab();
							for(int k=1; k<=deviceCount;k++)
							{
								Devices_Functions.AddDevicefromPanelAccessoriesGallery(sDeviceName,sType);
							}
						}
						else
						{
							// Click on Panel node
							Common_Functions.ClickOnNavigationTreeItem(PanelNode);
				
							// Click on Panel node
							Common_Functions.ClickOnNavigationTreeItem("Node1");
				
							
							// Click on Expander node
							//Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
				
							// Click on Expander node
							//Common_Functions.ClickOnNavigationTreeExpander("Backplane");
				
							// Click on Other slot cards tree item
							//Common_Functions.ClickOnNavigationTreeItem("Other Slot Cards");
							
							// Click on Expander node
							//Common_Functions.ClickOnNavigationTreeExpander("Backplane");
				
							Common_Functions.clickOnInventoryTab();
							
							for(int k=1; k<=deviceCount;k++)
							{
								Devices_Functions.AddDevicesfromPanelNodeGallery(sDeviceName,sType,PanelType);
							}
						}
					}
					
				}
				
				
				// Verify expected backplane1
				if(ExpectedBackplane1.Equals("Yes"))
				{
					//Common_Functions.ClickOnNavigationTreeExpander("Backplane  1/3");
					//Common_Functions.VerifyNavigationTreeItemText("Backplane  1/3");
					Common_Functions.VerifyAndClickNavigationTreeItemText(sBackplane1SlotCardName);
					//Common_Functions.ClickOnNavigationTreeExpander("Backplane  1/3");
					
				}
				else
				{
					Common_Functions.VerifyNavigationTreeItem("Backplane", ExpectedBackplane1);
				}
				
//				// Verify expected backplane2
//				if(ExpectedBackplane2.Equals("Yes"))
//				{
//					Common_Functions.ClickOnNavigationTreeExpander("Backplane  2/3");
//					Common_Functions.VerifyNavigationTreeItemText("Backplane  2/3");
//					Common_Functions.VerifyAndClickNavigationTreeItemText(sBackplane2SlotCardName);
//					Common_Functions.ClickOnNavigationTreeExpander("Backplane  2/3");
//				}
//				else
//				{
//					Common_Functions.VerifyNavigationTreeItem("Backplane  2/3", ExpectedBackplane2);
//					
//					
//				}
				
//				// Verify expected backplane3
//				if(ExpectedBackplane3.Equals("Yes"))
//				{
//					Common_Functions.ClickOnNavigationTreeExpander("Backplane  3/3");
//					Common_Functions.VerifyNavigationTreeItemText("Backplane  3/3");
//					Common_Functions.VerifyAndClickNavigationTreeItemText(sBackplane3SlotCardName);
//					Common_Functions.ClickOnNavigationTreeExpander("Backplane  3/3");
//						
//				}
//				else
//				{
//					Common_Functions.VerifyNavigationTreeItem("Backplane  3/3", ExpectedBackplane3);
//					
//				}
			}
			
			// Close Excel
			Excel_Utilities.CloseExcel();
		}
		
		/***********************************************************************************************************
		 * Function Name: VerifyDeletionOfDevicesFromOtherSlotCards
		 * Function Details: To Verify deletion Of devices From OtherSlotCards
		 * Parameter/Arguments: string sFileName,string sAddDevicesSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 13/06/2019 Alpesh Dhakad- 30/08/2019- Updated script as per new build xpath updates
		 * 04/05/2020 Alpesh Dhakad - Updated script as per new framework
		 ***********************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyDeletionOfDevicesFromOtherSlotCards(string sFileName,string sAddDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			int columns = Excel_Utilities.ExcelRange.Columns.Count;
			
			// Declared string type
			string ExpectedBackplane1,ExpectedBackplane2,ExpectedBackplane3,PanelName,PanelNode,CPUType;
			string sBackplane1SlotCardName,sBackplane2SlotCardName,sBackplane3SlotCardName;
			
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sOtherSlotCardName = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				ExpectedBackplane1 = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				ExpectedBackplane2 = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				ExpectedBackplane3 = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				sBackplane1SlotCardName = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				sBackplane2SlotCardName = ((Range)Excel_Utilities.ExcelRange.Cells[i,10]).Value.ToString();
				sBackplane3SlotCardName = ((Range)Excel_Utilities.ExcelRange.Cells[i,11]).Value.ToString();
				
				// Verify expected backplane1
				if(ExpectedBackplane1.Equals("Yes"))
				{
					//Common_Functions.ClickOnNavigationTreeExpander("Backplane  1/3");
					Common_Functions.VerifyNavigationTreeItemText("Backplane  1/3");
					//Common_Functions.VerifyAndClickNavigationTreeItemText(sBackplane1SlotCardName);
						
						Devices_Functions.SelectRowUsingLabelNameFromInventoryTab(sLabelName);
						Common_Functions.clickOnDeleteButton();
						
						Report.Log(ReportLevel.Success, "Device with " +sLabelName+ " deleted successfully");
					
					//Common_Functions.ClickOnNavigationTreeExpander("Backplane  1/3");
				}
				else
				{
					Common_Functions.VerifyNavigationTreeItem("Backplane", ExpectedBackplane1);
					
				}
				
//				// Verify expected backplane2
//				if(ExpectedBackplane2.Equals("Yes"))
//				{
//					Common_Functions.ClickOnNavigationTreeExpander("Backplane  2/3");
//					Common_Functions.VerifyNavigationTreeItemText("Backplane  2/3");
//					Common_Functions.VerifyAndClickNavigationTreeItemText(sBackplane2SlotCardName);
//						
//						Devices_Functions.SelectRowUsingLabelName(sLabelName);
//						Common_Functions.clickOnDeleteButton();
//						
//						Report.Log(ReportLevel.Success, "Device with " +sLabelName+ " deleted successfully");
//				
//						Common_Functions.ClickOnNavigationTreeExpander("Backplane  2/3");
//				}
//				else
//				{
//					Common_Functions.VerifyNavigationTreeItem("Backplane  2/3", ExpectedBackplane2);
//					
//				}				
//				
//				// Verify expected backplane3
//				if(ExpectedBackplane3.Equals("Yes"))
//				{
//					Common_Functions.ClickOnNavigationTreeExpander("Backplane  3/3");
//					Common_Functions.VerifyNavigationTreeItemText("Backplane  3/3");
//					Common_Functions.VerifyAndClickNavigationTreeItemText(sBackplane3SlotCardName);
//						
//						Devices_Functions.SelectRowUsingLabelName(sLabelName);
//						Common_Functions.clickOnDeleteButton();
//						
//						Report.Log(ReportLevel.Success, "Device with " +sLabelName+ " deleted successfully");
//						
//						Common_Functions.ClickOnNavigationTreeExpander("Backplane  3/3");
//				}
//				else
//				{
//					Common_Functions.VerifyNavigationTreeItem("Backplane  3/3", ExpectedBackplane3);
//					
//				}
//				
				
			}
			
			// Close Excel
			Excel_Utilities.CloseExcel();

		}

		/********************************************************************************************************************************************
		 * Function Name: VerifyOtherSlotCardGrid
		 * Function Details: To verify other slot cards grid
		 * Parameter/Arguments: sFileName, sAddDevicesSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : : 14/06/2019 Alpesh Dhakad- 30/07/2019,30/08/2019,06/09/2019 & 10/09/2019 - Updated script as per new build xpath updates
		 * 06/05/2020 - Alpesh Dhakad - Updated script as per new implemenation
		 ********************************************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyOtherSlotCardGrid(string sFileName,string sAddDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			int columns = Excel_Utilities.ExcelRange.Columns.Count;
			int deviceCount;
			// Declared string type
			string sType,sDeviceCount,sDeviceName,PanelType,ExpectedBackplane1,ExpectedBackplane2,ExpectedBackplane3,PanelName,PanelNode,CPUType;
			string sBackplane1SlotCardName,sBackplane2SlotCardName,sBackplane3SlotCardName,sPLXLoopCardBackplane1;

			for(int i=10; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				ExpectedBackplane1 = ((Range)Excel_Utilities.ExcelRange.Cells[i,11]).Value.ToString();
				ExpectedBackplane2 = ((Range)Excel_Utilities.ExcelRange.Cells[i,12]).Value.ToString();
				ExpectedBackplane3 = ((Range)Excel_Utilities.ExcelRange.Cells[i,13]).Value.ToString();
				sBackplane1SlotCardName = ((Range)Excel_Utilities.ExcelRange.Cells[i,14]).Value.ToString();
				sBackplane2SlotCardName = ((Range)Excel_Utilities.ExcelRange.Cells[i,15]).Value.ToString();
				sBackplane3SlotCardName = ((Range)Excel_Utilities.ExcelRange.Cells[i,16]).Value.ToString();
				sPLXLoopCardBackplane1 = ((Range)Excel_Utilities.ExcelRange.Cells[i,17]).Value.ToString();
				
				for(int j=4; j<=10; j++){
					
					
					sDeviceName =  ((Range)Excel_Utilities.ExcelRange.Cells[8,j]).Value.ToString();
					sType = ((Range)Excel_Utilities.ExcelRange.Cells[9,j]).Value.ToString();
					sDeviceCount = ((Range)Excel_Utilities.ExcelRange.Cells[i,j]).Value.ToString();
					
					
					PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[4,7]).Value.ToString();
					
					int.TryParse(sDeviceCount, out deviceCount);
					
					// Verify device count and then add devices from panel accessories gallery or panel node gallery
					if(deviceCount>0)
					{
						if (sType.Equals("Accessories"))
						{
							// Click on Panel Node
							Common_Functions.ClickOnNavigationTreeItem(PanelNode);
					
							Common_Functions.clickOnPanelAccessoriesTab();
							for(int k=1; k<=deviceCount;k++)
							{
								Devices_Functions.AddDevicefromPanelAccessoriesGallery(sDeviceName,sType);
							}
						}
						else
						{
							// Click on Panel Node
							Common_Functions.ClickOnNavigationTreeItem(PanelNode);
					
							// Click on Panel Node
							Common_Functions.ClickOnNavigationTreeItem("Node");
					
							
							Common_Functions.clickOnInventoryTab();
							
							for(int k=1; k<=deviceCount;k++)
							{
								Devices_Functions.AddDevicesfromPanelNodeGallery(sDeviceName,sType,PanelType);
							}
						}
					}
					
				}
				
					// Verify expected backplane1
				if(ExpectedBackplane1.Equals("Yes"))
				{
					//Common_Functions.ClickOnNavigationTreeExpander("Backplane  1/1");
					Common_Functions.VerifyNavigationTreeItemText("Backplane  1/1");
					Common_Functions.VerifyAndClickNavigationTreeItemText(sBackplane1SlotCardName);
					Common_Functions.VerifyAndClickNavigationTreeItemText(sPLXLoopCardBackplane1);
					//Common_Functions.ClickOnNavigationTreeExpander("Backplane  1/1");
					
				}
				else
				{
					Common_Functions.VerifyNavigationTreeItem("Backplane  1/1", ExpectedBackplane1);
				}
				
				
			}
			
			// Close Excel
			Excel_Utilities.CloseExcel();
			
		}
		
		/***********************************************************************************************************************************************
		 * Function Name: VerifyAddUnitDetails
		 * Function Details: Verify Add Unit Details and its status
		 * Parameter/Arguments: sFileName, sAddDevicesSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 25/06/2019 Alpesh Dhakad - 29/07/2019 - Updated scripts as per new build xpaths
		 * Alpesh Dhakad - 16/08/2019,19/08/2019,05/09/2019 & 06/09/2019 - Updated with new navigation tree method, xpath and devices gallery 
		 * Alpesh Dhakad - 09/05/2020 Updated as per new internal structure and implemenation changes
		 ***********************************************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyAddUnitDetails(string sFileName,string sAddDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			int columns = Excel_Utilities.ExcelRange.Columns.Count;
			int deviceCount;
			// Declared string type
			string sType,sDeviceCount,sDeviceName,PanelType,ExpectedBackplane1,ExpectedBackplane2,ExpectedBackplane3,PanelName,PanelNode,CPUType;
			string sBackplane1SlotCardName,sBackplane2SlotCardName,sBackplane3SlotCardName,sDeviceState;
			
			for(int i=10; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				ExpectedBackplane1 = ((Range)Excel_Utilities.ExcelRange.Cells[i,12]).Value.ToString();
				ExpectedBackplane2 = ((Range)Excel_Utilities.ExcelRange.Cells[i,13]).Value.ToString();
				ExpectedBackplane3 = ((Range)Excel_Utilities.ExcelRange.Cells[i,14]).Value.ToString();
				sBackplane1SlotCardName = ((Range)Excel_Utilities.ExcelRange.Cells[i,15]).Value.ToString();
				sBackplane2SlotCardName = ((Range)Excel_Utilities.ExcelRange.Cells[i,16]).Value.ToString();
				sBackplane3SlotCardName = ((Range)Excel_Utilities.ExcelRange.Cells[i,17]).Value.ToString();
				sDeviceState = ((Range)Excel_Utilities.ExcelRange.Cells[i,18]).Value.ToString();
				
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				
				// Click on Expander node
					Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
				
				
				for(int j=4; j<=11; j++){
					
					
					sDeviceName =  ((Range)Excel_Utilities.ExcelRange.Cells[8,j]).Value.ToString();
					sType = ((Range)Excel_Utilities.ExcelRange.Cells[9,j]).Value.ToString();
					sDeviceCount = ((Range)Excel_Utilities.ExcelRange.Cells[i,j]).Value.ToString();
					
					
					PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[4,7]).Value.ToString();
					
					int.TryParse(sDeviceCount, out deviceCount);
					
					// Click on Expander node
					//Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
				
					// Verify device count and then add devices from panel accessories gallery or panel node gallery
					if(deviceCount>0)
					{
						if (sType.Equals("Accessories"))
						{
							// Click on Panel node
							Common_Functions.ClickOnNavigationTreeItem(PanelNode);
				
							Common_Functions.clickOnPanelAccessoriesTab();
							
							for(int k=1; k<=deviceCount;k++)
							{
								//Devices_Functions.AddDevicefromPanelAccessoriesGallery(sDeviceName,sType);
								Devices_Functions.AddDevicesfromGallery(sDeviceName,sType);
							}
						}
						else
						{
							// Click on Panel node
							Common_Functions.ClickOnNavigationTreeItem(PanelNode);
				
							Common_Functions.clickOnInventoryTab();
							
							for(int k=1; k<=deviceCount;k++)
							{
								//Devices_Functions.AddDevicesfromPanelNodeGallery(sDeviceName,sType,PanelType);
								Devices_Functions.AddDevicesfromGallery(sDeviceName,sType);
							}
						}
						
//						if(sType.Equals("Accessories"))
//						{
//							//Devices_Functions.VerifyPanelNodePanelAccessoriesGallery(sDeviceName,sType,sDeviceState);
//							Devices_Functions.AddDevicesfromGallery(sDeviceName,sType);
//						}
//						else
//						{
//							Devices_Functions.VerifyNodeGallery(sDeviceName,sType,sDeviceState,PanelType);
//						}
					}
				}
				
				// Click on Expander node
				//	Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
					
				// Verify expected backplane1
				if(ExpectedBackplane1.Equals("Yes"))
				{
					//Common_Functions.ClickOnNavigationTreeExpander("Backplane");
					//Common_Functions.VerifyNavigationTreeItemText("Backplane");
					Common_Functions.VerifyAndClickNavigationTreeItemText(sBackplane1SlotCardName);
					//Common_Functions.ClickOnNavigationTreeExpander("Backplane");
					
				}
				else
				{
					Common_Functions.VerifyNavigationTreeItem("Backplane", ExpectedBackplane1);
				}
				
				// Verify expected backplane2
				if(ExpectedBackplane2.Equals("Yes"))
				{
					Common_Functions.ClickOnNavigationTreeExpander("Backplane  2/3");
					Common_Functions.VerifyNavigationTreeItemText("Backplane  2/3");
					Common_Functions.VerifyAndClickNavigationTreeItemText(sBackplane2SlotCardName);
					Common_Functions.ClickOnNavigationTreeExpander("Backplane  2/3");
				}
				else
				{
					Common_Functions.VerifyNavigationTreeItem("Backplane  2/3", ExpectedBackplane2);
					
					
				}
				
				// Verify expected backplane3
//				if(ExpectedBackplane3.Equals("Yes"))
//				{
//					Common_Functions.ClickOnNavigationTreeExpander("Backplane  3/3");
//					Common_Functions.VerifyNavigationTreeItemText("Backplane  3/3");
//					Common_Functions.VerifyAndClickNavigationTreeItemText(sBackplane3SlotCardName);
//					Common_Functions.ClickOnNavigationTreeExpander("Backplane  3/3");
//						
//				}
//				else
//				{
//					Common_Functions.VerifyNavigationTreeItem("Backplane  3/3", ExpectedBackplane3);
//					
//				}
				
			
				// Delete Panel
				Panel_Functions.DeletePanel(1,PanelNode,1);
				
				
				
			}
			
			// Close Excel
			Excel_Utilities.CloseExcel();
		}
		

		/********************************************************************************************************************
		 * Function Name: VerifySlotCardsAndLoopCardsDistribution
		 * Function Details: To verify slot cards and backplane distribution
		 * Parameter/Arguments: sFileName, sAddDevicesSheet
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 26/06/2019  Alpesh Dhakad - 29/07/2019 - Updated script as per new build xpaths
		 * Alpesh Dhakad - 16/08/2019 & 05/09/2019 - Updated with new navigation tree method, xpath and devices gallery 
		 ********************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifySlotCardsAndLoopCardsDistribution(string sFileName,string sAddDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			int columns = Excel_Utilities.ExcelRange.Columns.Count;
			int deviceCount,MaxInOtherSlotCard;
			// Declared string type
			string sType,sDeviceCount,sDeviceName,PanelType,ExpectedBackplane1,ExpectedBackplane2,ExpectedBackplane3,PanelName,PanelNode,CPUType;
			string sMaxInOtherSlotCard,sMaxLimitOfDevice;

			for(int i=10; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				ExpectedBackplane1 = ((Range)Excel_Utilities.ExcelRange.Cells[i,11]).Value.ToString();
				ExpectedBackplane2 = ((Range)Excel_Utilities.ExcelRange.Cells[i,12]).Value.ToString();
				ExpectedBackplane3 = ((Range)Excel_Utilities.ExcelRange.Cells[i,13]).Value.ToString();
				sMaxInOtherSlotCard = ((Range)Excel_Utilities.ExcelRange.Cells[i,14]).Value.ToString();
				sMaxLimitOfDevice = ((Range)Excel_Utilities.ExcelRange.Cells[i,15]).Value.ToString();
				
				
				// Add panels
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				
				// Click on Loop A node
							Common_Functions.ClickOnNavigationTreeItem(PanelNode);
				
							// Click on Expander node
							Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
						
						
				
				for(int j=4; j<=10; j++)
				{
					
					
					sDeviceName =  ((Range)Excel_Utilities.ExcelRange.Cells[8,j]).Value.ToString();
					sType = ((Range)Excel_Utilities.ExcelRange.Cells[9,j]).Value.ToString();
					sDeviceCount = ((Range)Excel_Utilities.ExcelRange.Cells[i,j]).Value.ToString();
					
					
					PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[4,7]).Value.ToString();
					
					int.TryParse(sDeviceCount, out deviceCount);
					
					// Click on Loop A node
							Common_Functions.ClickOnNavigationTreeItem(PanelNode);
				
							// Click on Expander node
							//Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
							
					
					
					// Verify device count and then add devices from panel accessories gallery or panel node gallery
					if(deviceCount>0)
					{
						if (sType.Equals("Accessories"))
						{
							
							Common_Functions.clickOnPanelAccessoriesTab();
							
							for(int k=1; k<=deviceCount;k++)
							{
								//Devices_Functions.AddDevicefromPanelAccessoriesGallery(sDeviceName,sType);
								Devices_Functions.AddDevicesfromGallery(sDeviceName,sType);
							}
						}
						else
						{
							
							Common_Functions.clickOnInventoryTab();
							
							for(int k=1; k<=deviceCount;k++)
							{
								//Devices_Functions.AddDevicesfromPanelNodeGallery(sDeviceName,sType,PanelType);
								Devices_Functions.AddDevicesfromGallery(sDeviceName,sType);
								
								// Verify expected backplane1
								if(ExpectedBackplane1.Equals("Yes"))
								{
									// Click on Expander node
									//Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
							
									// Click on Backplane Node
									Common_Functions.ClickOnNavigationTreeItem("Backplane");
				
										
										if(sType.Equals("Loops"))
										{
											int NoofPLXTree = k+1;
											string sNoofPLXTree = NoofPLXTree.ToString();
											
											string sPLXTreeName = "PLX/External Loop Card "+sNoofPLXTree+"";
											
											Common_Functions.VerifyAndClickNavigationTreeItemText(sPLXTreeName);
											
											//Conversion from string to int then again string
											int.TryParse(sMaxInOtherSlotCard, out MaxInOtherSlotCard);
											MaxInOtherSlotCard = MaxInOtherSlotCard-1;
											sMaxInOtherSlotCard = MaxInOtherSlotCard.ToString();
											
											string sBackplane1SlotCardName = "Other Slot Cards  (0 of "+sMaxInOtherSlotCard+")";
							
											Common_Functions.VerifyAndClickNavigationTreeItemText(sBackplane1SlotCardName);
					
											// Click on Panel node
											Common_Functions.ClickOnNavigationTreeItem(PanelNode);
											
											// Click on Expander node
											//Common_Functions.ClickOnNavigationTreeExpander("Backplane");
				
										}
										else
										{
											string sCount = k.ToString();
											string sBackplane1SlotCardName = "Other Slot Cards  ("+sCount+" of "+sMaxInOtherSlotCard+")";
																					
											Common_Functions.VerifyAndClickNavigationTreeItemText(sBackplane1SlotCardName);
											
											
											// Click on Panel node
											Common_Functions.ClickOnNavigationTreeItem(PanelNode);
				
										}
																		
									}
									else
									{
										Common_Functions.VerifyNavigationTreeItem("Backplane", ExpectedBackplane1);
									}
									
								}
								
									
								}
								
							// Click on Expander node
							//Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
								
							}
							// Click on Expander node
							//Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
							
					
				}

				
			}
			
			// Close Excel
			Excel_Utilities.CloseExcel();
		}
	
	}
}


