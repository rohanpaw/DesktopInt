﻿/*
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
		
		/***********************************************************************************************************
		 * Function Name: VerifyandClickOtherSlotCardsForBackplane1
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 06/05/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyandClickOtherSlotCardsForBackplane1(string slotCardName)
		{
			sOtherSlotCardName = slotCardName;
			repo.FormMe.Backplane1_OtherSlotCards.Click();
			Report.Log(ReportLevel.Info," Slot card name " +slotCardName + " is displayed  ");
		}
		
		/***********************************************************************************************************
		 * Function Name: VerifySlotCardsTextForBackplane2
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 02/05/2019  14/05/2019 - Alpesh Dhakad - Updated code
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifySlotCardsTextForBackplane2(string expectedText)
		{
			
			if(repo.FormMe.MainProcessorList.Backplane2_OtherSlotCardsWithPLXInfo.Exists())
			{
				string ActualText = repo.FormMe.MainProcessorList.Backplane2_OtherSlotCardsWithPLX.TextValue;
				
				if(ActualText.Equals(expectedText))
				{
					Report.Log(ReportLevel.Success,"Other slot cards text is as expected");
				}
				else
				{
					Report.Log(ReportLevel.Failure,"Other slot cards text is displayed as " +ActualText+ "instead of " +expectedText);
				}
			}
			else
			{
				string ActualText = repo.FormMe.Backplane2_OtherSlotCards.TextValue;
				
				if(ActualText.Equals(expectedText))
				{
					Report.Log(ReportLevel.Success,"Other slot cards text is as expected");
				}
				else
				{
					Report.Log(ReportLevel.Failure,"Other slot cards text is displayed as " +ActualText+ "instead of " +expectedText);
				}
			}
			
		}

		/********************************************************************
		 * Function Name: VerifySlotCardsAndBackplanesDistribution
		 * Function Details: To verify slot cards and backplane distribution
		 * Parameter/Arguments: sFileName, sAddDevicesSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 11/05/2019
		 ********************************************************************/
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
							repo.ProfileConsys1.NavigationTree.Expander.Click();
							repo.FormMe.tab_PanelAccessories.Click();
							for(int k=1; k<=deviceCount;k++)
							{
								Devices_Functions.AddDevicefromPanelAccessoriesGallery(sDeviceName,sType);
							}
						}
						else
						{
							repo.ProfileConsys1.NavigationTree.Expander.Click();
							repo.FormMe.tab_Inventory.Click();
							
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
					if(repo.FormMe.BackplaneOrXLMExternalLoopCard_ExpanderInfo.Exists())
					{
						repo.FormMe.BackplaneOrXLMExternalLoopCard_Expander.Click();
						Report.Log(ReportLevel.Success, "Backplane1 is available and displaying correctly");
						
						VerifyandClickOtherSlotCardsForBackplane1(sBackplane1SlotCardName);
						VerifySlotCardsTextForBackplane1(sBackplane1SlotCardName);
					}
					else
					{
						Report.Log(ReportLevel.Failure, "Backplane1 is not displayed");
					}
					
				}
				else
				{
					if(repo.FormMe.BackplaneOrXLMExternalLoopCard_ExpanderInfo.Exists())
					{
						Report.Log(ReportLevel.Failure, "Backplane1 should not be displayed");
					}
					
				}
				
				// Verify expected backplane2
				if(ExpectedBackplane2.Equals("Yes"))
				{
					if(repo.FormMe.Backplane2_ExpanderInfo.Exists())
					{
						repo.FormMe.Backplane2_Expander.Click();
						Report.Log(ReportLevel.Success, "Backplane2 is available and displaying correctly");
						
						VerifyandClickOtherSlotCardsForBackplane2(sBackplane2SlotCardName);
						VerifySlotCardsTextForBackplane2(sBackplane2SlotCardName);
					}
					else
					{
						Report.Log(ReportLevel.Failure, "Backplane2 is not displayed");
					}
					
				}
				else
				{
					if(repo.FormMe.Backplane2_ExpanderInfo.Exists())
					{
						Report.Log(ReportLevel.Failure, "Backplane2 should not be displayed");
					}
					
				}
				
				// Verify expected backplane3
				if(ExpectedBackplane3.Equals("Yes"))
				{
					if(repo.FormMe.Backplane3_ExpanderInfo.Exists())
					{
						repo.FormMe.Backplane3_Expander.Click();
						Report.Log(ReportLevel.Success, "Backplane3 is available and displaying correctly");
						
						VerifyandClickOtherSlotCardsForBackplane3(sBackplane3SlotCardName);
						VerifySlotCardsTextForBackplane3(sBackplane3SlotCardName);
					}
					else
					{
						Report.Log(ReportLevel.Failure, "Backplane3 is not displayed");
					}
				}
				else
				{
					if(repo.FormMe.Backplane3_ExpanderInfo.Exists())
					{
						Report.Log(ReportLevel.Failure, "Backplane3 should not be displayed");
					}
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
		
		
		
		/********************************************************************
		 * Function Name: VerifySlotCardsAndBackplanesDistributionWithOnePanel
		 * Function Details: To verify slot cards and backplane distribution
		 * Parameter/Arguments: sFileName, sAddDevicesSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 13/05/2019
		 ********************************************************************/
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
							repo.ProfileConsys1.NavigationTree.Expander.Click();
							repo.FormMe.tab_PanelAccessories.Click();
							for(int k=1; k<=deviceCount;k++)
							{
								Devices_Functions.AddDevicefromPanelAccessoriesGallery(sDeviceName,sType);
							}
						}
						else
						{
							repo.ProfileConsys1.NavigationTree.Expander.Click();
							repo.FormMe.tab_Inventory.Click();
							
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
					if(repo.FormMe.BackplaneOrXLMExternalLoopCard_ExpanderInfo.Exists())
					{
						repo.FormMe.BackplaneOrXLMExternalLoopCard_Expander.Click();
						Report.Log(ReportLevel.Success, "Backplane1 is available and displaying correctly");
						
						VerifyandClickOtherSlotCardsForBackplane1(sBackplane1SlotCardName);
						VerifySlotCardsTextForBackplane1(sBackplane1SlotCardName);
						
					}
					else
					{
						Report.Log(ReportLevel.Failure, "Backplane1 is not displayed");
					}
					
				}
				else
				{
					if(repo.FormMe.BackplaneOrXLMExternalLoopCard_ExpanderInfo.Exists())
					{
						Report.Log(ReportLevel.Failure, "Backplane1 should not be displayed");
					}
					
				}
				
				// Verify expected backplane2
				if(ExpectedBackplane2.Equals("Yes"))
				{
					if(repo.FormMe.Backplane2_ExpanderInfo.Exists())
					{
						repo.FormMe.Backplane2_Expander.Click();
						Report.Log(ReportLevel.Success, "Backplane2 is available and displaying correctly");
						
						VerifyandClickOtherSlotCardsForBackplane2(sBackplane2SlotCardName);
						VerifySlotCardsTextForBackplane2(sBackplane2SlotCardName);
						
					}
					else
					{
						Report.Log(ReportLevel.Failure, "Backplane2 is not displayed");
					}
					
				}
				else
				{
					if(repo.FormMe.Backplane2_ExpanderInfo.Exists())
					{
						Report.Log(ReportLevel.Failure, "Backplane2 should not be displayed");
					}
					
				}
				
				// Verify expected backplane3
				if(ExpectedBackplane3.Equals("Yes"))
				{
					if(repo.FormMe.Backplane3_ExpanderInfo.Exists())
					{
						repo.FormMe.Backplane3_Expander.Click();
						Report.Log(ReportLevel.Success, "Backplane3 is available and displaying correctly");
						
						VerifyandClickOtherSlotCardsForBackplane3(sBackplane3SlotCardName);
						VerifySlotCardsTextForBackplane3(sBackplane3SlotCardName);
					}
					else
					{
						Report.Log(ReportLevel.Failure, "Backplane3 is not displayed");
					}
				}
				else
				{
					if(repo.FormMe.Backplane3_ExpanderInfo.Exists())
					{
						Report.Log(ReportLevel.Failure, "Backplane3 should not be displayed");
					}
				}
				
			}
			
			// Close Excel
			Excel_Utilities.CloseExcel();
		}
		
		
		/***********************************************************************************************************
		 * Function Name: VerifySlotCardsTextForBackplane1
		 * Function Details: To Verify other SlotCards Text from Backplane1
		 * Parameter/Arguments: expectedText
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 13/05/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifySlotCardsTextForBackplane1(string expectedText)
		{
			string ActualText = repo.FormMe.Backplane1_OtherSlotCards.TextValue;
			
			if(ActualText.Equals(expectedText))
			{
				Report.Log(ReportLevel.Success,"Other slot cards text is as expected");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Other slot cards text is displayed as " +ActualText+ "instead of " +expectedText);
			}
			
		}
		
		/***********************************************************************************************************
		 * Function Name: VerifyandClickOtherSlotCardsForBackplane2
		 * Function Details: To verify and click on backplane 2
		 * Parameter/Arguments: slotCardName
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 13/05/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyandClickOtherSlotCardsForBackplane2(string slotCardName)
		{
			sOtherSlotCardName = slotCardName;
			
			if(repo.FormMe.MainProcessorList.Backplane2_OtherSlotCardsWithPLXInfo.Exists())
			{
				repo.FormMe.MainProcessorList.Backplane2_OtherSlotCardsWithPLX.Click();
			}
			else
			{
				repo.FormMe.Backplane2_OtherSlotCards.Click();
			}
			
			Report.Log(ReportLevel.Info," Slot card name " +slotCardName + " is displayed  ");
		}
		
		
		/***********************************************************************************************************
		 * Function Name: VerifyandClickOtherSlotCardsForBackplane3
		 * Function Details: To verify and click on backplane 3
		 * Parameter/Arguments: slotCardName
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 13/05/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyandClickOtherSlotCardsForBackplane3(string slotCardName)
		{
			sOtherSlotCardName = slotCardName;
			repo.FormMe.Backplane3_OtherSlotCards.Click();
			Report.Log(ReportLevel.Info," Slot card name " +slotCardName + " is displayed  ");
		}
		
		/***********************************************************************************************************
		 * Function Name: VerifySlotCardsTextForBackplane3
		 * Function Details: To Verify other SlotCards Text from Backplane3
		 * Parameter/Arguments: expectedText
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 13/05/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifySlotCardsTextForBackplane3(string expectedText)
		{
			string ActualText = repo.FormMe.Backplane3_OtherSlotCards.TextValue;
			
			if(ActualText.Equals(expectedText))
			{
				Report.Log(ReportLevel.Success,"Other slot cards text is as expected");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Other slot cards text is displayed as " +ActualText+ "instead of " +expectedText);
			}
			
		}

		/***********************************************************************************************************
		 * Function Name: VerifySlotCardsTextForBackplane3OnReopen
		 * Function Details: To Verify other SlotCards Text from Backplane3
		 * Parameter/Arguments: expectedText
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 13/05/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifySlotCardsTextForBackplane3OnReopen(string expectedText)
		{
			string ActualText = repo.FormMe.Backplane3_OtherSlotCards_Reopen.TextValue;
			
			if(ActualText.Equals(expectedText))
			{
				Report.Log(ReportLevel.Success,"Other slot cards text is as expected");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Other slot cards text is displayed as " +ActualText+ "instead of " +expectedText);
			}
			
		}
		
		/***********************************************************************************************************
		 * Function Name: VerifyPanelTypeDropdownOnSlotCardsPosition
		 * Function Details: string sFileName,string sAddDevicesSheet, string sPanelName
		 * Parameter/Arguments: expectedText
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 15/05/2019
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
							repo.ProfileConsys1.NavigationTree.Expander.Click();
							repo.FormMe.tab_PanelAccessories.Click();
							for(int k=1; k<=deviceCount;k++)
							{
								Devices_Functions.AddDevicefromPanelAccessoriesGallery(sDeviceName,sType);
							}
						}
						else
						{
							repo.ProfileConsys1.NavigationTree.Expander.Click();
							repo.FormMe.tab_Inventory.Click();
							
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
					if(repo.FormMe.BackplaneOrXLMExternalLoopCard_ExpanderInfo.Exists())
					{
						repo.FormMe.BackplaneOrXLMExternalLoopCard_Expander.Click();
						Report.Log(ReportLevel.Success, "Backplane1 is available and displaying correctly");
						
						VerifyandClickOtherSlotCardsForBackplane1(sBackplane1SlotCardName);
						VerifySlotCardsTextForBackplane1(sBackplane1SlotCardName);
						
					}
					else
					{
						Report.Log(ReportLevel.Failure, "Backplane1 is not displayed");
					}
					
				}
				else
				{
					if(repo.FormMe.BackplaneOrXLMExternalLoopCard_ExpanderInfo.Exists())
					{
						Report.Log(ReportLevel.Failure, "Backplane1 should not be displayed");
					}
					
				}
				
				// Verify expected backplane2
				if(ExpectedBackplane2.Equals("Yes"))
				{
					if(repo.FormMe.Backplane2_ExpanderInfo.Exists())
					{
						repo.FormMe.Backplane2_Expander.Click();
						Report.Log(ReportLevel.Success, "Backplane2 is available and displaying correctly");
						
						VerifyandClickOtherSlotCardsForBackplane2(sBackplane2SlotCardName);
						VerifySlotCardsTextForBackplane2(sBackplane2SlotCardName);
						
					}
					else
					{
						Report.Log(ReportLevel.Failure, "Backplane2 is not displayed");
					}
					
				}
				else
				{
					if(repo.FormMe.Backplane2_ExpanderInfo.Exists())
					{
						Report.Log(ReportLevel.Failure, "Backplane2 should not be displayed");
					}
					
				}
				
				// Verify expected backplane3
				if(ExpectedBackplane3.Equals("Yes"))
				{
					if(repo.FormMe.Backplane3_ExpanderInfo.Exists())
					{
						repo.FormMe.Backplane3_Expander.Click();
						Report.Log(ReportLevel.Success, "Backplane3 is available and displaying correctly");
						
						VerifyandClickOtherSlotCardsForBackplane3(sBackplane3SlotCardName);
						VerifySlotCardsTextForBackplane3(sBackplane3SlotCardName);
					}
					else
					{
						Report.Log(ReportLevel.Failure, "Backplane3 is not displayed");
					}
				}
				else
				{
					if(repo.FormMe.Backplane3_ExpanderInfo.Exists())
					{
						Report.Log(ReportLevel.Failure, "Backplane3 should not be displayed");
					}
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
		 * Last Update : 24/05/2019
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
				
				repo.ProfileConsys1.NavigationTree.Expander.Click();
				
				repo.FormMe.tab_PanelAccessories.Click();
				
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
		 * Last Update : 24/05/2019
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
				
				// Click on navigation tree expander
				repo.ProfileConsys1.NavigationTree.Expander.Click();
				
				// Click on panel accessories tab
				repo.FormMe.tab_PanelAccessories.Click();
				
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
		 * Last Update : 30/05/2019
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
				
				// Click on navigation tree expander
				repo.ProfileConsys1.NavigationTree.Expander.Click();
				
				// Click on panel accessories tab
				repo.FormMe.tab_PanelAccessories.Click();
				
				// Add devices from panel accessories gallery
				Devices_Functions.AddDevicefromPanelAccessoriesGallery(sDeviceName,sType);
				
				// Change FOM Value in Search properties
				Devices_Functions.ChangeFOMInSearchProperties(sFOMChange);
				
				// Verify and perform check or uncheck MPM checkbox in search properties
				Devices_Functions.CheckUncheckMPMCheckboxInSearchProperties(changeCheckboxStateTo);
				
				// Click on Site node
				repo.ProfileConsys1.SiteNode.Click();
				
				// Click on Shopping list tab
				repo.FormMe.ShoppingList.Click();
				
				// Verify shopping list count
				Devices_Functions.verifyShoppingList(shoppingListCount);
				Delay.Milliseconds(200);
				
				// Click on Export button
				repo.FormMe.Export2ndTime.Click();
				Delay.Milliseconds(200);
				
				// Click on Maximize button
				repo.PrintPreview.PARTMaximize.Click();
				
				// Click on export drop down button
				repo.PrintPreview.ExportDropdown.Click();
				
				// Click on excel format document
				repo.ExportDocument.ExcelFormat.Click();
				Delay.Duration(5000, false);
				
				// Set the attribute value to xls
				repo.ExportDocument.ExcelFormat.Element.SetAttributeValue("Text", "Xls");
				Delay.Duration(5000, false);
				
				// Click on OK button of export document
				repo.ExportDocument.ButtonOK.Click();
				Delay.Milliseconds(200);
				
				// Click on OK button of export document again
				repo.ExportDocument.ButtonOK.Click();
				
				// Click on shopping list Cell 18 of excel sheet
				repo.ShoppingListCompatibilityModeE.Cell18.Click();
				Delay.Milliseconds(200);
				
				// Verify Cell 18 text of excel sheet
				Libraries.Devices_Functions.verifyShoppingListDevicesTextForPxD(sFOMExpectedText);
				Delay.Milliseconds(0);
				
				// Verify Cell 22 text of excel sheet
				Devices_Functions.verifyShoppingListDevicesTextForPSC(sMPMExpectedText);
				
				// Verify Cell 26 text of excel sheet
				Devices_Functions.verifyShoppingListDevicesTextForThirdDevice(sANNExpectedText);
				
				// Click on button to close excel
				repo.ShoppingListCompatibilityModeE.btn_CloseExcel.Click();
				
				// Click on button to close print preview window
				repo.PrintPreview.btn_CloseB.Click();
				
				// Click on navigation tree expander
				repo.ProfileConsys1.NavigationTree.Expander.Click();
				
				// Click on panel accessories tab
				repo.FormMe.tab_PanelAccessories.Click();
				
				// Verify and perform check or uncheck MPM checkbox in search properties
				Devices_Functions.CheckUncheckMPMCheckboxInSearchProperties(changeCheckboxStateToAgain);
				
				// Click on site node
				repo.ProfileConsys1.SiteNode.Click();
				
				// Click on shopping list tab
				repo.FormMe.ShoppingList.Click();
				
				// Verify shopping list count
				Devices_Functions.verifyShoppingList(shoppingListCountAfterUncheck);
				Delay.Milliseconds(500);
				
				// Click on Export button
				repo.FormMe.Export2ndTime.Click();
				Delay.Milliseconds(200);
				
				// Click on maximize button
				repo.PrintPreview.PARTMaximize.Click();
				
				// Click on export drop down button
				repo.PrintPreview.ExportDropdown.Click();
				
				// Click on Export document to select excel format
				repo.ExportDocument.ExcelFormat.Click();
				Delay.Duration(5000, false);
				
				// Set the attribute value to xls
				repo.ExportDocument.ExcelFormat.Element.SetAttributeValue("Text", "Xls");
				Delay.Duration(5000, false);
				
				// Click on Ok button
				repo.ExportDocument.ButtonOK.Click();
				Delay.Milliseconds(200);
				
				// Click on Ok button again
				repo.ExportDocument.ButtonOK.Click();
				
				// Click Cell 18 text of excel sheet
				repo.ShoppingListCompatibilityModeE.Cell18.Click();
				Delay.Milliseconds(200);
				
				// Verify Cell 18 text of excel sheet
				Libraries.Devices_Functions.verifyShoppingListDevicesTextForPxD(sFOMExpectedText);
				Delay.Milliseconds(0);
				
				// Verify Cell 22 text of excel sheet
				Devices_Functions.verifyShoppingListDevicesTextForPSC(sMPMExpectedTextAfterUncheck);
				
				// Verify Cell 26 text of excel sheet
				Devices_Functions.verifyShoppingListDevicesTextForThirdDevice(sANNExpectedTextAfterUncheck);
				
				// Click to close excel sheet
				repo.ShoppingListCompatibilityModeE.btn_CloseExcel.Click();
				
				// Click on close button
				repo.PrintPreview.btn_CloseB.Click();
				
				// Click on Site node
				repo.ProfileConsys1.SiteNode.Click();
				
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
		 * Function Name: VerifyShoppingListOnSelectingFOMandMPM
		 * Function Details:
		 * Parameter/Arguments: string sFileName,string sAddDevicesSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 27/05/2019
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
			string sSecondCPUType,sSecondDeviceName,sSecondDeviceType,sSecondShoppingListCount,secondDisableDevice,secondDeviceState;
			
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
				
				
				int.TryParse(sShoppingListCount, out shoppingListCount);
				int.TryParse(sSecondShoppingListCount, out secondShoppingListCount);
				
				
				// Add panels
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				
				// Click on navigation tree expander
				repo.ProfileConsys1.NavigationTree.Expander.Click();
				
				// Click on panel accessories tab
				repo.FormMe.tab_PanelAccessories.Click();
				
				// Add devices from panel accessories gallery
				Devices_Functions.AddDevicefromPanelAccessoriesGallery(sDeviceName,sType);
				
				// Set newDevice name value in sDeviceName
				sDeviceName = newDeviceName;
				
				// Verify Enable or disable of devices in panel accessories gallery
				Devices_Functions.VerifyEnableDisablePanelAccessoriesGallery(sType,sDeviceName,initialState);
				
				// Click on site node
				repo.ProfileConsys1.SiteNode.Click();
				
				// Click on shopping list tab
				repo.FormMe.ShoppingList.Click();
				
				// Verify shopping list count
				Devices_Functions.verifyShoppingList(shoppingListCount);
				
				// Click on panel accessories tab
				repo.FormMe.tab_Panel_Network.Click();
				
				// Click on site node
				repo.ProfileConsys1.SiteNode.Click();
				
				// Add one panel after adding 1 one panel
				Panel_Functions.AddOnePanel(2,secondPanelName,sSecondCPUType);
			
				// Click on navigation tree expander
				repo.ProfileConsys1.NavigationTree.Expander.Click();
				
				// Click on panel accessories tab
				repo.FormMe.tab_PanelAccessories.Click();
				
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
				
				// Click on site node
				repo.ProfileConsys1.SiteNode.Click();
				
				// Click on shopping list tab
				repo.FormMe.ShoppingList.Click();
				
				// Verify shopping list count
				Devices_Functions.verifyShoppingList(secondShoppingListCount);
				Delay.Milliseconds(500);
				
				// Click on Export button
				repo.FormMe.Export2ndTime.Click();
				Delay.Milliseconds(200);
				
				// Click on maximize button
				repo.PrintPreview.PARTMaximize.Click();
				
				// Click on export drop down button
				repo.PrintPreview.ExportDropdown.Click();
				
				// Click on Export document to select excel format
				repo.ExportDocument.ExcelFormat.Click();
				Delay.Duration(5000, false);
				
				// Set the attribute value to xls
				repo.ExportDocument.ExcelFormat.Element.SetAttributeValue("Text", "Xls");
				Delay.Duration(5000, false);
				
				// Click on Ok button
				repo.ExportDocument.ButtonOK.Click();
				Delay.Milliseconds(200);
				
				// Click on Ok button again
				repo.ExportDocument.ButtonOK.Click();
			
				// Verify shopping list excel text for first device and panel
				Devices_Functions.verifyShoppingListDevicesTextForCell3And14(PanelName,sDeviceName);
				
				// Verify shopping list excel text for second device and panel
				Devices_Functions.verifyShoppingListDevicesTextForCell17And21(secondPanelName,sSecondDeviceName);
				
				// Click to close excel sheet
				repo.ShoppingListCompatibilityModeE.btn_CloseExcel.Click();
				
				// Click on close button
				repo.PrintPreview.btn_CloseB.Click();
				
				// Click on Site node
				repo.ProfileConsys1.SiteNode.Click();
				
				// Verify if row count is more than 8 then delete the panel
				if(rows!=8)
				{
					// Delete Panel
					Panel_Functions.DeletePanel(2,PanelNode,1);
				}
				
			}
			
			
			
			Excel_Utilities.CloseExcel();
		}
	}
}


		