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
		
		
		
    }
}
