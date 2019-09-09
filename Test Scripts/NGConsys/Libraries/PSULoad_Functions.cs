/*
 * Created by Ranorex
 * User: jdhakaa
 * Date: 12/26/2018
 * Time: 4:53 PM
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
	public class PSULoad_Functions
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
		
		static string sCell
		{
			get { return repo.sCell; }
			set { repo.sCell = value; }
		}
		
		static string sPsuV
		{
			get { return repo.sPsuV; }
			set { repo.sPsuV = value; }
		}
		
		static string sMainProcessorGalleryIndex
		{
			get { return repo.sMainProcessorGalleryIndex; }
			set { repo.sMainProcessorGalleryIndex = value; }
		}
		
		static string sLabelName
		{
			get { return repo.sLabelName; }
			set { repo.sLabelName = value; }
		}
		
		static string sPhysicalLayoutDeviceIndex
		{
			get { return repo.sPhysicalLayoutDeviceIndex; }
			set { repo.sPhysicalLayoutDeviceIndex = value; }
		}
		
		static string sRowIndex
		{
			get { return repo.sRowIndex; }
			set { repo.sRowIndex = value; }
		}
		
		
		/*****************************************************************************************************************
		 * Function Name: verifyMax5VPSULoad
		 * Function Details: To Verify maximum 5V PSU load value
		 * Parameter/Arguments:   expected Maximum value, panel type (FIM or PFI)  and row number is 12 by default for FIM
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 26/12/2018
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyMax5VPSULoad(string expectedMax5VPSU, string PanelType, int rowNumber)
		{
			// Verify panel type and then accordingly assign sRow value
			if(PanelType.Equals("FIM"))
			{
				sRow = (rowNumber).ToString();
			}
			else
			{
				sRow = (rowNumber+1).ToString();
			}
			
			// Click on Physical layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			// Fetch 5V PSU Load maximum limit value
			string max5VPsu = repo.ProfileConsys1.Max5VPsu.TextValue;
			
			// Compare max5VPSU value with expected value
			if(max5VPsu.Equals(expectedMax5VPSU))
			{
				Report.Log(ReportLevel.Success,"Max 5V PSU value " + max5VPsu + " is displayed correctly " );
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Max 5V PSU value is not displayed correctly, it is displayed as: " + max5VPsu + " instead of : " +expectedMax5VPSU);
			}
			
			// Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
		}
		
		
		/*****************************************************************************************************************
		 * Function Name: verify5VPSULoadValue
		 * Function Details: To Verify 5V PSU load value
		 * Parameter/Arguments:   expected value, panel type (FIM or PFI)
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 09/01/2019 Removed 1 argument 25/01/19 Alpesh Dhakad - Add sCell and also updated Xpath of Psu5VLoad
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verify5VPSULoadValue(string expected5VPSU, string PanelType)
		{
			// Verify panel type and then accordingly assign sRow value
			if(PanelType.Equals("FIM"))
			{
				sRow = (12).ToString();
				sCell= "[0]";
			}
			else
			{
				sRow = (13).ToString();
				sCell= "[0]";
			}
			
			// Assign sPsuV value from sPSU5VLoad parameter
			sPsuV=expected5VPSU;
			
			// Click on Physical layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			// Fetch PSU5V value and store in Actual 5VPSU value
			string Actual5VPSUValue = repo.FormMe.Psu5VLoad.TextValue;
			
			// Compare Actual and Expected 5V PSU load value
			if(Actual5VPSUValue.Equals(expected5VPSU))
			{
				Report.Log(ReportLevel.Success,"5V PSU value " + Actual5VPSUValue + " is displayed correctly " );
			}
			else
			{
				Report.Log(ReportLevel.Failure,"5V PSU value is not displayed correctly, it is displayed as: " + Actual5VPSUValue + " instead of : " +expected5VPSU);
			}
			
			// CLick on Points tab
			repo.ProfileConsys1.tab_Points.Click();
		}
		
		/*****************************************************************************************************************
		 * Function Name: verify5VPsuLoadOnAdditionDeletionOfAccessories
		 * Function Details: verify 5V Psu Load On Addition and Deletion Of Accessories
		 * Parameter/Arguments: file name and add panel sheet name  and row number is 12 by default for FIM and 13 for PFI
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 28/01/2019  Alpesh Dhakad- 29/07/2019 - Updated script as per new build xpath updates
		 * Alpesh Dhakad - 16/08/2019 - Updated with new navigation tree method, xpath and devices gallery 
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verify5VPsuLoadOnAdditionDeletionOfAccessories(string sFileName,string sAddPanelandDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelandDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared string type
			string PanelName,PanelNode,CPUType,sRowNumber,PanelType,expected5VPSU,expected2nd5VPSU,expected3rd5VPSU,sType;
			int rowNumber;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sRowNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				expected5VPSU = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				ModelNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				expected2nd5VPSU = ((Range)Excel_Utilities.ExcelRange.Cells[i,10]).Value.ToString();
				expected3rd5VPSU = ((Range)Excel_Utilities.ExcelRange.Cells[i,11]).Value.ToString();
				
				int.TryParse(sRowNumber, out rowNumber);
				
				// Add panels using test data in excel sheet
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				Report.Log(ReportLevel.Info, "Panel "+PanelName+" added successfully");
				
				// Click on Expander node
				Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
				
				// Click on Loop Card node
				Common_Functions.ClickOnNavigationTreeExpander(PanelType);
				
				// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				// Verify 24V PSU load value
				verify5VPSULoadValue(expected5VPSU,PanelType);
				
				// Click on Expander node
				Common_Functions.ClickOnNavigationTreeItem(PanelNode);
				
				
				// Add devices from Panel node gallery
				Devices_Functions.AddDevicesfromMainProcessorGallery(ModelNumber,sType,PanelType);
				
				// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				
				// Verify 24V PSU load value
				verify5VPSULoadValue(expected2nd5VPSU,PanelType);
				
				// Click on Expander node
				Common_Functions.ClickOnNavigationTreeItem(PanelNode);
				
				// Delete devices using its Label name
				Devices_Functions.DeleteDeviceUsingLabel(sLabelName);
				
				// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				// Verify 24V PSU load value
				verify5VPSULoadValue(expected3rd5VPSU,PanelType);
				
				// Delete added Panel
				Panel_Functions.DeletePanel(1,PanelNode,1);
			}
			
			//Close opened excel sheet
			Excel_Utilities.CloseExcel();
			
		}

		
		

		/*****************************************************************************************************************
		 * Function Name: verifyMax24VPSULoad
		 * Function Details: To Verify maximum 24V PSU load value
		 * Parameter/Arguments:   expected Maximum value, panel type (FIM or PFI)  and row number is 13 by default for FIM
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 07/01/2019
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyMax24VPSULoad(string expectedMax24VPSU, string PanelType, int rowNumber)
		{
			// Verify panel type and then accordingly assign sRow value
			if(PanelType.Equals("FIM"))
			{
				sRow = (rowNumber).ToString();
			}
			else
			{
				sRow = (rowNumber+1).ToString();
			}
			
			// Click on Physical layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			// Fetch 24V PSU Load maximum limit value
			string max24VPsu = repo.ProfileConsys1.Max24VPsu.TextValue;
			
			// Compare max5VPSU value with expected value
			if(max24VPsu.Equals(expectedMax24VPSU))
			{
				Report.Log(ReportLevel.Success,"Max 24V PSU value " + max24VPsu + " is displayed correctly " );
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Max 24V PSU value is not displayed correctly, it is displayed as: " + max24VPsu + " instead of : " +expectedMax24VPSU);
			}
			
			//Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
		}
		
		
		/*****************************************************************************************************************
		 * Function Name: verify24VPSULoadValue
		 * Function Details: To Verify 24V PSU load value
		 * Parameter/Arguments:   expected value, panel type (FIM or PFI)  and row number is 13 by default for FIM
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 09/01/2019 Removed 1 argument 25/01/19 Alpesh Dhakad - Add sCell and also updated Xpath of Psu24VLoad
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verify24VPSULoadValue(string expected24VPSU, string PanelType)
		{
			// Verify panel type and then accordingly assign sRow value
			if(PanelType.Equals("FIM"))
			{
				sRow = (13).ToString();
				sCell= "[1]";
			}
			else
			{
				sRow = (14).ToString();
				sCell= "[1]";
			}
			
			// Assign sPsuV value from sPSU24VLoad parameter
			sPsuV=expected24VPSU;
			
			//Click on Physical Layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			// Fetch PSU24V value and store in Actual 24VPSU value
			string Actual24VPSUValue = repo.FormMe.Psu24VLoad.TextValue;
			
			// Compare Actual and Expected 24V PSU load value
			if(Actual24VPSUValue.Equals(expected24VPSU))
			{
				Report.Log(ReportLevel.Success,"24V PSU value " + Actual24VPSUValue + " is displayed correctly " );
			}
			else
			{
				Report.Log(ReportLevel.Failure,"24V PSU value is not displayed correctly, it is displayed as: " + Actual24VPSUValue + " instead of : " +expected24VPSU);
			}
			
			//Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
		}
		
		
		/*****************************************************************************************************************
		 * Function Name: verifyMax24VPSULoadOnAdditionOfPanels
		 * Function Details: To Verify maximum 24V PSU load value after addition of panels
		 * Parameter/Arguments:   Filename and Add devices sheet as excel input and row number is 13 by default for FIM
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 07/01/2019 Alpesh Dhakad - 30/07/2019 & 21/08/2019 - Updated test scripts as per new build and xpaths
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyMax24VPSULoadOnAdditionOfPanels(string sFileName,string sAddDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared string type
			string PanelName,PanelNode,CPUType,sRowNumber,expectedMax24VPSU,PanelType;
			int rowNumber;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sRowNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				expectedMax24VPSU = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				
				int.TryParse(sRowNumber, out rowNumber);
				
				// Add panels using test data in excel sheet
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				
				
				// Click on Expander node
				Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
				
				// Click on Loop Card node
				Common_Functions.ClickOnNavigationTreeExpander(PanelType);
				
				// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				
				// Verify max 24V PSU load value
				verifyMax24VPSULoad(expectedMax24VPSU,PanelType,rowNumber);
				
				// Delete panel using PanelNode details from excel sheet
				Panel_Functions.DeletePanel(1,PanelNode,1);
				
				
			}
			//Close opened excel sheet
			Excel_Utilities.CloseExcel();
			
		}
		
		/*****************************************************************************************************************
		 * Function Name: verify24VLoadOnChangingCPU
		 * Function Details: verify 24V Load On Changing CPU of the panel
		 * Parameter/Arguments: file name and add panel sheet name  and row number is 13 by default for FIM
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 08/01/2019  Alpesh Dhakad - 30/07/2019 & 21/08/2019- Updated test scripts as per new build and xpaths
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verify24VLoadOnChangingCPU(string sFileName,string sAddPanelSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared string type
			string PanelName,PanelNode,CPUType,sRowNumber,changeCPUType,PanelType,expectedMax24VPSU,expected24VPSU,change2CPUType,expected2nd24VPSU;
			int rowNumber;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sRowNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				changeCPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				expectedMax24VPSU = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				expected24VPSU = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				change2CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				expected2nd24VPSU = ((Range)Excel_Utilities.ExcelRange.Cells[i,10]).Value.ToString();
				
				int.TryParse(sRowNumber, out rowNumber);
				
				// Add panels using test data in excel sheet
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				
				// Click on Panel node
				Common_Functions.ClickOnNavigationTreeItem(PanelNode);
				
				// Change CPU Type as per test data in sheet
				Panel_Functions.ChangeCPUType(changeCPUType);
				
				
				// Click on Expander node
				Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
				
				// Click on Loop Card node
				Common_Functions.ClickOnNavigationTreeExpander(PanelType);
				
				// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				
				// Verify max 24V PSU load value
				verifyMax24VPSULoad(expectedMax24VPSU,PanelType,rowNumber);
				
				// Verify 24V PSU load value
				verify24VPSULoadValue(expected24VPSU,PanelType);
				
				// Click on Panel node
				Common_Functions.ClickOnNavigationTreeItem(PanelNode);
				
				// Click on Search Properties text field
				repo.ProfileConsys1.txt_SearchProperties.Click();
				
				// Click on Panel node
				Common_Functions.ClickOnNavigationTreeItem(PanelNode);
				
				// Change CPU Type as per test data in sheet
				Panel_Functions.ChangeCPUType(change2CPUType);
				
				// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				// Verify 24V PSU load value
				verify24VPSULoadValue(expected2nd24VPSU,PanelType);

				// Delete panel using PanelNode details from excel sheet
				Panel_Functions.DeletePanel(1,PanelNode,1);
				
			}
			//Close opened excel sheet
			Excel_Utilities.CloseExcel();
			
		}
		
		
		/*****************************************************************************************************************
		 * Function Name: verify24VPsuLoadOnAdditionDeletionOfLoopCards
		 * Function Details: verify 24VPsu Load On Addition and Deletion Of LoopCards devices
		 * Parameter/Arguments: file name and add panel sheet name  and row number is 13 by default for FIM and 14 for PFI
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 16/01/2019  Alpesh Dhakad - 30/07/2019 - Updated test scripts as per new build and xpaths
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verify24VPsuLoadOnAdditionDeletionOfLoopCards(string sFileName,string sAddPanelandDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelandDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared string type
			string PanelName,PanelNode,CPUType,sRowNumber,PanelType,expected24VPSU,expected2nd24VPSU,expected3rd24VPSU,sType;
			int rowNumber;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sRowNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				expected24VPSU = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				ModelNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				expected2nd24VPSU = ((Range)Excel_Utilities.ExcelRange.Cells[i,10]).Value.ToString();
				expected3rd24VPSU = ((Range)Excel_Utilities.ExcelRange.Cells[i,11]).Value.ToString();
				
				int.TryParse(sRowNumber, out rowNumber);
				
				// Add panels using test data in excel sheet
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				// Click on Expander node
				Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
				
				// Click on Loop Card node
				Common_Functions.ClickOnNavigationTreeExpander(PanelType);
				
				// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
				// Verify 24V PSU load value
				verify24VPSULoadValue(expected24VPSU,PanelType);
				
				// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
				// Add Devices from gallery
				Devices_Functions.AddDevicesfromMainProcessorGallery(ModelNumber,sType,PanelType);
				
				// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				// Verify 24V PSU load value
				verify24VPSULoadValue(expected2nd24VPSU,PanelType);
				
				// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
				repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1.Click();
				
				repo.ProfileConsys1.btn_Delete.Click();
				Report.Log(ReportLevel.Success, "Device "+sLabelName+" deleted successfully");
				
				// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				// Verify 24V PSU load value
				verify24VPSULoadValue(expected3rd24VPSU,PanelType);
				
				// Delete added Panel
				Panel_Functions.DeletePanel(1,PanelNode,1);
			}
			
			//Close opened excel sheet
			Excel_Utilities.CloseExcel();
			
		}
		
		
		/*****************************************************************************************************************
		 * Function Name: verify24VPsuLoadOnAdditionDeletionOfSlotCards
		 * Function Details: verify 24VPsu Load On Addition and Deletion Of Slot Card
		 * Parameter/Arguments: file name and add panel sheet name  and row number is 13 by default for FIM and 14 for PFI
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 18/01/2019   Alpesh Dhakad - 30/07/2019 & 21/08/2019 - Updated test scripts as per new build and xpaths
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verify24VPsuLoadOnAdditionDeletionOfSlotCards(string sFileName,string sAddPanelandDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelandDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared string type
			string PanelName,PanelNode,CPUType,sRowNumber,PanelType,expected24VPSU,expected2nd24VPSU,expected3rd24VPSU,expected4th24VPSU,sType,ModelNumber1,sLabelName1,sType1;
			int rowNumber;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sRowNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				expected24VPSU = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				ModelNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				expected2nd24VPSU = ((Range)Excel_Utilities.ExcelRange.Cells[i,10]).Value.ToString();
				ModelNumber1 = ((Range)Excel_Utilities.ExcelRange.Cells[i,11]).Value.ToString();
				sLabelName1 = ((Range)Excel_Utilities.ExcelRange.Cells[i,12]).Value.ToString();
				sType1 = ((Range)Excel_Utilities.ExcelRange.Cells[i,13]).Value.ToString();
				expected3rd24VPSU = ((Range)Excel_Utilities.ExcelRange.Cells[i,14]).Value.ToString();
				expected4th24VPSU = ((Range)Excel_Utilities.ExcelRange.Cells[i,15]).Value.ToString();
				
				
				int.TryParse(sRowNumber, out rowNumber);
				
				// Add panels using test data in excel sheet
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				// Click on Expander node
				Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
				
				// Click on Loop Card node
				Common_Functions.ClickOnNavigationTreeExpander(PanelType);
				
				// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				// Verify 24V PSU load value
				verify24VPSULoadValue(expected24VPSU,PanelType);
				
				// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
				// Split Device name and then add devices as per the device name and number of devices from Panel node gallery
				string[] splitDeviceName  = ModelNumber.Split(',');
				int splitDevicesCount  = ModelNumber.Split(',').Length;
				
				for(int j=0; j<=(splitDevicesCount-1); j++)
				{
					ModelNumber = splitDeviceName[j];
					Devices_Functions.AddDevicesfromMainProcessorGallery(ModelNumber,sType,PanelType);
				}
				
				// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				// Verify 24V PSU load value
				verify24VPSULoadValue(expected2nd24VPSU,PanelType);
				
					// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
				// Click on Panel Accessories in Panel node
				repo.FormMe.tab_PanelAccessories.Click();

				// Split Device name and then add devices as per the device name and number of devices from Panel node gallery
				ModelNumber = ModelNumber1;
				sType = sType1;
				string[] splitDeviceName1  = ModelNumber.Split(',');
				int splitDevicesCount1  = ModelNumber.Split(',').Length;
				
				for(int k=0; k<=(splitDevicesCount1-1); k++)
				{
					ModelNumber = splitDeviceName1[k];
					Devices_Functions.AddDevicefromPanelAccessoriesGallery(ModelNumber,sType);
				}
				
				
				// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				// Verify 24V PSU load value
				verify24VPSULoadValue(expected3rd24VPSU,PanelType);
				
				// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
				// Split Device name and then delete devices using label name
				string[] splitLabelName  = sLabelName.Split(',');
				int splitLabelCount  = sLabelName.Split(',').Length;
				
				for(int l=0; l<=(splitLabelCount-1); l++)
				{
					sLabelName = splitLabelName[l];
					Devices_Functions.DeleteDeviceUsingLabel(sLabelName);
				}
				
				// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				// Verify 24V PSU load value
				verify24VPSULoadValue(expected4th24VPSU,PanelType);
				
				// Delete added Panel
				Panel_Functions.DeletePanel(1,PanelNode,1);
			}
			
			//Close opened excel sheet
			Excel_Utilities.CloseExcel();
			
		}
		
		
		/*****************************************************************************************************************
		 * Function Name: verify24VPsuLoadOnAdditionDeletionOfAccessories
		 * Function Details: verify 24VPsu Load On Addition and Deletion Of Accessories
		 * Parameter/Arguments: file name and add panel sheet name  and row number is 13 by default for FIM and 14 for PFI
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 21/01/2019  Alpesh Dhakad - 30/07/2019 & 21/08/2019- Updated test scripts as per new build and xpaths
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verify24VPsuLoadOnAdditionDeletionOfAccessories(string sFileName,string sAddPanelandDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelandDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared string type
			string PanelName,PanelNode,CPUType,sRowNumber,PanelType,expected24VPSU,expected2nd24VPSU,expected3rd24VPSU,sType;
			int rowNumber;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sRowNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				expected24VPSU = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				ModelNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				expected2nd24VPSU = ((Range)Excel_Utilities.ExcelRange.Cells[i,10]).Value.ToString();
				expected3rd24VPSU = ((Range)Excel_Utilities.ExcelRange.Cells[i,11]).Value.ToString();
				
				int.TryParse(sRowNumber, out rowNumber);
				
				// Add panels using test data in excel sheet
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				Report.Log(ReportLevel.Info, "Panel "+PanelName+" added successfully");
				
				// Click on Expander node
				Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
				
				// Click on Loop Card node
				Common_Functions.ClickOnNavigationTreeExpander(PanelType);
				
				// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				// Verify 24V PSU load value
				verify24VPSULoadValue(expected24VPSU,PanelType);
				
				// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
				// Add devices from Panel node gallery
				Devices_Functions.AddDevicesfromMainProcessorGallery(ModelNumber,sType,PanelType);
				
				// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				// Verify 24V PSU load value
				verify24VPSULoadValue(expected2nd24VPSU,PanelType);
				
				// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
				// Delete devices using its Label name
				Devices_Functions.DeleteDeviceUsingLabel(sLabelName);
				
				// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				// Verify 24V PSU load value
				verify24VPSULoadValue(expected3rd24VPSU,PanelType);
				
				// Delete added Panel
				Panel_Functions.DeletePanel(1,PanelNode,1);
			}
			
			//Close opened excel sheet
			Excel_Utilities.CloseExcel();
			
		}

		/*****************************************************************************************************************
		 * Function Name: verify24VPsuLoadOnAdditionDeletionOfLoopDevicesInZetfastLoop
		 * Function Details: verify 24VPsu Load On Addition and Deletion Of loop devices in Zetfast loop
		 * Parameter/Arguments: file name and add panel sheet name  and row number is 13 by default for FIM and 14 for PFI
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 22/01/2019  Alpesh Dhakad - 30/07/2019 & 21/08/2019 - Updated test scripts as per new build and xpaths
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verify24VPsuLoadOnAdditionDeletionOfLoopDevicesInZetfastLoop(string sFileName,string sAddPanelandDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelandDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared string type
			string PanelName,PanelNode,CPUType,sRowNumber,PanelType,expected24VPSU,expected2nd24VPSU,expected3rd24VPSU,expected4th24VPSU,sType,ModelNumber1,sLabelName1,sType1;
			int rowNumber;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sRowNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				expected24VPSU = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				ModelNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				expected2nd24VPSU = ((Range)Excel_Utilities.ExcelRange.Cells[i,10]).Value.ToString();
				ModelNumber1 = ((Range)Excel_Utilities.ExcelRange.Cells[i,11]).Value.ToString();
				sLabelName1 = ((Range)Excel_Utilities.ExcelRange.Cells[i,12]).Value.ToString();
				sType1 = ((Range)Excel_Utilities.ExcelRange.Cells[i,13]).Value.ToString();
				expected3rd24VPSU = ((Range)Excel_Utilities.ExcelRange.Cells[i,14]).Value.ToString();
				expected4th24VPSU = ((Range)Excel_Utilities.ExcelRange.Cells[i,15]).Value.ToString();
				
				int.TryParse(sRowNumber, out rowNumber);
				
				// Add panels using test data in excel sheet
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				Report.Log(ReportLevel.Info, "Panel "+PanelName+" added successfully");
				
				// Click on Expander node
					Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
					
					// Click on Loop Card node
					Common_Functions.ClickOnNavigationTreeExpander(PanelType);
					
					// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
				// Verify 24V PSU load value
				verify24VPSULoadValue(expected24VPSU,PanelType);
				
				// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
				// Add devices from Panel node gallery
				Devices_Functions.AddDevicesfromMainProcessorGallery(ModelNumber,sType,PanelType);
				
				// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
				// Verify 24V PSU load value
				verify24VPSULoadValue(expected2nd24VPSU,PanelType);
				
				// Click on Backplane expander 
					Common_Functions.ClickOnNavigationTreeExpander("XLM");
					
					// Click on Zetfas C node
					Common_Functions.ClickOnNavigationTreeItem("XLM800-Zetfas-C");
					
				
				// Split Device name and then add devices as per the device name and number of devices from main gallery
				ModelNumber = ModelNumber1;
				string[] splitDeviceName  = ModelNumber.Split(',');
				int splitDevicesCount  = ModelNumber.Split(',').Length;
				
				for(int j=0; j<=(splitDevicesCount-1); j++)
				{
					ModelNumber = splitDeviceName[j];
					Devices_Functions.AddDevicesfromGallery(ModelNumber,sType1);
				}
				
				// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				// Verify 24V PSU load value
				verify24VPSULoadValue(expected3rd24VPSU,PanelType);
				
				// Click on Zetfas C node
					Common_Functions.ClickOnNavigationTreeItem("XLM800-Zetfas-C");
					
				// Split Device name and then delete devices using label name
				string[] splitLabelName  = sLabelName1.Split(',');
				int splitLabelCount  = sLabelName1.Split(',').Length;
				
				for(int k=0; k<=(splitLabelCount-1); k++)
				{
					sLabelName1 = splitLabelName[k];
					Devices_Functions.DeleteDeviceUsingLabel(sLabelName1);
				}
				
				// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				// Verify 24V PSU load value
				verify24VPSULoadValue(expected4th24VPSU,PanelType);
				
				// Delete added Panel
				Panel_Functions.DeletePanel(1,PanelNode,1);
			}
			
			//Close opened excel sheet
			Excel_Utilities.CloseExcel();
			
		}
		
		
		/*****************************************************************************************************************
		 * Function Name: verify24VPsuLoadOnAdditionDeletionOfLoopDevicesInXLMLoop
		 * Function Details: verify 24VPsu Load On Addition and Deletion Of loop devices in XLM loop
		 * Parameter/Arguments: file name and add panel sheet name  and row number is 13 by default for FIM and 14 for PFI
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 22/01/2019   Alpesh Dhakad - 30/07/2019 & 21/08/2019 - Updated test scripts as per new build and xpaths
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verify24VPsuLoadOnAdditionDeletionOfLoopDevicesInXLMLoop(string sFileName,string sAddPanelandDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelandDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared string type
			string PanelName,PanelNode,CPUType,sRowNumber,PanelType,expected24VPSU,expected2nd24VPSU,expected3rd24VPSU,expected4th24VPSU,sType,ModelNumber1,sLabelName1,sType1;
			int rowNumber;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sRowNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				expected24VPSU = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				ModelNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				expected2nd24VPSU = ((Range)Excel_Utilities.ExcelRange.Cells[i,10]).Value.ToString();
				ModelNumber1 = ((Range)Excel_Utilities.ExcelRange.Cells[i,11]).Value.ToString();
				sLabelName1 = ((Range)Excel_Utilities.ExcelRange.Cells[i,12]).Value.ToString();
				sType1 = ((Range)Excel_Utilities.ExcelRange.Cells[i,13]).Value.ToString();
				expected3rd24VPSU = ((Range)Excel_Utilities.ExcelRange.Cells[i,14]).Value.ToString();
				expected4th24VPSU = ((Range)Excel_Utilities.ExcelRange.Cells[i,15]).Value.ToString();
				
				int.TryParse(sRowNumber, out rowNumber);
				
				// Add panels using test data in excel sheet
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				Report.Log(ReportLevel.Info, "Panel "+PanelName+" added successfully");
				
				// Click on Expander node
					Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
					
					// Click on Loop Card node
					Common_Functions.ClickOnNavigationTreeExpander(PanelType);
					
					// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
				// Verify 24V PSU load value
				verify24VPSULoadValue(expected24VPSU,PanelType);
				
				// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
				// Add devices from Panel node gallery
				Devices_Functions.AddDevicesfromMainProcessorGallery(ModelNumber,sType,PanelType);
				
				// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
				// Verify 24V PSU load value
				verify24VPSULoadValue(expected2nd24VPSU,PanelType);
				
				// Click on Backplane expander 
					Common_Functions.ClickOnNavigationTreeExpander("XLM");
					
					// Click on Zetfas C node
					Common_Functions.ClickOnNavigationTreeItem("XLM800-Zetfas-C");
					
				// Split Device name and then add devices as per the device name and number of devices from main gallery
				ModelNumber = ModelNumber1;
				string[] splitDeviceName  = ModelNumber.Split(',');
				int splitDevicesCount  = ModelNumber.Split(',').Length;
				
				for(int j=0; j<=(splitDevicesCount-1); j++)
				{
					ModelNumber = splitDeviceName[j];
					Devices_Functions.AddDevicesfromGallery(ModelNumber,sType1);
				}
				
				// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				// Verify 24V PSU load value
				verify24VPSULoadValue(expected3rd24VPSU,PanelType);
				
					// Click on Zetfas C node
					Common_Functions.ClickOnNavigationTreeItem("XLM800-Zetfas-C");
					
				// Split Device name and then delete devices using label name
				string[] splitLabelName  = sLabelName1.Split(',');
				int splitLabelCount  = sLabelName1.Split(',').Length;
				
				for(int k=0; k<=(splitLabelCount-1); k++)
				{
					sLabelName1 = splitLabelName[k];
					Devices_Functions.DeleteDeviceUsingLabel(sLabelName1);
				}
				
				// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				// Verify 24V PSU load value
				verify24VPSULoadValue(expected4th24VPSU,PanelType);
				
				// Delete added Panel
				Panel_Functions.DeletePanel(1,PanelNode,1);
			}
			
			//Close opened excel sheet
			Excel_Utilities.CloseExcel();
			
		}
		
		
		/*****************************************************************************************************************
		 * Function Name: verify24VPsuLoadOnAdditionDeletionOfLoopDevicesInPLXLoop
		 * Function Details: verify 24VPsu Load On Addition and Deletion Of loop devices in PLX loop
		 * Parameter/Arguments: file name and add panel sheet name  and row number is 13 by default for FIM and 14 for PFI
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 23/01/2019  Alpesh Dhakad - 30/07/2019 & 21/08/2019 - Updated test scripts as per new build and xpaths
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verify24VPsuLoadOnAdditionDeletionOfLoopDevicesInPLXLoop(string sFileName,string sAddPanelandDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelandDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared string type
			string PanelName,PanelNode,CPUType,sRowNumber,PanelType,expected24VPSU,expected2nd24VPSU,expected3rd24VPSU,expected4th24VPSU,sType,ModelNumber1,sLabelName1,sType1;
			int rowNumber;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sRowNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				expected24VPSU = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				ModelNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				expected2nd24VPSU = ((Range)Excel_Utilities.ExcelRange.Cells[i,10]).Value.ToString();
				ModelNumber1 = ((Range)Excel_Utilities.ExcelRange.Cells[i,11]).Value.ToString();
				sLabelName1 = ((Range)Excel_Utilities.ExcelRange.Cells[i,12]).Value.ToString();
				sType1 = ((Range)Excel_Utilities.ExcelRange.Cells[i,13]).Value.ToString();
				expected3rd24VPSU = ((Range)Excel_Utilities.ExcelRange.Cells[i,14]).Value.ToString();
				expected4th24VPSU = ((Range)Excel_Utilities.ExcelRange.Cells[i,15]).Value.ToString();
				
				int.TryParse(sRowNumber, out rowNumber);
				
				// Add panels using test data in excel sheet
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				Report.Log(ReportLevel.Info, "Panel "+PanelName+" added successfully");
				
				// Click on Expander node
					Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
					
					// Click on Loop Card node
					Common_Functions.ClickOnNavigationTreeExpander(PanelType);
					
					// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
				// Verify 24V PSU load value
				verify24VPSULoadValue(expected24VPSU,PanelType);
				
				// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
				// Add devices from Panel node gallery
				Devices_Functions.AddDevicesfromMainProcessorGallery(ModelNumber,sType,PanelType);
				
				// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
					// Verify 24V PSU load value
				verify24VPSULoadValue(expected2nd24VPSU,PanelType);

				// Click on Backplane expander 
					Common_Functions.ClickOnNavigationTreeExpander("Backplane");
					
					// Click on Zetfas C node
					Common_Functions.ClickOnNavigationTreeExpander("PLX");
					
					// Click on PLX node
					Common_Functions.ClickOnNavigationTreeItem("PLX800-E");
				
				// Split Device name and then add devices as per the device name and number of devices from main gallery
				ModelNumber = ModelNumber1;
				string[] splitDeviceName  = ModelNumber.Split(',');
				int splitDevicesCount  = ModelNumber.Split(',').Length;
				
				for(int j=0; j<=(splitDevicesCount-1); j++)
				{
					ModelNumber = splitDeviceName[j];
					Devices_Functions.AddDevicesfromGallery(ModelNumber,sType1);
				}
				
				// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
				// Verify 24V PSU load value
				verify24VPSULoadValue(expected3rd24VPSU,PanelType);
				
				// Click on PLX node to add device
					Common_Functions.ClickOnNavigationTreeItem("PLX800-E");
				
				// Split Device name and then delete devices using label name
				string[] splitLabelName  = sLabelName1.Split(',');
				int splitLabelCount  = sLabelName1.Split(',').Length;
				
				for(int k=0; k<=(splitLabelCount-1); k++)
				{
					sLabelName1 = splitLabelName[k];
					Devices_Functions.DeleteDeviceUsingLabel(sLabelName1);
				}
				
				// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				// Verify 24V PSU load value
				verify24VPSULoadValue(expected4th24VPSU,PanelType);
				
				// Delete added Panel
				Panel_Functions.DeletePanel(1,PanelNode,1);
			}
			
			//Close opened excel sheet
			Excel_Utilities.CloseExcel();
			
		}
		
		/*************************************************************************************************************************
		 * Function Name: verify40VLoadOnChangingCPU
		 * Function Details: To Verify maximum 40V PSU load on CPU change
		 * Parameter/Arguments:   expected Maximum value, panel type (FIM or PFI)  and row number is 13 by default for FIM
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update : 08/01/2019 Alpesh Dhakad - 31/07/2019 & 21/08/2019 - Updated test scripts as per new build and xpaths
		 *************************************************************************************************************************/
		[UserCodeMethod]
		public static void verify40VLoadOnChangingCPU(string sFileName,string sAddPanelSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared string type
			string PanelName,PanelNode,CPUType,sRowNumber,changeCPUType,PanelType,expectedMax40VPSU,expected40VPSU,changePSUType;
			int rowNumber;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sRowNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				changeCPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				changePSUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				expectedMax40VPSU = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				expected40VPSU = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				// sPSUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,10]).Value.ToString();
				
				int.TryParse(sRowNumber, out rowNumber);
				
				// Add panels using test data in excel sheet
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				
				// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
				// Change CPU Type as per test data in sheet
				if (!changeCPUType.IsEmpty())
				{
					Panel_Functions.ChangeCPUType(changeCPUType);
				}
				
				//Change PSU of panel
				if (!changePSUType.IsEmpty())
				{
					Panel_Functions.ChangePSUType(changePSUType);
				}
				
				// Click on Expander node
					Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
					
					// Click on Loop Card node
					Common_Functions.ClickOnNavigationTreeExpander(PanelType);
					
					// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
				// Verify max 40V PSU load value
				verifyMax40VPSULoad(expectedMax40VPSU,PanelType);
				
				// Verify 40V PSU load value
				verify40VPSULoadValue(expected40VPSU,PanelType);
				
				// Delete panel using PanelNode details from excel sheet
				Panel_Functions.DeletePanel(1,PanelNode,1);
				
			}
			//Close opened excel sheet
			Excel_Utilities.CloseExcel();
			
		}
		
		/*****************************************************************************************************************
		 * Function Name: verify40VPSULoadValue
		 * Function Details: To Verify 40V PSU load value
		 * Parameter/Arguments:   expected value, panel type (FIM or PFI)
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update : 08/01/2019
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verify40VPSULoadValue(string expected40VPSU, string PanelType)
		{
			// Verify panel type and then accordingly assign sRow value
			if(PanelType.Equals("FIM"))
			{
				sRow = (14).ToString();
				sCell= "[2]";
			}
			else
			{
				sRow = (6).ToString();
				sCell= "[5]";
			}
			
			//Click on Physical Layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			// Fetch PSU40V value and store in Actual 40VPSU value
			string Actual40VPSUValue = repo.FormMe.Psu40VLoad.TextValue;
			
			// Compare Actual and Expected 40V PSU load value
			if(Actual40VPSUValue.Equals(expected40VPSU))
			{
				Report.Log(ReportLevel.Success,"40V PSU value " + Actual40VPSUValue + " is displayed correctly " );
			}
			else
			{
				Report.Log(ReportLevel.Failure,"40V PSU value is not displayed correctly, it is displayed as: " + Actual40VPSUValue + " instead of : " +expected40VPSU);
			}
			
			//Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
		}
		
		
		/*****************************************************************************************************************
		 * Function Name: verifyMax40VPSULoad
		 * Function Details: To Verify maximum 40V PSU load value
		 * Parameter/Arguments:   expected Maximum value, panel type (FIM or PFI)  and row number is 13 by default for FIM
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update : 09/01/2019
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyMax40VPSULoad(string expectedMax40VPSU, string PanelType)
		{
			// Verify panel type and then accordingly assign sRow value
			if(PanelType.Equals("FIM"))
			{
				sRow = (14).ToString();
			}
			else
			{
				sRow = (6).ToString();
			}
			
			// Click on Physical layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			// Fetch 40V PSU Load maximum limit value
			string max40VPsu = repo.FormMe.Max40VPsu.TextValue;
			
			// Compare max40VPSU value with expected value
			if(max40VPsu.Equals(expectedMax40VPSU))
			{
				Report.Log(ReportLevel.Success,"Max 40V PSU value " + max40VPsu + " is displayed correctly " );
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Max 40V PSU value is not displayed correctly, it is displayed as: " + max40VPsu + " instead of : " +expectedMax40VPSU);
			}
			
			//Click on Points tab
			repo.ProfileConsys1.tab_Points.Click();
		}
		
		/*******************************************************************************************************************************
		 * Function Name: verify40VLoadOnEthernetAddDelete
		 * Function Details: To Verify maximum 40V PSU load on CPU change
		 * Parameter/Arguments:   expected Maximum value, panel type (FIM or PFI)  and row number is 13 by default for FIM
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update : 08/01/2019  Alpesh Dhakad - 31/07/2019 & 21/08/2019 - Updated test scripts as per new build and xpaths
		 *******************************************************************************************************************************/
		[UserCodeMethod]
		public static void verify40VLoadOnEthernetAddDelete(string sFileName,string sAddPanelSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared variables
			string PanelName,PanelNode,s40VLoad,sDefault40V,CPUType,sRowNumber,sType,PanelType,sExpected40VPSU;
			int rowNumber;
			float FourtyVLoad,Default40V;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sRowNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				
				
				int.TryParse(sRowNumber, out rowNumber);
				
				// Add panels using test data in excel sheet
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				
				// Click on Expander node
					Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
					
					// Click on Loop Card node
					Common_Functions.ClickOnNavigationTreeExpander(PanelType);
					
				// Click on Loop Card node
					Common_Functions.ClickOnNavigationTreeExpander("Main");
					
				// Click on Ethernet node
			Common_Functions.ClickOnNavigationTreeItem("Ethernet");
			
				
				for(int j=8; j<=9; j++)
				{
					
					ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[j,7]).Value.ToString();
					sType = ((Range)Excel_Utilities.ExcelRange.Cells[j,8]).Value.ToString();
					sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[j,9]).Value.ToString();
					s40VLoad = ((Range)Excel_Utilities.ExcelRange.Cells[j,10]).Value.ToString();
					
					
					// Click on Ethernet node
			Common_Functions.ClickOnNavigationTreeItem("Ethernet");
			
					
					float.TryParse(s40VLoad, out FourtyVLoad);
					Devices_Functions.AddDevicesfromMainProcessorGallery(ModelNumber,sType,PanelType);
					Report.Log(ReportLevel.Info, "Device "+ModelNumber+" added successfully");
					
					
					// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
					//Get 40V load from UI
					sDefault40V = Get40VPSULoadValue(PanelType);
					
					//Generate expected 40V load
					float.TryParse(sDefault40V, out Default40V);
					float Expected40VPSU = Default40V+FourtyVLoad;
					sExpected40VPSU= Expected40VPSU.ToString("0.000");
					
					// Verify 40V PSU load value on addition of Ethernet
					verify40VPSULoadValue(sExpected40VPSU,PanelType);
					
					//Get 40V load from UI
					sDefault40V = Get40VPSULoadValue(PanelType);
					
					//Generate expected 40V load on Deletion
					float.TryParse(sDefault40V, out Default40V);
					Expected40VPSU = Default40V-FourtyVLoad;
					sExpected40VPSU = Expected40VPSU.ToString("0.000");
					
					// Click on Ethernet node
			Common_Functions.ClickOnNavigationTreeItem("Ethernet");
			
					repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1.Click();
					
					if(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info.Exists())
					{
						repo.ProfileConsys1.btn_Delete.Click();
						Validate.AttributeEqual(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info, "Text", sLabelName);
						Report.Log(ReportLevel.Success, "Device "+sLabelName+" deleted successfully");
						
						// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
						// Verify 40V PSU load value on addition of Ethernet
						verify40VPSULoadValue(sExpected40VPSU,PanelType);
					}
					
					else
					{
						
						Report.Log(ReportLevel.Failure, "Device "+sLabelName+" not found");
					}

					
				}
				
				
				// Delete panel using PanelNode details from excel sheet
				Panel_Functions.DeletePanel(1,PanelNode,1);
				
			}
			//Close opened excel sheet
			Excel_Utilities.CloseExcel();
			
		}
		
		/*****************************************************************************************************************
		 * Function Name: verify40VLoadOnRbusAddDelete
		 * Function Details: To Verify 40V load on addition/deletion of R-Bus connection
		 * Parameter/Arguments:   expected Maximum value, panel type (FIM or PFI)  and row number is 13 by default for FIM
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update : 22/01/2019 Alpesh Dhakad - 31/07/2019 & 21/08/2019 - Updated test scripts as per new build and xpaths
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verify40VLoadOnRbusAddDelete(string sFileName,string sAddPanelSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared variables
			string PanelName,PanelNode,s40VLoad,sDefault40V,CPUType,sRowNumber,sType,PanelType,sExpected40VPSU,sXBus40VLoad;
			int rowNumber;
			float RBusFourtyVLoad,Default40V,XBusFourtyVLoad;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sRowNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				
				
				int.TryParse(sRowNumber, out rowNumber);
				
				// Add panels using test data in excel sheet
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				Report.Log(ReportLevel.Info, "Panel "+PanelName+" added successfully");
				
				// Click on Expander node
					Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
					
					// Click on Loop Card node
					Common_Functions.ClickOnNavigationTreeExpander(PanelType);
					
				// Click on Loop Card node
					Common_Functions.ClickOnNavigationTreeExpander("Main");
					
				// Click on RBUS node
			Common_Functions.ClickOnNavigationTreeItem("R-BUS");
				
				for(int j=8; j<9; j++)
				{
					
					ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[j,7]).Value.ToString();
					sType = ((Range)Excel_Utilities.ExcelRange.Cells[j,8]).Value.ToString();
					sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[j,9]).Value.ToString();
					s40VLoad = ((Range)Excel_Utilities.ExcelRange.Cells[j,10]).Value.ToString();
					
					//Add RBus connection
					// Click on RBUS node
			Common_Functions.ClickOnNavigationTreeItem("R-BUS");
				
					float.TryParse(s40VLoad, out RBusFourtyVLoad);
					Devices_Functions.AddDevicesfromMainProcessorGallery(ModelNumber,sType,PanelType);
					Report.Log(ReportLevel.Info, "Device "+ModelNumber+" added successfully");
					
					// Add X-Bus to R-Bus
					ModelNumber = ((Range)Excel_Utilities.ExcelRange.Cells[j,11]).Value.ToString();
					sType = ((Range)Excel_Utilities.ExcelRange.Cells[j,12]).Value.ToString();
					//s40VLoad = ((Range)Excel_Utilities.ExcelRange.Cells[j,13]).Value.ToString();
					sXBus40VLoad = ((Range)Excel_Utilities.ExcelRange.Cells[j,14]).Value.ToString();
					
					//Select R-Bus node
					repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1.Click();
					float.TryParse(sXBus40VLoad, out XBusFourtyVLoad);
					Devices_Functions.AddDevicesfromMainProcessorGallery(ModelNumber,sType,PanelType);
					Report.Log(ReportLevel.Info, "Device "+ModelNumber+" added successfully");
					
					// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
					//Get 40V load from UI
					sDefault40V = Get40VPSULoadValue(PanelType);
					
					//Generate expected 40V load
					float.TryParse(sDefault40V, out Default40V);
					float Expected40VPSU = Default40V+RBusFourtyVLoad+XBusFourtyVLoad;
					sExpected40VPSU= Expected40VPSU.ToString("0.000");
					
					// Verify 40V PSU load value on addition of R-Bus & X-Bus template
					verify40VPSULoadValue(sExpected40VPSU,PanelType);
					
					//Get 40V load from UI
					sDefault40V = Get40VPSULoadValue(PanelType);
					
					//Generate expected 40V load on Deletion
					float.TryParse(sDefault40V, out Default40V);
					Expected40VPSU = Default40V-RBusFourtyVLoad-XBusFourtyVLoad;
					sExpected40VPSU = Expected40VPSU.ToString("0.000");
					
					// Click on RBUS node
			Common_Functions.ClickOnNavigationTreeItem("R-BUS");
				
					repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1.Click();
					
					if(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info.Exists())
					{
						repo.ProfileConsys1.btn_Delete.Click();
						Validate.AttributeEqual(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info, "Text", sLabelName);
						Report.Log(ReportLevel.Success, "Device "+sLabelName+" deleted successfully");
						
						// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
						// Verify 40V PSU load value on addition of Ethernet
						verify40VPSULoadValue(sExpected40VPSU,PanelType);
					}
					
					else
					{
						
						Report.Log(ReportLevel.Failure, "Device "+sLabelName+" not found");
					}

					
				}
				
				
				// Delete panel using PanelNode details from excel sheet
				Panel_Functions.DeletePanel(1,PanelNode,1);
				
			}
			//Close opened excel sheet
			Excel_Utilities.CloseExcel();
			
		}

		/*****************************************************************************************************************
		 * Function Name: Get40VPSULoadValue
		 * Function Details: To get 40V PSU load value
		 * Parameter/Arguments:   expected value, panel type (FIM or PFI)
		 * Output:40V load displayed on UI
		 * Function Owner: Shweta Bhosale
		 * Last Update : 22/01/2019
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static string Get40VPSULoadValue(string PanelType)
		{
			// Verify panel type and then accordingly assign sRow value
			if(PanelType.Equals("FIM"))
			{
				sRow = (14).ToString();
				sCell= "[2]";
			}
			else
			{
				sRow = (6).ToString();
				sCell= "[5]";
			}
			
			//Click on Physical Layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			// Fetch PSU40V value and store in Actual 40VPSU value
			string Actual40VPSUValue = repo.FormMe.Psu40VLoad.TextValue;
			
			return Actual40VPSUValue;
		}
		
		
		/**********************************************************************************************************************
		 * Function Name: verify40VLoadOnAccessoriesAddDelete
		 * Function Details: To Verify 40V load on addition/deletion of Accessory
		 * Parameter/Arguments:   expected Maximum value, panel type (FIM or PFI)  and row number is 13 by default for FIM
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update : 23/01/2019  Alpesh Dhakad - 31/07/2019 & 21/08/2019 - Updated test scripts as per new build and xpaths
		 **********************************************************************************************************************/
		[UserCodeMethod]
		public static void verify40VLoadOnAccessoriesAddDelete(string sFileName,string sAddPanelSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared variables
			string PanelName,PanelNode,s40VLoad,sDefault40V,CPUType,sRowNumber,sType,PanelType,sExpected40VPSU;
			int rowNumber;
			float RBusFourtyVLoad,Default40V;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sRowNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				
				
				int.TryParse(sRowNumber, out rowNumber);
				
				// Add panels using test data in excel sheet
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				
				// Click on Expander node
					Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
					
					// Click on Loop Card node
					Common_Functions.ClickOnNavigationTreeExpander(PanelType);
					
				// Click on Loop Card node
					Common_Functions.ClickOnNavigationTreeExpander("Main");
					
				// Click on RBUS node
			Common_Functions.ClickOnNavigationTreeItem("R-BUS");
				
				
				for(int j=8; j<9; j++)
				{
					
					ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[j,7]).Value.ToString();
					sType = ((Range)Excel_Utilities.ExcelRange.Cells[j,8]).Value.ToString();
					sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[j,9]).Value.ToString();
					s40VLoad = ((Range)Excel_Utilities.ExcelRange.Cells[j,10]).Value.ToString();
					
					//Add Printer
						// Click on RBUS node
			Common_Functions.ClickOnNavigationTreeItem("R-BUS");
				
					float.TryParse(s40VLoad, out RBusFourtyVLoad);
					Devices_Functions.AddDevicesfromMainProcessorGallery(ModelNumber,sType,PanelType);
					Report.Log(ReportLevel.Info, "Device "+ModelNumber+" added successfully");
					
					// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
					//Get 40V load from UI
					sDefault40V = Get40VPSULoadValue(PanelType);
					
					//Generate expected 40V load
					float.TryParse(sDefault40V, out Default40V);
					float Expected40VPSU = Default40V+RBusFourtyVLoad;
					sExpected40VPSU= Expected40VPSU.ToString("0.000");
					
					// Verify 40V PSU load value on addition printer
					verify40VPSULoadValue(sExpected40VPSU,PanelType);
					
					//Get 40V load from UI
					sDefault40V = Get40VPSULoadValue(PanelType);
					
					//Generate expected 40V load on Deletion
					float.TryParse(sDefault40V, out Default40V);
					Expected40VPSU = Default40V-RBusFourtyVLoad;
					sExpected40VPSU = Expected40VPSU.ToString("0.000");
					
					// Click on RBUS node
			Common_Functions.ClickOnNavigationTreeItem("R-BUS");
				
					repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1.Click();
					
					if(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info.Exists())
					{
						repo.ProfileConsys1.btn_Delete.Click();
						Validate.AttributeEqual(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info, "Text", sLabelName);
						Report.Log(ReportLevel.Success, "Device "+sLabelName+" deleted successfully");
						
						// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
						// Verify 40V PSU load value on addition of Ethernet
						verify40VPSULoadValue(sExpected40VPSU,PanelType);
					}
					
					else
					{
						
						Report.Log(ReportLevel.Failure, "Device "+sLabelName+" not found");
					}

					
				}
				
				
				// Delete panel using PanelNode details from excel sheet
				Panel_Functions.DeletePanel(1,PanelNode,1);
				
			}
			//Close opened excel sheet
			Excel_Utilities.CloseExcel();
			
		}

		
		/*****************************************************************************************************************
		 * Function Name: verify40VLoadOnZetfastLoopAddDelete
		 * Function Details: To Verify 40V load on addition/deletion of Zetfast loop with devices
		 * Parameter/Arguments:   expected Maximum value, panel type (FIM or PFI)  and row number is 13 by default for FIM
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update : 23/01/2019 Alpesh Dhakad - 01/08/2019 & 21/08/2019 - Updated test scripts as per new build and xpaths
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verify40VLoadOnZetfastLoopAddDelete(string sFileName,string sAddPanelSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared variables
			string PanelName,PanelNode,s40VLoad,sDefault40V,CPUType,sRowNumber,sType,PanelType,sExpected40VPSU;
			int rowNumber;
			float ZetfastFourtyVLoad,Default40V,Expected40VPSU;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sRowNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				
				int.TryParse(sRowNumber, out rowNumber);
				
				// Add panels using test data in excel sheet
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				
				// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
				//Add zetfast loop and devices and verify 40 V load
				for(int j=7; j<=9; j++)
				{
					ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[j,7]).Value.ToString();
					sType = ((Range)Excel_Utilities.ExcelRange.Cells[j,8]).Value.ToString();
					sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[j,9]).Value.ToString();
					s40VLoad = ((Range)Excel_Utilities.ExcelRange.Cells[j,10]).Value.ToString();
					
					if(j==7)
					{
						Devices_Functions.AddDevicesfromMainProcessorGallery(ModelNumber,sType,PanelType);
						
					}
					
					else
					{
						//                		// Click on XLM Loop Card Expander
//						repo.FormMe.BackplaneOrXLMExternalLoopCard_Expander.Click();
						
						// Click on XLM Loop C Node to add device
						repo.FormMe.XLMExternalLoopCardDevices_C.Click();

						Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
						Report.Log(ReportLevel.Info, "Device "+ModelNumber+" added successfully");
						
					}
					
					// Click on Expander node
					Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
					
					
					// Click on XLM Loop Card Expander
					repo.FormMe.BackplaneOrXLMExternalLoopCard_Expander.Click();
					
					// Click on XLM Loop C Node to add device
					repo.FormMe.XLMExternalLoopCardDevices_C.Click();
					
					float.TryParse(s40VLoad, out ZetfastFourtyVLoad);
					
					//Get 40V load from UI
					sDefault40V = Get40VPSULoadValue(PanelType);
					
					//Generate expected 40V load
					float.TryParse(sDefault40V, out Default40V);
					Expected40VPSU = Default40V+ZetfastFourtyVLoad;
					sExpected40VPSU= Expected40VPSU.ToString("0.000");
					
					// Verify 40V PSU load value on addition of zetfast loop with devices
					verify40VPSULoadValue(sExpected40VPSU,PanelType);
					
				}
				
				for(int k=9; k<=7; k--)
				{
					ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[k,7]).Value.ToString();
					sType = ((Range)Excel_Utilities.ExcelRange.Cells[k,8]).Value.ToString();
					sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[k,9]).Value.ToString();
					s40VLoad = ((Range)Excel_Utilities.ExcelRange.Cells[k,10]).Value.ToString();
					
					//Get 40V load from UI
					sDefault40V = Get40VPSULoadValue(PanelType);
					
					//Generate expected 40V load on Deletion
					float.TryParse(sDefault40V, out Default40V);
					float.TryParse(s40VLoad, out ZetfastFourtyVLoad);
					Expected40VPSU = Default40V-ZetfastFourtyVLoad;
					sExpected40VPSU = Expected40VPSU.ToString("0.000");
					
					if(k==8)
					{
						// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
						repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1.Click();
						
						if(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info.Exists())
						{
							repo.ProfileConsys1.btn_Delete.Click();
							Validate.AttributeEqual(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info, "Text", sLabelName);
							Report.Log(ReportLevel.Success, "Device "+sLabelName+" deleted successfully");
							
							// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
							// Verify 40V PSU load value on deletion of Zetfast loop
							verify40VPSULoadValue(sExpected40VPSU,PanelType);
						}
					}
					
					
					else
					{
						// Click on XLM Loop Card Expander
						repo.FormMe.BackplaneOrXLMExternalLoopCard_Expander.Click();
						
						// Click on XLM Loop C Node to add device
						repo.FormMe.XLMExternalLoopCardDevices_C.Click();

						repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1.Click();
						
						if(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info.Exists())
						{
							repo.ProfileConsys1.btn_Delete.Click();
							Validate.AttributeEqual(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info, "Text", sLabelName);
							Report.Log(ReportLevel.Success, "Device "+sLabelName+" deleted successfully");
							
							// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
							// Verify 40V PSU load value on deletion of Zetfast loop
							verify40VPSULoadValue(sExpected40VPSU,PanelType);
							
						}
						else
						{
							
							Report.Log(ReportLevel.Failure, "Device "+sLabelName+" not found");
						}
						
					}
				}
				
				
				
				
				
				// Delete panel using PanelNode details from excel sheet
				Panel_Functions.DeletePanel(1,PanelNode,1);
				
			}
			//Close opened excel sheet
			Excel_Utilities.CloseExcel();
			
		}
		
		
		
		/***********************************************************************************************************************************
		 * Function Name: verify40VLoadOnSlotCardsAddDelete
		 * Function Details: To Verify 40V load on addition/deletion of Slot Cards
		 * Parameter/Arguments:   expected Maximum value, panel type (FIM or PFI)  and row number is 13 by default for FIM
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update : 3/02/2019  Alpesh Dhakad - 31/07/2019 & 21/08/2019 - Updated test scripts as per new build and xpaths
		 ***********************************************************************************************************************************/
		[UserCodeMethod]
		public static void verify40VLoadOnSlotCardsAddDelete(string sFileName,string sAddPanelSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared variables
			string PanelName,PanelNode,s40VLoad,sDefault40V,CPUType,sRowNumber,sType,PanelType,sExpected40VPSU;
			int rowNumber;
			float AccessoryFourtyVLoad,Default40V;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sRowNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				
				
				int.TryParse(sRowNumber, out rowNumber);
				
				// Add panels using test data in excel sheet
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				Report.Log(ReportLevel.Info, "Panel "+PanelName+" added successfully");
				
				// Click on Expander node
					Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
					
					// Click on Loop Card node
					Common_Functions.ClickOnNavigationTreeExpander(PanelType);
					
					// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Main");
					
				
				
				//Click on Panel Accessories tab
				//repo.FormMe.tab_PanelAccessories.Click();
				
				for(int j=8; j<=9; j++)
				{
					
					ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[j,7]).Value.ToString();
					sType = ((Range)Excel_Utilities.ExcelRange.Cells[j,8]).Value.ToString();
					sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[j,9]).Value.ToString();
					s40VLoad = ((Range)Excel_Utilities.ExcelRange.Cells[j,10]).Value.ToString();
					
					// Click on Panel node
					Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
					//Click on Panel Accessories tab
					repo.FormMe.tab_PanelAccessories.Click();
					
					float.TryParse(s40VLoad, out AccessoryFourtyVLoad);
					Devices_Functions.AddDevicefromPanelAccessoriesGallery(ModelNumber,sType);
					
					// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
					//Get 40V load from UI
					sDefault40V = Get40VPSULoadValue(PanelType);
					
					//Generate expected 40V load
					float.TryParse(sDefault40V, out Default40V);
					float Expected40VPSU = Default40V+AccessoryFourtyVLoad;
					sExpected40VPSU= Expected40VPSU.ToString("0.000");
					
					// Verify 40V PSU load value on addition printer
					verify40VPSULoadValue(sExpected40VPSU,PanelType);
					
					//Get 40V load from UI
					sDefault40V = Get40VPSULoadValue(PanelType);
					
					//Generate expected 40V load on Deletion
					float.TryParse(sDefault40V, out Default40V);
					Expected40VPSU = Default40V-AccessoryFourtyVLoad;
					sExpected40VPSU = Expected40VPSU.ToString("0.000");
					
					// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
					//Click on Panel Accessories tab
					repo.FormMe.tab_PanelAccessories.Click();
					
					//repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1.Click();
					
					repo.ProfileConsys1.PanelInvetoryGrid.txt_LabelNameofAccessory.Click();
					
					if(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info.Exists())
					{
						repo.ProfileConsys1.btn_Delete.Click();
						//Validate.AttributeEqual(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info, "Text", sLabelName);
						Report.Log(ReportLevel.Success, "Device "+sLabelName+" deleted successfully");
						
						// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
						// Verify 40V PSU load value on deletion of Accessory
						verify40VPSULoadValue(sExpected40VPSU,PanelType);
					}
					
					else
					{
						
						Report.Log(ReportLevel.Failure, "Device "+sLabelName+" not found");
					}

					
				}
				
				
				// Delete panel using PanelNode details from excel sheet
				Panel_Functions.DeletePanel(1,PanelNode,1);
				
			}
			//Close opened excel sheet
			Excel_Utilities.CloseExcel();
			
		}

		/*******************************************************************************************************************************
		 * Function Name: verify40VCalculationforPFI
		 * Function Details: To Verify 40V load on addition/deletion of PLX loop card with devices
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update : 8/02/2019  Alpesh Dhakad - 31/07/2019 & 21/08/2019 - Updated test scripts as per new build and xpaths
		 *******************************************************************************************************************************/
		[UserCodeMethod]
		public static void verify40VCalculationforPFI(string sFileName,string sAddPanelSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared variables
			string PanelName,PanelNode,CPUType,sRowNumber,sType,PanelType,sExpected40VPSU,sIsPLXSupported;
			int rowNumber;
			bool isPLXSupported;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sRowNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				sExpected40VPSU = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				sIsPLXSupported = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				if (sIsPLXSupported.Equals("YES"))
				{
					isPLXSupported = true;
				}
				else
				{
					isPLXSupported = false;
				}
				int.TryParse(sRowNumber, out rowNumber);
				
				// Add panels using test data in excel sheet
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				
				// Click on Expander node
					Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
					
					// Click on Loop Card node
					Common_Functions.ClickOnNavigationTreeExpander(PanelType);
					
					// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					;
				
				// Verify 40V PSU load value of Built in PLX loop card
				verify40VPSULoadValue(sExpected40VPSU,PanelType);
				
				if(isPLXSupported)
				{
					// Add PLX loop and devices and verify 40 V load
					for(int j=8; j<9; j++)
					{
						
						// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
						//Fetch PLX card details
						ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[j,9]).Value.ToString();
						sType = ((Range)Excel_Utilities.ExcelRange.Cells[j,10]).Value.ToString();
						sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[j,11]).Value.ToString();
						sExpected40VPSU = ((Range)Excel_Utilities.ExcelRange.Cells[j,12]).Value.ToString();
						
						//Add External Loop card
						Devices_Functions.AddDevicesfromMainProcessorGallery(ModelNumber,sType,PanelType);
						
						// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
						// Click on Expander 
					Common_Functions.ClickOnNavigationTreeExpander("Backplane");
					
					// Click on PLX expander button
					Common_Functions.ClickOnNavigationTreeExpander("PLX");
					
					// Click on PLX800 node
			Common_Functions.ClickOnNavigationTreeItem("PLX800-E");
			
						// Click on PLX loop E
						repo.FormMe.PLX800LoopCard_E.Click();
						
						// Verify 40V PSU load value of Built in PLX loop card
						verify40VPSULoadValue(sExpected40VPSU,PanelType);
						
						//Delete External PLX loop card
						
						// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
						repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1.Click();
						
						if(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info.Exists())
						{
							repo.ProfileConsys1.btn_Delete.Click();
							Validate.AttributeEqual(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info, "Text", sLabelName);
							Report.Log(ReportLevel.Success, "Device "+sLabelName+" deleted successfully");
							
						}
					}
				}
				// Delete panel using PanelNode details from excel sheet
				Panel_Functions.DeletePanel(1,PanelNode,1);
				
			}
			//Close opened excel sheet
			Excel_Utilities.CloseExcel();
		}
		
		
		/*****************************************************************************************************************************
		 * Function Name: verify40VCalculationforFIM
		 * Function Details: To Verify 40V load on addition/deletion of XLM loop card
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update : 11/02/2019  Alpesh Dhakad - 31/07/2019 & 21/08/2019 - Updated test scripts as per new build and xpaths
		 *****************************************************************************************************************************/
		[UserCodeMethod]
		public static void verify40VCalculationforFIM(string sFileName,string sAddPanelSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared variables
			string PanelName,PanelNode,CPUType,sRowNumber,sType,PanelType,sXLMFortyVLoad,sIsXLMSupported,sDefault40V,sExpected40VPSU;
			int rowNumber;
			float Default40V,XLMFortyVLoad;
			bool isXLMSupported;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sRowNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				sDefault40V = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				sIsXLMSupported = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				if (sIsXLMSupported.Equals("YES"))
				{
					isXLMSupported = true;
				}
				else
				{
					isXLMSupported = false;
				}
				int.TryParse(sRowNumber, out rowNumber);
				
				// Add panels using test data in excel sheet
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				
				// Click on Expander node
					Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
					
					// Click on Loop Card node
					Common_Functions.ClickOnNavigationTreeExpander(PanelType);
					
					// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
				
				// Verify 40V PSU load value of Built in PLX loop card
				verify40VPSULoadValue(sDefault40V,PanelType);
				
				if( isXLMSupported)
				{
					// Add PLX loop and devices and verify 40 V load
					for(int j=8; j<9; j++)
					{
						
						// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
						//Fetch PLX card details
						ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[j,9]).Value.ToString();
						sType = ((Range)Excel_Utilities.ExcelRange.Cells[j,10]).Value.ToString();
						sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[j,11]).Value.ToString();
						sXLMFortyVLoad = ((Range)Excel_Utilities.ExcelRange.Cells[j,12]).Value.ToString();
						
						//Add External Loop card
						Devices_Functions.AddDevicesfromMainProcessorGallery(ModelNumber,sType,PanelType);
						Report.Log(ReportLevel.Info, "Device "+ModelNumber+" added successfully");
						
						// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
						// Expand Backplane node
						repo.FormMe.BackplaneOrXLMExternalLoopCard_Expander.Click();
						
						// Expand external loop card node
						repo.FormMe.BackplaneOrXLMExternalLoopCard_Expander.Click();
						
						// Click on PLX loop E
						repo.FormMe.XLMExternalLoopCardDevices_C.Click();
						
						//Generate expected 40V load
						float.TryParse(sDefault40V, out Default40V);
						float.TryParse(sXLMFortyVLoad, out XLMFortyVLoad);
						float Expected40VPSU = Default40V+XLMFortyVLoad;
						sExpected40VPSU= Expected40VPSU.ToString("0.000");
						
						// Verify 40V PSU load value of after addition of XLM loop card
						verify40VPSULoadValue(sExpected40VPSU,PanelType);
						
						//Generate expected 40V load on deletion
						float.TryParse(sDefault40V, out Default40V);
						float.TryParse(sXLMFortyVLoad, out XLMFortyVLoad);
						Expected40VPSU = Expected40VPSU-XLMFortyVLoad;
						sExpected40VPSU= Expected40VPSU.ToString("0.000");
						
						// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
						
						repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1.Click();
						
						if(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info.Exists())
						{
							repo.ProfileConsys1.btn_Delete.Click();
							Validate.AttributeEqual(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info, "Text", sLabelName);
							Report.Log(ReportLevel.Success, "Device "+sLabelName+" deleted successfully");
						}
						
						
						// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
						
						// Verify 40V PSU load value
						verify40VPSULoadValue(sExpected40VPSU,PanelType);
						
					}
				}
				// Delete panel using PanelNode details from excel sheet
				Panel_Functions.DeletePanel(1,PanelNode,1);
				
			}
			//Close opened excel sheet
			Excel_Utilities.CloseExcel();
		}

		
		/***************************************************************************************************************************
		 * Function Name: verify40VCalculationforPLXLoopWithDevices
		 * Function Details: To Verify 40V load on addition/deletion of PLX loop card with devices
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update : 13/02/2019   Alpesh Dhakad - 31/07/2019 & 21/08/2019 - Updated test scripts as per new build and xpaths
		 ***************************************************************************************************************************/
		[UserCodeMethod]
		public static void verify40VCalculationforPLXLoopWithDevices(string sFileName,string sAddPanelSheet)
		{
			//Open excel sheet and read it's values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared variables
			string PanelName,PanelNode,CPUType,sRowNumber,sType,PanelType,sExpected40VPSU,sIsPLXSupported,sLoopsSupported,sDefault40V,sExpected40VLoadofDevices;
			int rowNumber,iLoopsSupported,k;
			bool isPLXSupported;
			float Default40V,Expected40VLoadofDevices;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sRowNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				sDefault40V = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				sIsPLXSupported = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				sLoopsSupported = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				int.TryParse(sLoopsSupported,out iLoopsSupported);
				if (sIsPLXSupported.Equals("YES"))
				{
					isPLXSupported = true;
				}
				else
				{
					isPLXSupported = false;
				}
				int.TryParse(sRowNumber, out rowNumber);
				
				// Add panels using test data in excel sheet
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				
				// Click on Expander node
					Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
					
					// Click on Loop Card node
					Common_Functions.ClickOnNavigationTreeExpander(PanelType);
					
					// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
				
				// Verify 40V PSU load value of Built in PLX loop card
				verify40VPSULoadValue(sDefault40V,PanelType);
				
				if(isPLXSupported)
				{
					// Add PLX loop verify 40 V load
					for(int j=8; j<9; j++)
					{
						
							// Click on Panel node
						Common_Functions.ClickOnNavigationTreeItem(PanelNode);
				
						//Fetch PLX card details
						ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[j,10]).Value.ToString();
						sType = ((Range)Excel_Utilities.ExcelRange.Cells[j,11]).Value.ToString();
						sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[j,12]).Value.ToString();
						sExpected40VPSU = ((Range)Excel_Utilities.ExcelRange.Cells[j,13]).Value.ToString();
						
						//Add External Loop card
						Devices_Functions.AddDevicesfromMainProcessorGallery(ModelNumber,sType,PanelType);
							
						// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
						// Expand Backplane node
						repo.FormMe.BackplaneOrXLMExternalLoopCard_Expander.Click();
						
						// Expand external loop card node
						repo.FormMe.PLXExternalLoopCard_Expander.Click();
						
						// Click on PLX loop E
						repo.FormMe.PLX800LoopCard_E.Click();
						
						// Verify 40V PSU load value of loop card
						verify40VPSULoadValue(sExpected40VPSU,PanelType);
					}
					// 40 V load on Addition of devices
					sExpected40VLoadofDevices = ((Range)Excel_Utilities.ExcelRange.Cells[6,16]).Value.ToString();
					
					//Generate expected 40V load
					float.TryParse(sDefault40V, out Default40V);
					float.TryParse(sExpected40VLoadofDevices, out Expected40VLoadofDevices);
					float Expected40VPSU = Default40V+Expected40VLoadofDevices;
					sExpected40VPSU= Expected40VPSU.ToString("0.000");
					
					// Select Loop E and Add devices
					repo.FormMe.PLX800LoopCard_E.Click();
					
					for(k=8;k<=9;k++)
					{
						// Fetch devices data and add devices in PLX loop card
						ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[k,15]).Value.ToString();
						sType = ((Range)Excel_Utilities.ExcelRange.Cells[k,16]).Value.ToString();
						Devices_Functions.AddDevicesfromMainProcessorGallery(ModelNumber,sType,PanelType);
					}
					
					// Verify 40V PSU load value of loop after addition of devices
					verify40VPSULoadValue(sExpected40VPSU,PanelType);
					
					// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
					for(k=8;k<=9;k++)
					{
						// Fetch devices data and add devices in PLX loop card
						ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[k,15]).Value.ToString();
						sType = ((Range)Excel_Utilities.ExcelRange.Cells[k,16]).Value.ToString();
						Devices_Functions.AddDevicesfromMainProcessorGallery(ModelNumber,sType,PanelType);
					}
					
					// Verify 40V PSU load value of loop after addition of devices
					verify40VPSULoadValue(sExpected40VPSU,PanelType);
				}
				
				else
				{
					// 40 V load on Addition of devices
					sExpected40VLoadofDevices = ((Range)Excel_Utilities.ExcelRange.Cells[6,16]).Value.ToString();
					
					//Generate expected 40V load
					float.TryParse(sDefault40V, out Default40V);
					float.TryParse(sExpected40VLoadofDevices, out Expected40VLoadofDevices);
					float Expected40VPSU = Default40V+Expected40VLoadofDevices;
					sExpected40VPSU= Expected40VPSU.ToString("0.000");
					
					// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
					for(k=8;k<=9;k++)
					{
						// Fetch devices data and add devices in PLX loop card
						ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[k,15]).Value.ToString();
						sType = ((Range)Excel_Utilities.ExcelRange.Cells[k,16]).Value.ToString();
						Devices_Functions.AddDevicesfromMainProcessorGallery(ModelNumber,sType,PanelType);
					}
					
					// Verify 40V PSU load value of loop after addition of devices
					verify40VPSULoadValue(sExpected40VPSU,PanelType);
					
				}
				// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
				repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1.Click();
				
				if(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info.Exists())
				{
					repo.ProfileConsys1.btn_Delete.Click();
					Validate.AttributeEqual(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info, "Text", sLabelName);
					Report.Log(ReportLevel.Success, "Device "+sLabelName+" deleted successfully");
					
				}
				
				
				// Delete panel using PanelNode details from excel sheet
				Panel_Functions.DeletePanel(1,PanelNode,1);
				
			}
			//Close opened excel sheet
			Excel_Utilities.CloseExcel();
		}
		
		
		/*****************************************************************************************************************************
		 * Function Name: verify40VCalculationforXLMLoopWithDevices
		 * Function Details: To Verify 40V load on addition/deletion of XLM loop card with devices
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update : 13/02/2019  Alpesh Dhakad - 31/07/2019 & 21/08/2019- Updated test scripts as per new build and xpaths
		 *****************************************************************************************************************************/
		[UserCodeMethod]
		public static void verify40VCalculationforXLMLoopWithDevices(string sFileName,string sAddPanelSheet)
		{
			//Open excel sheet and read it's values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared variables
			string PanelName,PanelNode,CPUType,sRowNumber,sType,PanelType,sExpected40VPSU,sIsXLMSupported,sCalcExpected40VPSU,sDefault40V,sExpected40VLoadofDevices;
			int rowNumber,k;
			bool isXLMSupported;
			float Default40V,Expected40VLoadofDevices,Expected40VPSU,CalcExpected40VPSU;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sRowNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				sDefault40V = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				sIsXLMSupported = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				
				if (sIsXLMSupported.Equals("YES"))
				{
					isXLMSupported = true;
				}
				else
				{
					isXLMSupported = false;
				}
				int.TryParse(sRowNumber, out rowNumber);
				
				// Add panels using test data in excel sheet
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				Report.Log(ReportLevel.Info, "Panel "+PanelName+" added successfully");
				
				// Click on Expander node
				Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
				
				// Click on Loop Card node
				Common_Functions.ClickOnNavigationTreeExpander(PanelType);
				
				// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				
				
				// Verify 40V PSU load value of Built in XLM loop card
				verify40VPSULoadValue(sDefault40V,PanelType);
				
				if(isXLMSupported)
				{
					// Add XLM loop verify 40 V load
					for(int j=8; j<9; j++)
					{
						
						// Click on Panel node
						Common_Functions.ClickOnNavigationTreeItem(PanelNode);
						
						//Fetch XLM card details
						ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[j,9]).Value.ToString();
						sType = ((Range)Excel_Utilities.ExcelRange.Cells[j,10]).Value.ToString();
						sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[j,11]).Value.ToString();
						sExpected40VPSU = ((Range)Excel_Utilities.ExcelRange.Cells[j,12]).Value.ToString();
						
						//Add External Loop card
						Devices_Functions.AddDevicesfromMainProcessorGallery(ModelNumber,sType,PanelType);
						
						// Click on Panel node
						Common_Functions.ClickOnNavigationTreeItem(PanelNode);
						
						//Expand Backplane node
						Common_Functions.ClickOnNavigationTreeExpander("XLM/External");
						
						// Expand Backplane node
						//repo.FormMe.BackplaneOrXLMExternalLoopCard_Expander.Click();
						
						// Expand external loop card node
						//repo.FormMe.BackplaneOrXLMExternalLoopCard_Expander.Click();
						
						// Click on Loop A node
						Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
						
						//Generate expected 40V load
						float.TryParse(sDefault40V, out Default40V);
						float.TryParse(sExpected40VPSU, out Expected40VPSU);
						CalcExpected40VPSU = Default40V+Expected40VPSU;
						sCalcExpected40VPSU= CalcExpected40VPSU.ToString("0.000");
						
						// Verify 40V PSU load value of loop card
						verify40VPSULoadValue(sCalcExpected40VPSU,PanelType);
						
						
						// Click on Loop C
						Common_Functions.ClickOnNavigationTreeItem("XLM800-C");
						
						//repo.FormMe.XLMExternalLoopCardDevices_C.Click();
						
						// Verify 40V PSU load value of loop card
						verify40VPSULoadValue(sCalcExpected40VPSU,PanelType);
						
						// 40 V load on Addition of devices
						sExpected40VLoadofDevices = ((Range)Excel_Utilities.ExcelRange.Cells[6,15]).Value.ToString();
						
						//Generate expected 40V load
						float.TryParse(sDefault40V, out Default40V);
						float.TryParse(sExpected40VPSU, out Expected40VPSU);
						float.TryParse(sExpected40VLoadofDevices, out Expected40VLoadofDevices);
						CalcExpected40VPSU = Default40V+Expected40VPSU+Expected40VLoadofDevices;
						sCalcExpected40VPSU= CalcExpected40VPSU.ToString("0.000");
						
						// Click on Loop C and add devices
						Common_Functions.ClickOnNavigationTreeItem("XLM800-C");
						
						//repo.FormMe.XLMExternalLoopCardDevices_C.Click();
						
						for(k=8;k<=9;k++)
						{
							// Fetch devices data and add devices in XLM loop card
							ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[k,14]).Value.ToString();
							sType = ((Range)Excel_Utilities.ExcelRange.Cells[k,15]).Value.ToString();
							Devices_Functions.AddDevicesfromMainProcessorGallery(ModelNumber,sType,PanelType);
						}
						
						// Verify 40V PSU load value of loop after addition of devices
						verify40VPSULoadValue(sCalcExpected40VPSU,PanelType);
						
						// Click on Loop A node
						Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
						
						//Generate expected 40V load
						float.TryParse(sDefault40V, out Default40V);
						float.TryParse(sExpected40VPSU, out Expected40VPSU);
						float.TryParse(sExpected40VLoadofDevices, out Expected40VLoadofDevices);
						CalcExpected40VPSU = Default40V+Expected40VPSU+Expected40VLoadofDevices+Expected40VLoadofDevices;
						sCalcExpected40VPSU= CalcExpected40VPSU.ToString("0.000");
						
						for(k=8;k<=9;k++)
						{
							// Fetch devices data and add devices in XLM loop card
							ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[k,14]).Value.ToString();
							sType = ((Range)Excel_Utilities.ExcelRange.Cells[k,15]).Value.ToString();
							Devices_Functions.AddDevicesfromMainProcessorGallery(ModelNumber,sType,PanelType);
						}
						
						// Verify 40V PSU load value of loop after addition of devices
						verify40VPSULoadValue(sCalcExpected40VPSU,PanelType);
						
					}
				}
				
				else
				{
					// 40 V load on Addition of devices
					sExpected40VLoadofDevices = ((Range)Excel_Utilities.ExcelRange.Cells[6,15]).Value.ToString();
					
					//Generate expected 40V load
					float.TryParse(sDefault40V, out Default40V);
					float.TryParse(sExpected40VLoadofDevices, out Expected40VLoadofDevices);
					CalcExpected40VPSU = Default40V+Expected40VLoadofDevices;
					sCalcExpected40VPSU= CalcExpected40VPSU.ToString("0.000");
					
					// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
					for(k=8;k<=9;k++)
					{
						// Fetch devices data and add devices in XLM loop card
						ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[k,14]).Value.ToString();
						sType = ((Range)Excel_Utilities.ExcelRange.Cells[k,15]).Value.ToString();
						Devices_Functions.AddDevicesfromMainProcessorGallery(ModelNumber,sType,PanelType);
					}
					
					// Verify 40V PSU load value of loop after addition of devices
					verify40VPSULoadValue(sCalcExpected40VPSU,PanelType);
					
				}
				
				// Delete panel using PanelNode details from excel sheet
				Panel_Functions.DeletePanel(1,PanelNode,1);
				
			}
			//Close opened excel sheet
			Excel_Utilities.CloseExcel();
		}
		
		
		
		/*****************************************************************************************************************************
		 * Function Name:verifyMaxBatteryStandbyAndAlarmLoad
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner:Purvi Bhasin
		 * Last Update :4/2/2019  Alpesh Dhakad - 30/07/2019 & 23/08/2019 - Updated test scripts as per new build and xpaths
		 *****************************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyMaxBatteryStandbyAndAlarmLoad(string sFileName,string sAddPanelSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared string typ
			string PanelName,PanelNode,CPUType,sRowNumber,PanelType,expectedMaxBatteryStandby,expectedMaxAlarmLoad;
			int rowNumber;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sRowNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				expectedMaxBatteryStandby = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				expectedMaxAlarmLoad = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				
				int.TryParse(sRowNumber, out rowNumber);
				
				// Add panels using test data in excel sheet
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				Report.Log(ReportLevel.Info, "Panel "+PanelName+" added successfully");

				// Click on Expander node
					Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
					
					// Click on Loop Card node
					Common_Functions.ClickOnNavigationTreeExpander(PanelType);
					
					// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
				
				// Verify max Battery Standby load value
				verifyMaxBatteryStandby(expectedMaxBatteryStandby,false);
				
				// Verify max Alarm load value
				verifyMaxAlarmLoad(expectedMaxAlarmLoad,false);
				
				// Click on Site node
			Common_Functions.ClickOnNavigationTreeItem("Site");
			
				// Delete panel using PanelNode details from excel sheet
				Panel_Functions.DeletePanel(1,PanelNode,1);
				
			}
			//Close opened excel sheet
			Excel_Utilities.CloseExcel();
			
		}
		
		/*****************************************************************************************************************
		 * Function Name: verifyMaxBatteryStandby
		 * Function Details:
		 * Parameter/Arguments:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 4/2/2019
		 *****************************************************************************************************************/
		public static void verifyMaxBatteryStandby(string expectedMaxBatteryStandby,bool isSecondPSU)
		{
			if(isSecondPSU)
			{
				sRow=(18).ToString();
			}
			else
			{
				sRow=(16).ToString();
			}
			
			// Click on Physical layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			// Fetch Battery Standby maximum limit value
			string maxBatteryStandby = repo.ProfileConsys1.MaxBatteryStandby.TextValue;
			
			// Compare max40VPSU value with expected value
			if(maxBatteryStandby.Equals(expectedMaxBatteryStandby))
			{
				Report.Log(ReportLevel.Success,"Max Battery Standby " + maxBatteryStandby + " is displayed correctly " );
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Max Max Battery Standby value is not displayed correctly, it is displayed as: " + maxBatteryStandby + " instead of : " +expectedMaxBatteryStandby);
			}
		}
		
		/*****************************************************************************************************************
		 * Function Name: verifyDefaultBatteryStandby
		 * Function Details:
		 * Parameter/Arguments:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 4/2/2019
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyBatteryStandby(string expectedBatteryStandby, bool isSecondPSU, string PanelType)
		{
			
			if(PanelType.Equals("FIM"))
			{
				sCell= "[4]";
				if(isSecondPSU)
				{
					sRow=(18).ToString();
				}
				else
				{
					sRow=(16).ToString();
				}
				
			}
			else
			{
				sCell= "[3]";
				sRow=(16).ToString();
			}
			// Click on Physical layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			// Fetch Default Battery Standby limit value
			string BatteryStandby = repo.FormMe.BatteryStandBy.TextValue;
			
			// Compare Default Battery Standby value with expected value
			if(BatteryStandby.Equals(expectedBatteryStandby))
			{
				Report.Log(ReportLevel.Success,"Battery Standby " + BatteryStandby + " is displayed correctly " );
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Battery Standby value is not displayed correctly, it is displayed as: " + BatteryStandby + " instead of : " +expectedBatteryStandby);
			}
		}
		
		/*****************************************************************************************************************
		 * Function Name: verifyBatteryStandbyAccToRow
		 * Function Details:
		 * Parameter/Arguments:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 4/2/2019
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyBatteryStandbyAccToRow(string expectedBatteryStandby, string rowNum, string PanelType)
		{
			
			sRow = rowNum; //rowNum should be present in Excell acc to the number of isolator Devices added
			if(PanelType.Equals("FIM"))
			{
				sCell= "[4]";
			}
			else
			{
				sCell= "[3]";
			}
			
			
			// Click on Physical layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			// Fetch Default Battery Standby limit value
			string BatteryStandby = repo.FormMe.BatteryStandBy.TextValue;
			
			// Compare Default Battery Standby value with expected value
			if(BatteryStandby.Equals(expectedBatteryStandby))
			{
				Report.Log(ReportLevel.Success,"Battery Standby " + BatteryStandby + " is displayed correctly " );
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Battery Standby value is not displayed correctly, it is displayed as: " + BatteryStandby + " instead of : " +expectedBatteryStandby);
			}
		}
		
		/*****************************************************************************************************************
		 * Function Name: verifyBatteryStandbyOnChangingCPU
		 * Function Details:
		 * Parameter/Arguments:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 4/2/2019
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyBatteryStandbyOnChangingCPU(string expectedBatteryStandby)
		{
			// Click on Physical layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			// Fetch Battery Standby limit value
			string BatteryStandby = repo.FormMe.BatteryStandBy.TextValue;
			
			// Compare Battery Standby value with expected value
			if(BatteryStandby.Equals(expectedBatteryStandby))
			{
				Report.Log(ReportLevel.Success," Battery Standby " + BatteryStandby + " is displayed correctly " );
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Default Battery Standby value is not displayed correctly, it is displayed as: " + BatteryStandby + " instead of : " +expectedBatteryStandby);
			}
		}

		/*****************************************************************************************************************
		 * Function Name: verifyMaxAlarmLoad
		 * Function Details:
		 * Parameter/Arguments:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 4/2/2019
		 *****************************************************************************************************************/
		public static void verifyMaxAlarmLoad(string expectedMaxAlarmLoad, bool isSecondPSU)
		{
			if(isSecondPSU)
			{
				sRow=(19).ToString();
			}
			else
			{
				sRow=(17).ToString();
			}
			// Fetch Battery Standby maximum limit value
			string maxAlarmLoad = repo.ProfileConsys1.MaxAlarmLoad.TextValue;
			
			// Compare max40VPSU value with expected value
			if(maxAlarmLoad.Equals(expectedMaxAlarmLoad))
			{
				Report.Log(ReportLevel.Success,"Max Alarm Load " + maxAlarmLoad + " is displayed correctly " );
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Max Alarm Load value is not displayed correctly, it is displayed as: " + maxAlarmLoad + " instead of : " +expectedMaxAlarmLoad);
			}
			
		}
		
		/*****************************************************************************************************************
		 * Function Name: verifyAlarmLoad
		 * Function Details:
		 * Parameter/Arguments:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 4/2/2019
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyAlarmLoad(string expectedAlarmLoad, bool isSecondPSU, string PanelType)
		{
			
			if(PanelType.Equals("FIM"))
			{
				sCell= "[5]";
				if(isSecondPSU)
				{
					sRow=(19).ToString();
				}
				else
				{
					sRow=(17).ToString();
				}
			}
			
			else
			{
				sCell= "[4]";
				sRow=(17).ToString();
			}
			// Click on Physical layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			// Fetch Default Alarm Load limit value
			string AlarmLoad = repo.FormMe.AlarmLoad.TextValue;
			
			// Compare Default Alarm Load value with expected value
			if(AlarmLoad.Equals(expectedAlarmLoad))
			{
				Report.Log(ReportLevel.Success,"Alarm Load " + AlarmLoad + " is displayed correctly " );
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Alarm Load value is not displayed correctly, it is displayed as: " + AlarmLoad + " instead of : " +expectedAlarmLoad);
			}
		}

		/*****************************************************************************************************************
		 * Function Name: verifyAlarmLoadOnChangingCPU
		 * Function Details:
		 * Parameter/Arguments:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 4/2/2019
		 *****************************************************************************************************************/
		public static void verifyAlarmLoadOnChangingCPU(string expectedAlarmLoad)
		{
			// Fetch Default Alarm Load
			string AlarmLoad = repo.FormMe.AlarmLoad.TextValue;
			
			// Compare Default Alarm Load value with expected value
			if(AlarmLoad.Equals(expectedAlarmLoad))
			{
				Report.Log(ReportLevel.Success,"Default Alarm Load " + AlarmLoad + " is displayed correctly " );
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Max Alarm Load value is not displayed correctly, it is displayed as: " + AlarmLoad + " instead of : " +expectedAlarmLoad);
			}

		}
		
		/*****************************************************************************************************************
		 * Function Name: verifyAlarmLoadAccToRow
		 * Function Details:
		 * Parameter/Arguments:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 4/2/2019
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyAlarmLoadAccToRow(string expectedAlarmLoad, string rowNum, string PanelType)
		{
			
			sRow = rowNum; //rowNum should be present in Excell acc to the number of isolator Devices added
			if(PanelType.Equals("FIM"))
			{
				sCell= "[5]";
			}
			else
			{
				sCell= "[4]";
			}
			
			// Click on Physical layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			// Fetch Default Alarm Load limit value
			string AlarmLoad = repo.FormMe.AlarmLoad.TextValue;
			
			// Compare Default Alarm Load value with expected value
			if(AlarmLoad.Equals(expectedAlarmLoad))
			{
				Report.Log(ReportLevel.Success,"Alarm Load " + AlarmLoad + " is displayed correctly " );
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Alarm Load value is not displayed correctly, it is displayed as: " + AlarmLoad + " instead of : " +expectedAlarmLoad);
			}
		}

		/*****************************************************************************************************************
		 * Function Name:verifyBatteryStandbyAndAlarmLoadOnChangingCPUAndPSU
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner:Purvi Bhasin
		 * Last Update :4/2/2019  Alpesh Dhakad - 30/07/2019 - Updated test scripts as per new build and xpaths
		 * Alpesh Dhakad - 08/08/2019 - Updated code from node expander to panel node
		 * Alpesh Dhakad - 21/08/2019 - Updated with new navigation tree method, xpath and devices gallery 
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyBatteryStandbyAndAlarmLoadOnChangingCPUAndPSU(string sFileName,string sAddPanelSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared string type
			string PanelName,PanelNode,CPUType,sRowNumber,changeCPUType,PanelType,expectedBatteryStandby,expectedDefaultBatteryStandby,expectedAlarmLoad,expectedDefaultAlarmLoad,changePSUType;
			int rowNumber;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sRowNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				changeCPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				changePSUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				expectedDefaultBatteryStandby = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				expectedDefaultAlarmLoad = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				expectedBatteryStandby = ((Range)Excel_Utilities.ExcelRange.Cells[i,10]).Value.ToString();
				expectedAlarmLoad = ((Range)Excel_Utilities.ExcelRange.Cells[i,11]).Value.ToString();
				// sPSUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,10]).Value.ToString();
				
				int.TryParse(sRowNumber, out rowNumber);
				
				// Add panels using test data in excel sheet
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				Report.Log(ReportLevel.Info, "Panel "+PanelName+" added successfully");

				// Click on Expander node
				Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
				
				// Click on Loop Card node
				Common_Functions.ClickOnNavigationTreeExpander(PanelType);
				
				// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				// Verify Default Battery Standby load value
				verifyBatteryStandby(expectedDefaultBatteryStandby,false,PanelType);
				
				// Verify Default Alarm load value
				verifyAlarmLoad(expectedDefaultAlarmLoad,false,PanelType);
				
				// Click on Site node
				Common_Functions.ClickOnNavigationTreeItem("Site");
			
				// Click on Panel node
				Common_Functions.ClickOnNavigationTreeItem(PanelNode);
				
				// Change CPU Type as per test data in sheet
				if (!changeCPUType.IsEmpty())
				{
					Panel_Functions.ChangeCPUType(changeCPUType);
				}
				
				//Change PSU of panel
				if (!changePSUType.IsEmpty())
				{
					Panel_Functions.ChangePSUType(changePSUType);
				}
				
				// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				// Verify Battery Standby on changing CPU load value
				verifyBatteryStandbyOnChangingCPU(expectedBatteryStandby);
				
				// Verify Alarm Load on changing CPU load value
				verifyAlarmLoadOnChangingCPU(expectedAlarmLoad);

				// Click on Site node
				Common_Functions.ClickOnNavigationTreeItem("Site");
			
				// Delete panel using PanelNode details from excel sheet
				Panel_Functions.DeletePanel(1,PanelNode,1);
				
			}
			//Close opened excel sheet
			Excel_Utilities.CloseExcel();
			
		}
		
		/*****************************************************************************************************************
		 * Function Name: verifyBatteryStandbyAndAlarmLoadOnEthernetAddDelete
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Purvi Bhasim
		 * Last Update : 08/01/2019 Alpesh Dhakad - 30/07/2019 - Updated test scripts as per new build and xpaths
		 * Alpesh Dhakad - 21/08/2019 & 30/08/2019 - Updated with new navigation tree method, xpath and devices gallery 
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyBatteryStandbyAndAlarmLoadOnEthernetAddDelete(string sFileName,string sAddPanelSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared variables
			string PanelName,PanelNode,sBatteryStandby,sAlarmLoad,expectedDefaultBatteryStandby,expectedDefaultAlarmLoad,CPUType,sRowNumber,sType,PanelType,expectedBatteryStandyby,expectedAlarmLoad,sDefaultBatteryStandyby,sDefaultAlarmLoad;
			int rowNumber;
			float BatteryStandby,AlarmLoad,DefaultBatteryStandby,DefaultAlarmLoad;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sRowNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				expectedDefaultBatteryStandby=((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				expectedDefaultAlarmLoad=((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				
				int.TryParse(sRowNumber, out rowNumber);
				
				// Add panels using test data in excel sheet
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				
				
				// Click on Expander node
				Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
				
				// Click on Loop Card node
				Common_Functions.ClickOnNavigationTreeExpander(PanelType);
				
				// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				
				// Verify Default Battery Standby load value
				verifyBatteryStandby(expectedDefaultBatteryStandby,false,PanelType);
				
				// Verify Default Alarm load value
				verifyAlarmLoad(expectedDefaultAlarmLoad,false,PanelType);
				
				// Click on Site node
			Common_Functions.ClickOnNavigationTreeItem("Site");
			
			// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
			// Click on Main processor expander node
				Common_Functions.ClickOnNavigationTreeExpander("Main");
					
			// Click on Ethernet node
			Common_Functions.ClickOnNavigationTreeItem("Ethernet");
			
				for(int j=8; j<=9; j++)
				{
					
					ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[j,8]).Value.ToString();
					sType = ((Range)Excel_Utilities.ExcelRange.Cells[j,9]).Value.ToString();
					sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[j,10]).Value.ToString();
					sBatteryStandby = ((Range)Excel_Utilities.ExcelRange.Cells[j,11]).Value.ToString();
					sAlarmLoad = ((Range)Excel_Utilities.ExcelRange.Cells[j,12]).Value.ToString();
					// Click on Ethernet node
			Common_Functions.ClickOnNavigationTreeItem("Ethernet");
			
					float.TryParse(sBatteryStandby, out BatteryStandby);
					float.TryParse(sAlarmLoad, out AlarmLoad);
					Devices_Functions.AddDevicesfromMainProcessorGallery(ModelNumber,sType,PanelType);
					Report.Log(ReportLevel.Info, "Device "+ModelNumber+" added successfully");
					
					// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
					//Get Battery Standby from UI
					sDefaultBatteryStandyby = GetBatteryStandbyValue(PanelType);
					sDefaultAlarmLoad = GetAlarmLoadValue(PanelType);
					
					//Generate expected Battery Standby and alarm load
					float.TryParse(sDefaultBatteryStandyby, out DefaultBatteryStandby);
					float.TryParse(sDefaultAlarmLoad, out DefaultAlarmLoad);
					float ExpectedBatteryStandby = DefaultBatteryStandby+BatteryStandby;
					float ExpectedAlarmLoad = DefaultAlarmLoad+AlarmLoad;
					expectedBatteryStandyby= ExpectedBatteryStandby.ToString("0.000");
					expectedAlarmLoad= ExpectedAlarmLoad.ToString("0.000");

					
					// Verify Battery Standby and alarm load value on addition of Ethernet
					verifyBatteryStandby(expectedBatteryStandyby,false,PanelType);
					verifyAlarmLoad(expectedAlarmLoad,false,PanelType);
					
					//Get Battery Standby and alarm load from UI
					sDefaultBatteryStandyby = GetBatteryStandbyValue(PanelType);
					sDefaultAlarmLoad = GetAlarmLoadValue(PanelType);
					
					//Generate expected Battery Standby and alarm load on Deletion
					float.TryParse(sDefaultBatteryStandyby, out DefaultBatteryStandby);
					float.TryParse(sDefaultAlarmLoad, out DefaultAlarmLoad);
					ExpectedBatteryStandby = DefaultBatteryStandby-BatteryStandby;
					ExpectedAlarmLoad = DefaultAlarmLoad-AlarmLoad;
					expectedBatteryStandyby = ExpectedBatteryStandby.ToString("0.000");
					expectedAlarmLoad = ExpectedAlarmLoad.ToString("0.000");
					
					// Click on Ethernet node
					Common_Functions.ClickOnNavigationTreeItem("Ethernet");
			
					repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1.Click();
					
					if(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info.Exists())
					{
						repo.ProfileConsys1.btn_Delete.Click();
						Validate.AttributeEqual(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info, "Text", sLabelName);
						Report.Log(ReportLevel.Success, "Device "+sLabelName+" deleted successfully");
						
						/// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
						// Verify Battery Standby and alarm load PSU load value on addition of Ethernet
						verifyBatteryStandby(expectedBatteryStandyby,false,PanelType);
						verifyAlarmLoad(expectedAlarmLoad,false,PanelType);
					}
					
					else
					{
						
						Report.Log(ReportLevel.Failure, "Device "+sLabelName+" not found");
					}

					
				}
				
				
				// Delete panel using PanelNode details from excel sheet
				Panel_Functions.DeletePanel(1,PanelNode,1);
				
			}
			//Close opened excel sheet
			Excel_Utilities.CloseExcel();
			
		}

		
		/*****************************************************************************************************************
		 * Function Name: verifyBatteryStandbyAndAlarmLoadOnRbusAddDelete
		 * Function Details: To Verify Battery Standby and Alarm Load on addition/deletion of R-Bus connection
		 * Parameter/Arguments:   expected Maximum value, panel type (FIM or PFI)  and row number is 16 and 17 by default
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 22/01/2019  Alpesh Dhakad - 30/07/2019 - Updated test scripts as per new build and xpaths
		 * Alpesh Dhakad - 07/09/2019 - Updated test scripts 
		 * Alpesh Dhakad - 21/08/2019 - Updated with new navigation tree method, xpath and devices gallery 
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyBatteryStandbyAndAlarmLoadOnRbusAddDelete(string sFileName,string sAddPanelSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared variables
			string PanelName,PanelNode,sBatteryStandby,sAlarmLoad,sDefaultBatteryStandby,sDefaultAlarmLoad,CPUType,sRowNumber,sType,PanelType,sExpectedBatteryStandby,sExpectedAlarmLoad,sXBusBatteryStandby,sXBusAlarmLoad,expectedDefaultBatteryStandby,expectedDefaultAlarmLoad;
			int rowNumber;
			float RBusBatteryStandby,RBusAlarmLoad,DefaultBatteryStandby,DefaultAlarmLoad,XBusBatteryStandby,XBusAlarmLoad;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sRowNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				expectedDefaultBatteryStandby = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				expectedDefaultAlarmLoad = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				
				int.TryParse(sRowNumber, out rowNumber);
				
				// Add panels using test data in excel sheet
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				
				// Click on Expander node
				Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
				
				// Click on Loop Card node
				Common_Functions.ClickOnNavigationTreeExpander(PanelType);
				
				// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				// Verify Default Battery Standby load value
				verifyBatteryStandby(expectedDefaultBatteryStandby,false,PanelType);
				
				// Verify Default Alarm load value
				verifyAlarmLoad(expectedDefaultAlarmLoad,false,PanelType);
				
				
				// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
			
			// Click on Main processor expander node
				Common_Functions.ClickOnNavigationTreeExpander("Main");
			
				
			// Click on RBus node
			Common_Functions.ClickOnNavigationTreeItem("R-BUS");
				
				
				for(int j=8; j<9; j++)
				{
					
					ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[j,9]).Value.ToString();
					sType = ((Range)Excel_Utilities.ExcelRange.Cells[j,10]).Value.ToString();
					sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[j,11]).Value.ToString();
					sBatteryStandby = ((Range)Excel_Utilities.ExcelRange.Cells[j,12]).Value.ToString();
					sAlarmLoad = ((Range)Excel_Utilities.ExcelRange.Cells[j,13]).Value.ToString();
					
					//Add RBus connection
					repo.FormMe.RBus1.Click();
					
					float.TryParse(sBatteryStandby, out RBusBatteryStandby);
					float.TryParse(sAlarmLoad, out RBusAlarmLoad);
					Devices_Functions.AddDevicesfromMainProcessorGallery(ModelNumber,sType,PanelType);
					Report.Log(ReportLevel.Info, "Device "+ModelNumber+" added successfully");
					
					
					// Add X-Bus to R-Bus
					ModelNumber = ((Range)Excel_Utilities.ExcelRange.Cells[j,14]).Value.ToString();
					sType = ((Range)Excel_Utilities.ExcelRange.Cells[j,15]).Value.ToString();
					//s40VLoad = ((Range)Excel_Utilities.ExcelRange.Cells[j,16]).Value.ToString();
					sXBusBatteryStandby = ((Range)Excel_Utilities.ExcelRange.Cells[j,17]).Value.ToString();
					sXBusAlarmLoad = ((Range)Excel_Utilities.ExcelRange.Cells[j,18]).Value.ToString();
					
					//Select R-Bus node
					repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1.Click();
					float.TryParse(sXBusBatteryStandby, out XBusBatteryStandby);
					float.TryParse(sXBusAlarmLoad, out XBusAlarmLoad);
					Devices_Functions.AddDevicesfromMainProcessorGallery(ModelNumber,sType,PanelType);
					Report.Log(ReportLevel.Info, "Device "+ModelNumber+" added successfully");
					
					// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
					//Get Battery Standby from UI
					sDefaultBatteryStandby = GetBatteryStandbyValue(PanelType);
					
					//Get Alarm Load from UI
					sDefaultAlarmLoad = GetAlarmLoadValue(PanelType);
					
					//Generate expected Battery Standby load
					float.TryParse(sDefaultBatteryStandby, out DefaultBatteryStandby);
					float ExpectedBatteryStandby = DefaultBatteryStandby+RBusBatteryStandby+XBusBatteryStandby;
					sExpectedBatteryStandby= ExpectedBatteryStandby.ToString("0.000");
					
					//Generate expected Alarm load
					float.TryParse(sDefaultAlarmLoad, out DefaultAlarmLoad);
					float ExpectedAlarmLoad = DefaultAlarmLoad+RBusAlarmLoad+XBusAlarmLoad;
					sExpectedAlarmLoad= ExpectedAlarmLoad.ToString("0.000");
					
					// Verify Battery Standby value on addition of R-Bus & X-Bus template
					verifyBatteryStandby(sExpectedBatteryStandby,false,PanelType);
					
					// Verify Alarm load value on addition of R-Bus & X-Bus template
					verifyAlarmLoad(sExpectedAlarmLoad,false,PanelType);
					
					//Get Battery Standby from UI
					sDefaultBatteryStandby = GetBatteryStandbyValue(PanelType);
					
					//Get Alarm load from UI
					sDefaultAlarmLoad = GetAlarmLoadValue(PanelType);
					
					//Generate expected Battery Standby load on Deletion
					float.TryParse(sDefaultBatteryStandby, out DefaultBatteryStandby);
					ExpectedBatteryStandby = DefaultBatteryStandby-RBusBatteryStandby-XBusBatteryStandby;
					sExpectedBatteryStandby = ExpectedBatteryStandby.ToString("0.000");
					
					//Generate expected Alarm load on Deletion
					float.TryParse(sDefaultAlarmLoad, out DefaultAlarmLoad);
					ExpectedAlarmLoad = DefaultAlarmLoad-RBusAlarmLoad-XBusAlarmLoad;
					sExpectedAlarmLoad = ExpectedAlarmLoad.ToString("0.000");
					
					// Click on RBus node
				Common_Functions.ClickOnNavigationTreeItem("R-BUS");
				
					repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1.Click();
					
					if(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info.Exists())
					{
						repo.ProfileConsys1.btn_Delete.Click();
						Validate.AttributeEqual(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info, "Text", sLabelName);
						Report.Log(ReportLevel.Success, "Device "+sLabelName+" deleted successfully");
						
						/// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
						// Verify Battery Standby and Alarm load value on addition of Ethernet
						verifyBatteryStandby(sExpectedBatteryStandby,false,PanelType);
						verifyAlarmLoad(sExpectedAlarmLoad,false,PanelType);
					}
					
					else
					{
						
						Report.Log(ReportLevel.Failure, "Device "+sLabelName+" not found");
					}

					
				}
				
				
				// Delete panel using PanelNode details from excel sheet
				Panel_Functions.DeletePanel(1,PanelNode,1);
				
			}
			//Close opened excel sheet
			Excel_Utilities.CloseExcel();
			
		}

		
		/*****************************************************************************************************************
		 * Function Name: GetBatteryStandbyValue
		 * Function Details: To get Battery Standby value
		 * Parameter/Arguments:   expected value, panel type (FIM or PFI)
		 * Output:40V load displayed on UI
		 * Function Owner: Purvi Bhasin
		 * Last Update : 22/01/2019
		 *****************************************************************************************************************/
		public static string GetBatteryStandbyValue(string PanelType)
		{
			// Verify panel type and then accordingly assign sRow value
			if(PanelType.Equals("FIM"))
			{
				sRow = (16).ToString();
				sCell= "[4]";
			}
			else
			{
				sRow = (16).ToString();
				sCell= "[5]";
			}
			
			//Click on Physical Layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			// Fetch BatteryStandby and store in Actual BatteryStandby value
			string ActualBatteryStandbyValue = repo.FormMe.BatteryStandBy.TextValue;
			
			return ActualBatteryStandbyValue;
		}
		


		/*****************************************************************************************************************
		 * Function Name: GetAlarmLoadValue
		 * Function Details: To get Alarm load value
		 * Parameter/Arguments:   expected value, panel type (FIM or PFI)
		 * Output:40V load displayed on UI
		 * Function Owner:Purvi Bhasin
		 * Last Update : 22/01/2019
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static string GetAlarmLoadValue(string PanelType)
		{
			
			// Verify panel type and then accordingly assign sRow value
			if(PanelType.Equals("FIM"))
			{
				sRow = (16).ToString();
				sCell= "[4]";
			}
			else
			{
				sRow = (16).ToString();
				sCell= "[5]";
			}
			//Click on Physical Layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			// Fetch BatteryStandby and store in Actual 40VPSU value
			string ActualAlarmLoadValue = repo.FormMe.AlarmLoad.TextValue;
			
			return ActualAlarmLoadValue;
		}
		
		/*****************************************************************************************************************
		 * Function Name: verifyBatteryStandbyAndAlarmLoadOnAdditionAndDeletionOfAccessories
		 * Function Details: To Verify 40V load on addition/deletion of R-Bus connection
		 * Parameter/Arguments:   expected Maximum value, panel type (FIM or PFI)  and row number is 13 by default for FIM
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 22/01/2019   Alpesh Dhakad - 30/07/2019 - Updated test scripts as per new build and xpaths
		 * Alpesh Dhakad - 08/08/2019 - Updated test script
		 * Alpesh Dhakad - 21/08/2019 - Updated with new navigation tree method, xpath and devices gallery
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyBatteryStandbyAndAlarmLoadOnAdditionAndDeletionOfAccessories(string sFileName,string sAddPanelSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared variables
			string PanelName,PanelNode,sBatteryStandby,sAlarmLoad,sDefaultBatteryStandby,sDefaultAlarmLoad,CPUType,sRowNumber,sType,PanelType,sExpectedBatteryStandby,sExpectedAlarmLoad,expectedDefaultBatteryStandby,expectedDefaultAlarmLoad;
			int rowNumber;
			float PrinterBatteryStandby,PrinterAlarmLoad,DefaultBatteryStandby,DefaultAlarmLoad;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sRowNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				expectedDefaultBatteryStandby = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				expectedDefaultAlarmLoad = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				
				int.TryParse(sRowNumber, out rowNumber);
				
				// Add panels using test data in excel sheet
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				Report.Log(ReportLevel.Info, "Panel "+PanelName+" added successfully");
				
				// Click on Expander node
				Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
				
				// Click on Loop Card node
				Common_Functions.ClickOnNavigationTreeExpander(PanelType);
				
				// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				// Verify Default Battery Standby load value
				verifyBatteryStandby(expectedDefaultBatteryStandby,false,PanelType);
				
				// Verify Default Alarm load value
				verifyAlarmLoad(expectedDefaultAlarmLoad,false,PanelType);
				
				// Click on Loop Card node
				Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
				
				for(int j=8; j<9; j++)
				{
					
					ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[j,9]).Value.ToString();
					sType = ((Range)Excel_Utilities.ExcelRange.Cells[j,10]).Value.ToString();
					sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[j,11]).Value.ToString();
					sBatteryStandby = ((Range)Excel_Utilities.ExcelRange.Cells[j,12]).Value.ToString();
					sAlarmLoad = ((Range)Excel_Utilities.ExcelRange.Cells[j,13]).Value.ToString();
					
					//Add Printer connection
					// Click on Loop Card node
					Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
				
					float.TryParse(sBatteryStandby, out PrinterBatteryStandby);
					float.TryParse(sAlarmLoad, out PrinterAlarmLoad);
					Devices_Functions.AddDevicesfromMainProcessorGallery(ModelNumber,sType,PanelType);
					Report.Log(ReportLevel.Info, "Device "+ModelNumber+" added successfully");
					
					// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
					//Get Battery Standby from UI
					sDefaultBatteryStandby = GetBatteryStandbyValue(PanelType);
					
					//Get Alarm Load from UI
					sDefaultAlarmLoad = GetAlarmLoadValue(PanelType);
					
					//Generate expected Battery Standby load
					float.TryParse(sDefaultBatteryStandby, out DefaultBatteryStandby);
					float ExpectedBatteryStandby = DefaultBatteryStandby+PrinterBatteryStandby;
					sExpectedBatteryStandby= ExpectedBatteryStandby.ToString("0.000");
					
					//Generate expected Alarm load
					float.TryParse(sDefaultAlarmLoad, out DefaultAlarmLoad);
					float ExpectedAlarmLoad = DefaultAlarmLoad+PrinterAlarmLoad;
					sExpectedAlarmLoad= ExpectedAlarmLoad.ToString("0.000");
					
					// Verify Battery Standby value on addition of Accessories
					verifyBatteryStandby(sExpectedBatteryStandby,false,PanelType);
					
					// Verify Alarm load value on addition of Accessories
					verifyAlarmLoad(sExpectedAlarmLoad,false,PanelType);
					
					//Get Battery Standby from UI
					sDefaultBatteryStandby = GetBatteryStandbyValue(PanelType);
					
					//Get Alarm load from UI
					sDefaultAlarmLoad = GetAlarmLoadValue(PanelType);
					
					//Generate expected Battery Standby load on Deletion
					float.TryParse(sDefaultBatteryStandby, out DefaultBatteryStandby);
					ExpectedBatteryStandby = DefaultBatteryStandby-PrinterBatteryStandby;
					sExpectedBatteryStandby = ExpectedBatteryStandby.ToString("0.000");
					
					//Generate expected Alarm load on Deletion
					float.TryParse(sDefaultAlarmLoad, out DefaultAlarmLoad);
					ExpectedAlarmLoad = DefaultAlarmLoad-PrinterAlarmLoad;
					sExpectedAlarmLoad = ExpectedAlarmLoad.ToString("0.000");
					
					// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem(PanelNode);
				
					repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1.Click();
					
					if(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info.Exists())
					{
						repo.ProfileConsys1.btn_Delete.Click();
						Validate.AttributeEqual(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info, "Text", sLabelName);
						Report.Log(ReportLevel.Success, "Device "+sLabelName+" deleted successfully");
						
						// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
						// Verify Battery Standby and Alarm load value on addition of Ethernet
						verifyBatteryStandby(sExpectedBatteryStandby,false,PanelType);
						verifyAlarmLoad(sExpectedAlarmLoad,false,PanelType);
					}
					
					else
					{
						
						Report.Log(ReportLevel.Failure, "Device "+sLabelName+" not found");
					}

					
				}
				
				
				// Delete panel using PanelNode details from excel sheet
				Panel_Functions.DeletePanel(1,PanelNode,1);
				
			}
			//Close opened excel sheet
			Excel_Utilities.CloseExcel();
			
		}


		/***********************************************************************************************************************************************
		 * Function Name: verifyBatteryStandbyAndAlarmLoadOnZetfastLoopAddDelete
		 * Function Details: To Verify 40V load on addition/deletion of Zetfast loop with devices
		 * Parameter/Arguments:   expected Maximum value, panel type (FIM or PFI)  and row number is 13 by default for FIM
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 23/01/2019 Alpesh Dhakad - 30/07/2019,21/08/2019,30/08/2019,08/09/2019- Updated test scripts as per new build and xpaths
		 ***********************************************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyBatteryStandbyAndAlarmLoadOnZetfastLoopAddDelete(string sFileName,string sAddPanelSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared variables
			string PanelName,PanelNode,sBatteryStandby,sAlarmLoad,sDefaultBatteryStandby,sDefaultAlarmLoad,CPUType,sRowNumber,sType,PanelType,sExpectedBatteryStandby,sExpectedAlarmLoad,expectedDefaultBatteryStandby,expectedDefaultAlarmLoad;
			int rowNumber;
			float ZetfastBatteryStandby,ZetfastAlarmLoad,DefaultBatteryStandby,DefaultAlarmLoad,ExpectedBatteryStandby,ExpectedAlarmLoad;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sRowNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				expectedDefaultBatteryStandby = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				expectedDefaultAlarmLoad = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				
				int.TryParse(sRowNumber, out rowNumber);
				
				// Add panels using test data in excel sheet
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				
				// Click on Expander node
					Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
					
				
				// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
				//Add zetfast loop and devices and verify Battery Standby and Alarm Load
				for(int j=6; j<=8; j++)
				{
					ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[j,8]).Value.ToString();
					sType = ((Range)Excel_Utilities.ExcelRange.Cells[j,9]).Value.ToString();
					sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[j,10]).Value.ToString();
					sBatteryStandby = ((Range)Excel_Utilities.ExcelRange.Cells[j,11]).Value.ToString();
					sAlarmLoad = ((Range)Excel_Utilities.ExcelRange.Cells[j,12]).Value.ToString();
					
					if(j==6)
					{
						Devices_Functions.AddDevicesfromMainProcessorGallery(ModelNumber,sType,PanelType);
						Report.Log(ReportLevel.Info, "Device "+ModelNumber+" added successfully");
						
						// Click on XLM Loop CardExpander node
					Common_Functions.ClickOnNavigationTreeExpander("XLM");
					
					
					}
					
					else
					{
						
						
						// Click on XLM Loop C Node to add device
						Common_Functions.ClickOnNavigationTreeItem("XLM800-Zetfas-C");
						

						Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
						
					}
					
					// Click on Panel node
					Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
					
					// Click on XLM Loop C Node to add device
					Common_Functions.ClickOnNavigationTreeItem("XLM800-Zetfas-C");
						
					
					float.TryParse(sBatteryStandby, out ZetfastBatteryStandby);
					float.TryParse(sAlarmLoad, out ZetfastAlarmLoad);
					
					//Get Battery Standby load from UI
					sDefaultBatteryStandby = GetBatteryStandbyValue(PanelType);
					
					//Get Alarm load from UI
					sDefaultAlarmLoad = GetAlarmLoadValue(PanelType);
					
					//Generate expected Battery Standby
					float.TryParse(sDefaultBatteryStandby, out DefaultBatteryStandby);
					ExpectedBatteryStandby = DefaultBatteryStandby+ZetfastBatteryStandby;
					sExpectedBatteryStandby= ExpectedBatteryStandby.ToString("0.000");
					
					//Generate expected Alarm load
					float.TryParse(sDefaultAlarmLoad, out DefaultAlarmLoad);
					ExpectedAlarmLoad = DefaultAlarmLoad+ZetfastAlarmLoad;
					sExpectedAlarmLoad= ExpectedAlarmLoad.ToString("0.000");
					
					// Verify Battery Standby value on addition of zetfast loop with devices
					verifyBatteryStandby(sExpectedBatteryStandby,false,PanelType);
					
					// Verify 40V PSU load value on addition of zetfast loop with devices
					verifyAlarmLoad(sExpectedAlarmLoad,false,PanelType);
					
					
					
					// Click on Site node
					Common_Functions.ClickOnNavigationTreeItem("Site");
			
					
				}
				
				for(int k=8; k<=6; k--)
				{
					ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[k,8]).Value.ToString();
					sType = ((Range)Excel_Utilities.ExcelRange.Cells[k,9]).Value.ToString();
					sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[k,10]).Value.ToString();
					sBatteryStandby = ((Range)Excel_Utilities.ExcelRange.Cells[k,11]).Value.ToString();
					sAlarmLoad = ((Range)Excel_Utilities.ExcelRange.Cells[k,12]).Value.ToString();
					
					//Get Battery Standby load from UI
					sDefaultBatteryStandby = GetBatteryStandbyValue(PanelType);
					
					//Get Alarm load from UI
					sDefaultAlarmLoad = GetAlarmLoadValue(PanelType);
					
					//Generate expected 40V load on Deletion
					float.TryParse(sDefaultBatteryStandby, out DefaultBatteryStandby);
					float.TryParse(sBatteryStandby, out ZetfastBatteryStandby);
					ExpectedBatteryStandby = DefaultBatteryStandby-ZetfastBatteryStandby;
					sExpectedBatteryStandby = ExpectedBatteryStandby.ToString("0.000");
					
					//Generate expected 40V load on Deletion
					float.TryParse(sDefaultAlarmLoad, out DefaultAlarmLoad);
					float.TryParse(sAlarmLoad, out ZetfastAlarmLoad);
					ExpectedAlarmLoad = DefaultAlarmLoad-ZetfastAlarmLoad;
					sExpectedAlarmLoad = ExpectedAlarmLoad.ToString("0.000");
					
					if(k==8)
					{
						// Click on Panel node
						Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
						repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1.Click();
						
						if(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info.Exists())
						{
							repo.ProfileConsys1.btn_Delete.Click();
							Validate.AttributeEqual(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info, "Text", sLabelName);
							Report.Log(ReportLevel.Success, "Device "+sLabelName+" deleted successfully");
							
							// Click on Loop A node
							Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
							// Verify Battery Standby load value on deletion of Zetfast loop
							verifyBatteryStandby(sExpectedBatteryStandby,false,PanelType);
							
							// Verify Alarm load value on deletion of Zetfast loop
							verifyAlarmLoad(sExpectedAlarmLoad,false,PanelType);
						}
					}
					
					
					else
					{
						
						
						// Click on XLM Loop C Node to add device
					Common_Functions.ClickOnNavigationTreeItem("XLM800-Zetfas-C");
					

						repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1.Click();
						
						if(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info.Exists())
						{
							repo.ProfileConsys1.btn_Delete.Click();
							Validate.AttributeEqual(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info, "Text", sLabelName);
							Report.Log(ReportLevel.Success, "Device "+sLabelName+" deleted successfully");
							
							// Click on Loop A node
							Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
							// Verify Battery Standby load value on deletion of Zetfast loop
							verifyBatteryStandby(sExpectedBatteryStandby,false,PanelType);
							
							// Verify Alarm load value on deletion of Zetfast loop
							verifyAlarmLoad(sExpectedAlarmLoad,false,PanelType);
							
						}
						else
						{
							
							Report.Log(ReportLevel.Failure, "Device "+sLabelName+" not found");
						}
						
					}
				}
				
				
				
				
				
				// Delete panel using PanelNode details from excel sheet
				Panel_Functions.DeletePanel(1,PanelNode,1);
				
			}
			//Close opened excel sheet
			Excel_Utilities.CloseExcel();
			
			
		}

		
		/***********************************************************************************************************************************************
		 * Function Name: verifyBatteryStandbyAndAlarmLoadOnSlotCardAddDelete
		 * Function Details: To Verify Battery Standby and Alarm load on addition/deletion of Slot Cards
		 * Parameter/Arguments:   expected value, panel type (FIM or PFI)  and row number is 14 by default for PFI
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 22/01/2019  Alpesh Dhakad - 30/07/2019,21/08/2019,08/09/2019- Updated test scripts as per new build and xpaths
		 ***********************************************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyBatteryStandbyAndAlarmLoadOnSlotCardAddDelete(string sFileName,string sAddPanelSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared variables
			string PanelName,PanelNode,CPUType,sBatteryStandby,sAlarmLoad,sDefaultBatteryStandby,sDefaultAlarmLoad,sRowNumber,sType,PanelType,sExpectedBatteryStandby,sExpectedAlarmLoad,expectedDefaultBatteryStandby,expectedDefaultAlarmLoad;
			int rowNumber;
			float SCBatteryStandby,SCAlarmLoad,PABatteryStandby,PAAlarmLoad,DefaultBatteryStandby,DefaultAlarmLoad;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sRowNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				expectedDefaultBatteryStandby = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				expectedDefaultAlarmLoad = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				
				int.TryParse(sRowNumber, out rowNumber);
				
				// Add panels using test data in excel sheet
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				Report.Log(ReportLevel.Info, "Panel "+PanelName+" added successfully");
				
				// Click on Expander node
				Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
				
				// Click on Loop Card node
				Common_Functions.ClickOnNavigationTreeExpander(PanelType);
				
				// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				
				// Verify Default Battery Standby load value
				verifyBatteryStandby(expectedDefaultBatteryStandby,false,PanelType);
				
				// Verify Default Alarm load value
				verifyAlarmLoad(expectedDefaultAlarmLoad,false,PanelType);
				
				// Click on Loop Card node
				Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
				
				for(int j=8; j<9; j++)
				{
					
					ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[j,9]).Value.ToString();
					sType = ((Range)Excel_Utilities.ExcelRange.Cells[j,10]).Value.ToString();
					sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[j,11]).Value.ToString();
					sBatteryStandby = ((Range)Excel_Utilities.ExcelRange.Cells[j,12]).Value.ToString();
					sAlarmLoad = ((Range)Excel_Utilities.ExcelRange.Cells[j,13]).Value.ToString();
					
					//Add Slot Card
					// Click on Loop Card node
					Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
				
					float.TryParse(sBatteryStandby, out SCBatteryStandby);
					float.TryParse(sAlarmLoad, out SCAlarmLoad);
					Devices_Functions.AddDevicesfromMainProcessorGallery(ModelNumber,sType,PanelType);
					Report.Log(ReportLevel.Info, "Device "+ModelNumber+" added successfully");
					
					// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
					//Get Battery Standby from UI
					sDefaultBatteryStandby = GetBatteryStandbyValue(PanelType);
					
					//Get Alarm Load from UI
					sDefaultAlarmLoad = GetAlarmLoadValue(PanelType);
					
					//Generate expected Battery Standby load
					float.TryParse(sDefaultBatteryStandby, out DefaultBatteryStandby);
					float ExpectedBatteryStandby = DefaultBatteryStandby+SCBatteryStandby;
					sExpectedBatteryStandby= ExpectedBatteryStandby.ToString("0.000");
					
					//Generate expected Alarm load
					float.TryParse(sDefaultAlarmLoad, out DefaultAlarmLoad);
					float ExpectedAlarmLoad = DefaultAlarmLoad+SCAlarmLoad;
					sExpectedAlarmLoad= ExpectedAlarmLoad.ToString("0.000");
					
					// Verify Battery Standby value on addition of Accessories
					verifyBatteryStandby(sExpectedBatteryStandby,false,PanelType);
					
					// Verify Alarm load value on addition of Accessories
					verifyAlarmLoad(sExpectedAlarmLoad,false,PanelType);
					
					//Get Battery Standby from UI
					sDefaultBatteryStandby = GetBatteryStandbyValue(PanelType);
					
					//Get Alarm load from UI
					sDefaultAlarmLoad = GetAlarmLoadValue(PanelType);
					
					//Generate expected Battery Standby load on Deletion
					float.TryParse(sDefaultBatteryStandby, out DefaultBatteryStandby);
					ExpectedBatteryStandby = DefaultBatteryStandby-SCBatteryStandby;
					sExpectedBatteryStandby = ExpectedBatteryStandby.ToString("0.000");
					
					//Generate expected Alarm load on Deletion
					float.TryParse(sDefaultAlarmLoad, out DefaultAlarmLoad);
					ExpectedAlarmLoad = DefaultAlarmLoad-SCAlarmLoad;
					sExpectedAlarmLoad = ExpectedAlarmLoad.ToString("0.000");
					
					// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem(PanelNode);
				
					repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1.Click();
					
					if(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info.Exists())
					{
						repo.ProfileConsys1.btn_Delete.Click();
						Validate.AttributeEqual(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info, "Text", sLabelName);
						Report.Log(ReportLevel.Success, "Device "+sLabelName+" deleted successfully");
						
						// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
						// Verify Battery Standby and Alarm load value on addition of Ethernet
						verifyBatteryStandby(sExpectedBatteryStandby,false,PanelType);
						verifyAlarmLoad(sExpectedAlarmLoad,false,PanelType);
					}
					
					else
					{
						
						Report.Log(ReportLevel.Failure, "Device "+sLabelName+" not found");
					}

					
				}
				
				//for adding panel accessories
				for(int j=8; j<=9; j++)
				{
					
					ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[j,14]).Value.ToString();
					sType = ((Range)Excel_Utilities.ExcelRange.Cells[j,15]).Value.ToString();
					sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[j,16]).Value.ToString();
					sBatteryStandby = ((Range)Excel_Utilities.ExcelRange.Cells[j,17]).Value.ToString();
					sAlarmLoad = ((Range)Excel_Utilities.ExcelRange.Cells[j,18]).Value.ToString();
					
					//Add Slot Card
					// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem(PanelNode);
				
					//click panel Accessories tab
					repo.FormMe.tab_PanelAccessories.Click();
					
					float.TryParse(sBatteryStandby, out PABatteryStandby);
					float.TryParse(sAlarmLoad, out PAAlarmLoad);
					Devices_Functions.AddDevicefromPanelAccessoriesGallery(ModelNumber,sType);
					Report.Log(ReportLevel.Info, "Device "+ModelNumber+" added successfully");
					
					// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
					//Get Battery Standby from UI
					sDefaultBatteryStandby = GetBatteryStandbyValue(PanelType);
					
					//Get Alarm Load from UI
					sDefaultAlarmLoad = GetAlarmLoadValue(PanelType);
					
					//Generate expected Battery Standby load
					float.TryParse(sDefaultBatteryStandby, out DefaultBatteryStandby);
					float ExpectedBatteryStandby = DefaultBatteryStandby+PABatteryStandby;
					sExpectedBatteryStandby= ExpectedBatteryStandby.ToString("0.000");
					
					//Generate expected Alarm load
					float.TryParse(sDefaultAlarmLoad, out DefaultAlarmLoad);
					float ExpectedAlarmLoad = DefaultAlarmLoad+PAAlarmLoad;
					sExpectedAlarmLoad= ExpectedAlarmLoad.ToString("0.000");
					
					// Verify Battery Standby value on addition of Accessories
					verifyBatteryStandby(sExpectedBatteryStandby,false,PanelType);
					
					// Verify Alarm load value on addition of Accessories
					verifyAlarmLoad(sExpectedAlarmLoad,false,PanelType);
					
					//Get Battery Standby from UI
					sDefaultBatteryStandby = GetBatteryStandbyValue(PanelType);
					
					//Get Alarm load from UI
					sDefaultAlarmLoad = GetAlarmLoadValue(PanelType);
					
					//Generate expected Battery Standby load on Deletion
					float.TryParse(sDefaultBatteryStandby, out DefaultBatteryStandby);
					ExpectedBatteryStandby = DefaultBatteryStandby-PABatteryStandby;
					sExpectedBatteryStandby = ExpectedBatteryStandby.ToString("0.000");
					
					//Generate expected Alarm load on Deletion
					float.TryParse(sDefaultAlarmLoad, out DefaultAlarmLoad);
					ExpectedAlarmLoad = DefaultAlarmLoad-PAAlarmLoad;
					sExpectedAlarmLoad = ExpectedAlarmLoad.ToString("0.000");
					
					// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem(PanelNode);
				
					//click on panel accessories tab
					repo.FormMe.tab_PanelAccessories.Click();
					
					repo.FormMe.cell_Label.Click();
					//repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1.Click();
					
					if(repo.FormMe.cell_LabelInfo.Exists())
					{
						repo.ProfileConsys1.btn_Delete.Click();
						Validate.AttributeEqual(repo.FormMe.cell_LabelInfo, "Text", sLabelName);
						Report.Log(ReportLevel.Success, "Device "+sLabelName+" deleted successfully");
						
						// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
						// Verify Battery Standby and Alarm load value on addition of Ethernet
						verifyBatteryStandby(sExpectedBatteryStandby,false,PanelType);
						verifyAlarmLoad(sExpectedAlarmLoad,false,PanelType);
					}
					
					else
					{
						
						Report.Log(ReportLevel.Failure, "Device "+sLabelName+" not found");
					}

					
				}
				
				
				// Delete panel using PanelNode details from excel sheet
				Panel_Functions.DeletePanel(1,PanelNode,1);
				
			}
			//Close opened excel sheet
			Excel_Utilities.CloseExcel();
			
		}

		/*************************************************************************************************************************
		 * Function Name: verifyMaxSystemLoadValue
		 * Function Details:
		 * Parameter/Arguments:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 4/2/2019  Alpesh Dhakad - 01/08/2019 & 23/08/2019 - Updated test scripts as per new build and xpaths
		 *************************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyMaxSystemLoadValue(string sFileName,string sAddPanelSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared string type
			string PanelName,PanelNode,CPUType,sRowNumber,PanelType,expectedMaxSystemLoad;
			int rowNumber;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sRowNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				
				expectedMaxSystemLoad = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				
				// sPSUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,10]).Value.ToString();
				
				int.TryParse(sRowNumber, out rowNumber);
				
				// Add panels using test data in excel sheet
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				Report.Log(ReportLevel.Info, "Panel "+PanelName+" added successfully");
				
				// Click on Expander node
					Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
					
					// Click on Loop Card node
					Common_Functions.ClickOnNavigationTreeExpander(PanelType);
					
					// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
				// Verify max System Load load value
				verifyMaxSystemLoad(expectedMaxSystemLoad);
				
				// Click on Site node
			Common_Functions.ClickOnNavigationTreeItem("Site");
			
				// Delete panel using PanelNode details from excel sheet
				Panel_Functions.DeletePanel(1,PanelNode,1);
				
			}
			//Close opened excel sheet
			Excel_Utilities.CloseExcel();
			
		}
		
		
		/*****************************************************************************************************************
		 * Function Name: verifyMaxSystemLoad
		 * Function Details:
		 * Parameter/Arguments:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 4/2/2019
		 *****************************************************************************************************************/
		
		public static void verifyMaxSystemLoad(string expectedMaxSystemLoad)
		{
			
			// Click on Physical layout tab
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			// Fetch System Load maximum limit value
			string maxSystemLoad = repo.FormMe.maxSystemLoad.TextValue;
			
			// Compare max40VPSU value with expected value
			if(maxSystemLoad.Equals(expectedMaxSystemLoad))
			{
				Report.Log(ReportLevel.Success,"Max System Load " + maxSystemLoad + " is displayed correctly " );
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Max System Load value is not displayed correctly, it is displayed as: " + maxSystemLoad + " instead of : " +expectedMaxSystemLoad);
			}
		}
		/*****************************************************************************************************************
		 * Function Name: verifyImpactOfSecondPSUOnBatteryAndAlarm
		 * Function Details: To Verify Impact of 2nd PSU on Battery and Alarm Load
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 22/01/2019  Alpesh Dhakad - 01/08/2019 - Updated test scripts as per new build and xpaths 
		 *  Purvi Bhasin - 07/08/2019 - Commented Node Expander so that Loop A remains visible
		 * Alpesh Dhakad - 23/08/2019 - Updated test scripts as per new build and xpaths 
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyImpactOfSecondPSUOnBatteryAndAlarm(string sFileName,string sAddPanelSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared variables
			string PanelName,PanelNode,CPUType,PSUType,ChangedPSUType,sRowNumber,sType,PanelType,sExpectedBatteryStandby,sExpectedAlarmLoad,expectedMaxBatteryStandby,expectedMaxAlarmLoad,SecondPSU,PoweredBy;
			int rowNumber;
			
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				PSUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				ChangedPSUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				sRowNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				expectedMaxBatteryStandby = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				expectedMaxAlarmLoad = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				sExpectedBatteryStandby = ((Range)Excel_Utilities.ExcelRange.Cells[i,10]).Value.ToString();
				sExpectedAlarmLoad = ((Range)Excel_Utilities.ExcelRange.Cells[i,11]).Value.ToString();
				
				int.TryParse(sRowNumber, out rowNumber);
				
				// Add panels using test data in excel sheet
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				Report.Log(ReportLevel.Info, "Panel "+PanelName+" added successfully");
				
				// Click on Expander node
					Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
					
					// Click on Loop Card node
					Common_Functions.ClickOnNavigationTreeExpander(PanelType);
					
					// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
				
				//Verify max Battery Standby and max Alarm Load
				verifyMaxBatteryStandby(expectedMaxBatteryStandby,false);
				verifyMaxAlarmLoad(expectedMaxAlarmLoad,false);
				
				// Verify Battery Standby load value
				verifyBatteryStandby(sExpectedBatteryStandby,false,PanelType);
				
				// Verify Alarm load value
				verifyAlarmLoad(sExpectedAlarmLoad,false,PanelType);
				
				
				for(int j=8; j<=9; j++)
				{
					SecondPSU = ((Range)Excel_Utilities.ExcelRange.Cells[j,12]).Value.ToString();
					expectedMaxBatteryStandby = ((Range)Excel_Utilities.ExcelRange.Cells[j,13]).Value.ToString();
					expectedMaxAlarmLoad = ((Range)Excel_Utilities.ExcelRange.Cells[j,14]).Value.ToString();
					sExpectedBatteryStandby = ((Range)Excel_Utilities.ExcelRange.Cells[j,15]).Value.ToString();
					sExpectedAlarmLoad = ((Range)Excel_Utilities.ExcelRange.Cells[j,16]).Value.ToString();
					
					//Click on points tab
					repo.ProfileConsys1.tab_Points.Click();
					
				
				// Click on Site node
			Common_Functions.ClickOnNavigationTreeItem("Site");
			
			// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
					
					Panel_Functions.ChangeSecondPSUType(SecondPSU);
					
					// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
					//Verify max Battery Standby and max Alarm Load
					verifyMaxBatteryStandby(expectedMaxBatteryStandby,true);
					verifyMaxAlarmLoad(expectedMaxAlarmLoad,true);
					
					// Verify Battery Standby load value
					verifyBatteryStandby(sExpectedBatteryStandby,true,PanelType);
					
					// Verify Alarm load value
					verifyAlarmLoad(sExpectedAlarmLoad,true,PanelType);
					
					// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
					
					
					
					for(int k=8; j<9; j++)
					{
						
						ModelNumber =  ((Range)Excel_Utilities.ExcelRange.Cells[k,18]).Value.ToString();
						sType = ((Range)Excel_Utilities.ExcelRange.Cells[k,19]).Value.ToString();
						sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[k,20]).Value.ToString();
						PoweredBy = ((Range)Excel_Utilities.ExcelRange.Cells[k,21]).Value.ToString();
						sExpectedBatteryStandby = ((Range)Excel_Utilities.ExcelRange.Cells[k,22]).Value.ToString();
						sExpectedAlarmLoad = ((Range)Excel_Utilities.ExcelRange.Cells[k,23]).Value.ToString();
						
						
						// Add devices from Panel node gallery
						Devices_Functions.AddDevicesfromMainProcessorGallery(ModelNumber,sType,PanelType);
						
						// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
						
						// Verify Battery Standby load value
						verifyBatteryStandby(sExpectedBatteryStandby,true,PanelType);
						
						// Verify Alarm load value
						verifyAlarmLoad(sExpectedAlarmLoad,true,PanelType);
						
						// Click on Site node
			Common_Functions.ClickOnNavigationTreeItem("Site");
			
						
						//Change Powered From
						
				// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
						repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1.Click();
						Panel_Functions.DevicePoweredFrom(PoweredBy);
						
						sExpectedBatteryStandby = ((Range)Excel_Utilities.ExcelRange.Cells[k,24]).Value.ToString();
						sExpectedAlarmLoad= ((Range)Excel_Utilities.ExcelRange.Cells[k,25]).Value.ToString();
						
						
						// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
						// Verify Battery Standby load value
						verifyBatteryStandby(sExpectedBatteryStandby,true,PanelType);
						
						// Verify Alarm load value
						verifyAlarmLoad(sExpectedAlarmLoad,true,PanelType);
						
						
						// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
						repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1.Click();
						
						if(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info.Exists())
						{
							repo.ProfileConsys1.btn_Delete.Click();
							Validate.AttributeEqual(repo.ProfileConsys1.PanelInvetoryGrid.txt_Label1Info, "Text", sLabelName);
							Report.Log(ReportLevel.Success, "Device "+sLabelName+" deleted successfully");
							
							// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
						}
						
						else
						{
							
							Report.Log(ReportLevel.Failure, "Device "+sLabelName+" not found");
						}

						
					}
					
					
					// Delete panel using PanelNode details from excel sheet
					Panel_Functions.DeletePanel(1,PanelNode,1);
					
				}
				//Close opened excel sheet
				Excel_Utilities.CloseExcel();
				
			}
			
			
		}
		
		/***************************************************************************
		 * Function Details: To Verify UI on adding Second PSU
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 22/01/2019
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyUIOnAddingSecondPSU(bool isSecondPSU)
		{
			if(isSecondPSU)
			{
				if(repo.FormMe.Cell_AdditionalPSUInfo.Exists())
				{
					Report.Log(ReportLevel.Success,"Additional PSU is present" );
				}
				else
				{
					Report.Log(ReportLevel.Failure,"Additional PSU properties are not displayed" );
				}
			}
			else
			{
				if(repo.FormMe.Cell_AdditionalPSUInfo.Exists())
				{
					Report.Log(ReportLevel.Failure,"Additional PSU is present" );
				}
				else
				{
					Report.Log(ReportLevel.Success,"Additional PSU properties are not displayed" );
				}
			}
		}
		
		
		
		/*****************************************************************************************************************
		 * Function Name: verifyPowerCalculations
		 * Function Details: To verify Power Calculations error and warning along with Indicators
		 * Parameter/Arguments: filename and sheetname
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 25/03/2019
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyPowerCalculationsFor24VAndSystemLoad(string sFileName,string sAddPanelandDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelandDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared variables type
			string PanelType,sExpectedPowerCalculationText,sDeviceName,sLabelName;
			int DeviceQty;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				sDeviceName=  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				DeviceQty=  int.Parse(((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString());
				
				sExpectedPowerCalculationText= ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				
				PanelType= ((Range)Excel_Utilities.ExcelRange.Cells[5,5]).Value.ToString();
				
				Devices_Functions.AddDevicesfromMultiplePointWizard(sDeviceName,DeviceQty);
				
				
				verifyPowerCalculationsFor24V(PanelType);
				
				verifyPowerCalculationsText(sExpectedPowerCalculationText);

			}
			// Close Excel sheet
			Excel_Utilities.CloseExcel();
		}
		
		
		/*****************************************************************************************************************
		 * Function Name:verifyPowerCalculationsFor40VAndDCUnits
		 * Function Details: To verify Power Calculations For40V And DCUnits
		 * Parameter/Arguments: filename and sheetname
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 25/03/2019
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyPowerCalculationsFor40VAndDCUnits(string sFileName,string sAddPanelandDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelandDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared variables type
			string PanelType,sExpectedPowerCalculationText,sDeviceName,sLabelName;
			int DeviceQty;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				sDeviceName=  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				DeviceQty=  int.Parse(((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString());
				
				sExpectedPowerCalculationText= ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				
				PanelType= ((Range)Excel_Utilities.ExcelRange.Cells[5,5]).Value.ToString();
				
				Devices_Functions.AddDevicesfromMultiplePointWizard(sDeviceName,DeviceQty);
				
				verifyPowerCalculationsFor40V(PanelType);
				
				// Verify
				verifyPowerCalculationsForDCUnits(PanelType);
				
				verifyPowerCalculationsText(sExpectedPowerCalculationText);
				
			}
			// Close Excel sheet
			Excel_Utilities.CloseExcel();
		}
		
		/*****************************************************************************************************************
		 * Function Name:verifyPowerCalculationsFor40VACAndDCUnits
		 * Function Details: To verify Power Calculations For40V, AC And DCUnits
		 * Parameter/Arguments: filename and sheetname
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 25/03/2019
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyPowerCalculationsFor40VACAndDCUnits(string sFileName,string sAddPanelandDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelandDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared variables type
			string PanelType,sExpectedPowerCalculationText,sDeviceName,sLabelName;
			int DeviceQty;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				sDeviceName=  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				DeviceQty=  int.Parse(((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString());
				
				sExpectedPowerCalculationText= ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				
				PanelType= ((Range)Excel_Utilities.ExcelRange.Cells[5,5]).Value.ToString();
				
				Devices_Functions.AddDevicesfromMultiplePointWizard(sDeviceName,DeviceQty);
				
				
				verifyPowerCalculationsForACUnits(PanelType);
				
				verifyPowerCalculationsText(sExpectedPowerCalculationText);
				
			}
			// Close Excel sheet
			Excel_Utilities.CloseExcel();
		}
		
		
		/*****************************************************************************************************************
		 * Function Name:
		 * Function Details: To verify Power Calculations For IS Units
		 * Parameter/Arguments: filename and sheetname
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 26/03/2019
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyPowerCalculationsForISUnitsAndACUnits(string sFileName,string sAddPanelandDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelandDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared variables type
			string PanelType,sExpectedPowerCalculationText,sDeviceName,sLabelName,sType,sRowNumber;
			
			PanelType= ((Range)Excel_Utilities.ExcelRange.Cells[5,5]).Value.ToString();
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				

				sRowNumber= (1).ToString();
				
				
				sDeviceName=  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sType=  (((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString());

				Devices_Functions.AddDevicesfromGallery(sDeviceName,sType);
				
				if(i>8){
					// Click on first added EXI800
					Devices_Functions.SelectPointsGridRow(sRowNumber);
				}
				
				
			}
			
			sExpectedPowerCalculationText= ((Range)Excel_Utilities.ExcelRange.Cells[2,7]).Value.ToString();
			verifyPowerCalculationsForISUnits(PanelType);
			
			verifyPowerCalculationsText(sExpectedPowerCalculationText);
			
			// Close Excel sheet
			Excel_Utilities.CloseExcel();
			
		}
		
		
		/*****************************************************************************************************************
		 * Function Name:verifyPowerCalculationsForExtraISUnits
		 * Function Details: To verify Power Calculations For Extra IS Units
		 * Parameter/Arguments: filename and sheetname
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 26/03/2019
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyPowerCalculationsForExtraISUnits(string sFileName,string sAddPanelandDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelandDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared variables type
			string PanelType,sExpectedPowerCalculationText,sDeviceName,sLabelName,sType;
			
			PanelType= ((Range)Excel_Utilities.ExcelRange.Cells[5,5]).Value.ToString();
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				
				sDeviceName=  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				sLabelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sType=  (((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString());
				string ChangedValue =  ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				
				Devices_Functions.AddDevicesfromGallery(sDeviceName,sType);
				
				
				repo.ProfileConsys1.cell_CableCapacitance.DoubleClick();
				repo.ProfileConsys1.txt_CableCapacitance.PressKeys((ChangedValue) +"{ENTER}");
			}
			
			sExpectedPowerCalculationText =  ((Range)Excel_Utilities.ExcelRange.Cells[2,7]).Value.ToString();
			sPhysicalLayoutDeviceIndex =(1).ToString();
			
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			repo.ProfileConsys1.tab_Points.Click();
			
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			repo.FormMe.PhysicalLayoutDeviceIndex.Click();
			
			
			// Click on Power Calculations tab
			repo.FormMe.tab_PowerCalculations.Click();
			
			// Retrieve PowerCalculation Text value
			string actualPowerCalculationText = repo.FormMe.PowerCalculationText_Single.TextValue;
			
			
			// Compare actual and expected power calculation text value
			if (actualPowerCalculationText.Equals(sExpectedPowerCalculationText))
			{
				Report.Log(ReportLevel.Success,"Power calculation text message " +actualPowerCalculationText+ " is correctly displayed" );
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Power calculation text message " +actualPowerCalculationText+ " is not displayed correctly" );
			}
			
			
			
			Excel_Utilities.CloseExcel();
		}
		
		/*****************************************************************************************************************
		 * Function Name: verifyPowerCalculationsForISUnits
		 * Function Details: To verify PowerCalculations for ISUnits
		 * Parameter/Arguments: Panel type
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 26/03/2019
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyPowerCalculationsForISUnits(string PanelType)
		{
			string actualColour,expectedColor;
			
			
			sPhysicalLayoutDeviceIndex =(1).ToString();
			
			//Go to Physical layout
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			
			//Go to Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			//Go to Physical layout
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			
			// Click to select First EXI800 device
			if(repo.FormMe.PhysicalLayoutDeviceIndexInfo.Exists())
			{
				repo.FormMe.PhysicalLayoutDeviceIndex.Click();
			}
			else
			{
				repo.FormMe.PhysicalLayoutIndex_ISUnits.Click();
			}
			
			if(PanelType.Equals("FIM"))
			{
				sRow = (10).ToString();
				sRowIndex = (10).ToString();
			}
			else
			{
				sRow = (11).ToString();
				sRowIndex = (11).ToString();
			}

			float ActualISUnits = float.Parse(repo.FormMe.ISUnits.TextValue);
			
			//Retrieve foreground color
			actualColour = repo.FormMe.ISUnitProgressBar.GetAttributeValue<string>("foreground");
			
			//Fetch max volt drop text value and storing it in string
			float maxISUnitsValue = float.Parse(repo.FormMe.MaxISUnits.TextValue);
			
			// To calculate and get the expected color value
			expectedColor = Devices_Functions.calculatePercentage(ActualISUnits, maxISUnitsValue);
			
			// To verify Percentage
			Devices_Functions.VerifyPercentage(expectedColor, actualColour);
			
			
			//Go to Points tab
			repo.ProfileConsys1.tab_Points.Click();
		}
		
		
		
		/*****************************************************************************************************************
		 * Function Name: verifyPowerCalculationsForACUnits
		 * Function Details: To verify PowerCalculations for ACUnits
		 * Parameter/Arguments: Panel type
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 25/03/2019
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyPowerCalculationsForACUnits(string PanelType)
		{
			
			string actualColour,expectedColor;
			sRow = (1).ToString();
			
			//Go to Physical layout
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			
			float ActualACUnits = float.Parse(repo.ProfileConsys1.ACUnits.TextValue);
			
			//Retrieve foreground color
			actualColour = repo.FormMe.LoadingDetailsProgressbar.GetAttributeValue<string>("foreground");
			
			//Fetch max volt drop text value and storing it in string
			float maxACUnitsValue = float.Parse(repo.ProfileConsys1.MaxACUnits.TextValue);
			
			// To calculate and get the expected color value
			expectedColor = Devices_Functions.calculatePercentage(ActualACUnits, maxACUnitsValue);
			
			// To verify Percentage
			Devices_Functions.VerifyPercentage(expectedColor, actualColour);
			
		}
		
		/*****************************************************************************************************************
		 * Function Name: verifyPowerCalculationsForDCUnits
		 * Function Details: To verify PowerCalculations for DCUnits
		 * Parameter/Arguments: Panel type
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 25/03/2019
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyPowerCalculationsForDCUnits(string PanelType)
		{
			
			string actualColour,expectedColor;
			sRow = (2).ToString();
			
			//Go to Physical layout
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			
			float ActualDCUnits = float.Parse(repo.ProfileConsys1.DCUnits.TextValue);
			
			
			//Retrieve foreground color
			actualColour = repo.FormMe.LoadingDetailsProgressbar.GetAttributeValue<string>("foreground");
			
			//Fetch max volt drop text value and storing it in string
			float maxDCUnitsValue = float.Parse(repo.ProfileConsys1.MaxDCUnits.TextValue);
			
			// To calculate and get the expected color value
			expectedColor = Devices_Functions.calculatePercentage(ActualDCUnits, maxDCUnitsValue);
			
			// To verify Percentage
			Devices_Functions.VerifyPercentage(expectedColor, actualColour);
			
		}
		
		
		/*****************************************************************************************************************
		 * Function Name: verifyPowerCalculationsFor24V
		 * Function Details: To verify PowerCalculations for 24V
		 * Parameter/Arguments: Panel type
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 25/03/2019
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyPowerCalculationsFor24V(string PanelType)
		{
			string actualColour,expectedColor;
			
			//Go to Physical layout
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			if(PanelType.Equals("FIM"))
			{
				sRow = (13).ToString();
				sCell= "[1]";
			}
			else
			{
				sRow = (14).ToString();
				sCell= "[1]";
			}
			
			float Actual24VPSUValue = float.Parse(repo.FormMe.Psu24VLoad.TextValue);
			
			//Retrieve foreground color
			actualColour = repo.FormMe.LoadingDetailsProgressbar.GetAttributeValue<string>("foreground");
			
			//Fetch max volt drop text value and storing it in string
			float max24VPSUValue = float.Parse(repo.ProfileConsys1.Max24VPsu.TextValue);
			
			// To calculate and get the expected color value
			expectedColor = Devices_Functions.calculatePercentage(Actual24VPSUValue, max24VPSUValue);
			
			// To verify Percentage
			Devices_Functions.VerifyPercentage(expectedColor, actualColour);
			
		}
		
		
		/*****************************************************************************************************************
		 * Function Name:
		 * Function Details:
		 * Parameter/Arguments: filename and sheetname
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 22/02/2019
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyPowerCalculationsFor40V(string PanelType)
		{
			
			string actualColour,expectedColor;
			
			//Go to Physical layout
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			
			if(PanelType.Equals("FIM"))
			{
				sRow = (14).ToString();
				sCell= "[2]";
			}
			else
			{
				sRow = (6).ToString();
				sCell= "[5]";
			}
			
			
			float Actual40VPSUValue = float.Parse(repo.FormMe.Psu40VLoad.TextValue);
			
			//Retrieve foreground color
			actualColour = repo.FormMe.LoadingDetailsProgressbar.GetAttributeValue<string>("foreground");
			
			//Fetch max volt drop text value and storing it in string
			float max40VPSUValue = float.Parse(repo.ProfileConsys1.Max40VPsu.TextValue);
			
			// To calculate and get the expected color value
			expectedColor = Devices_Functions.calculatePercentage(Actual40VPSUValue, max40VPSUValue);
			
			// To verify Percentage
			Devices_Functions.VerifyPercentage(expectedColor, actualColour);
			
		}
		
		/*****************************************************************************************************************
		 * Function Name:
		 * Function Details:
		 * Parameter/Arguments: filename and sheetname
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : Alpesh Dhakad - 30/07/2019 - Updated test scripts as per new build and xpaths
		 * Alpesh Dhakad - 09/08/2019 & 21/08/2019 - Updated code to fetch text for single row and added xpath also
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyPowerCalculationsText(string sExpectedPowerCalculationText)
		{

			// Click on Power Calculations tab
			repo.FormMe.tab_PowerCalculations.Click();
			
			// To calculated children counts present under Power calculation tab
			int PowerCalculationChildrenCount = repo.FormMe.PowerCalculationContainer.Children.Count;

			// Comapre Count value and then peforming action to verify the text
			if(PowerCalculationChildrenCount.Equals(0))
			{
				Report.Log(ReportLevel.Info, "Power Calculation container doesn't contains any warning or errors ");
			}
			else
			{
				
				// Split the text from excel sheet
				string[] splitPowerCalculationText  = sExpectedPowerCalculationText.Split(',');
				int splitPowerCalculationTextCount  = sExpectedPowerCalculationText.Split(',').Length;
				
				// Verify warning error text from power calculation tab
				for(int k=0; k<=(splitPowerCalculationTextCount-1); k++)
				{
					sExpectedPowerCalculationText = splitPowerCalculationText[k];

					// Set sRow value which is used in PowerCalculationText
					sRow = (k+1).ToString();
					
					if(splitPowerCalculationTextCount==1)
					{
						string actualPowerCalculationText = repo.FormMe.PowerCalculationText_Single.TextValue;
						
						// Compare actual and expected power calculation text value
					if (actualPowerCalculationText.Equals(sExpectedPowerCalculationText))
					{
						Report.Log(ReportLevel.Success,"Power calculation text message " +actualPowerCalculationText+ " is correctly displayed" );
					}
					else
					{
						Report.Log(ReportLevel.Failure,"Power calculation text message " +actualPowerCalculationText+ " is not displayed correctly" );
					}
					
					}
					else
					{
					// Retrieve PowerCalculation Text value
					string actualPowerCalculationText = repo.FormMe.PowerCalculationText.TextValue;

					// Compare actual and expected power calculation text value
					if (actualPowerCalculationText.Equals(sExpectedPowerCalculationText))
					{
						Report.Log(ReportLevel.Success,"Power calculation text message " +actualPowerCalculationText+ " is correctly displayed" );
					}
					else
					{
						Report.Log(ReportLevel.Failure,"Power calculation text message " +actualPowerCalculationText+ " is not displayed correctly" );
					}
					}
				}
				
				
			}
			//Go to Points tab
			repo.ProfileConsys1.tab_Points.Click();
			
			// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
			
		}
		
		
		/*******************************************************************************************************************************
		 * Function Name:VerifyNormalLoadandAlarmLoadPropertyOnChangingPowerSource
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner:Purvi Bhasin
		 * Last Update :4/2/2019   Alpesh Dhakad - 01/08/2019 & 23/08/2019- Updated test scripts as per new build and xpaths
		 *******************************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyNormalLoadandAlarmLoadPropertyOnChangingPowerSource(string sFileName,string sAddPanelSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared string type
			string PanelName,PanelNode,CPUType,sRowNumber,PanelType,ModelNumber,sType,sLabel,sPowerSupply,expectedDefaultBatteryStandby,expectedDefaultAlarmLoad,sChangePowerSupply,expectedBatteryStandby,expectedAlarmLoad;
			int rowNumber;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sRowNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				ModelNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				sLabel = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				sPowerSupply = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				expectedDefaultBatteryStandby = ((Range)Excel_Utilities.ExcelRange.Cells[i,10]).Value.ToString();
				expectedDefaultAlarmLoad = ((Range)Excel_Utilities.ExcelRange.Cells[i,11]).Value.ToString();

				
				int.TryParse(sRowNumber, out rowNumber);
				
				// Add panels using test data in excel sheet
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				Report.Log(ReportLevel.Info, "Panel "+PanelName+" added successfully");

				// Click on Expander node
					Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
					
					// Click on Loop Card node
					Common_Functions.ClickOnNavigationTreeExpander(PanelType);
					
					// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
				
				// Verify Default Battery Standby load value
				verifyBatteryStandby(expectedDefaultBatteryStandby,false,PanelType);
				
				// Verify Default Alarm load value
				verifyAlarmLoad(expectedDefaultAlarmLoad,false,PanelType);
				
				// Click on Site node
			Common_Functions.ClickOnNavigationTreeItem("Site");
			
				
				// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
				//Click on Main Processor
				Common_Functions.ClickOnNavigationTreeItem("Main");
			
				//Add Device from gallery
				Devices_Functions.AddDevicesfromMainProcessorGallery(ModelNumber, sType,PanelType);
				
				// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
				//Click on Device
				Devices_Functions.SelectRowUsingLabelName(sLabel);
				
				Devices_Functions.VerifyPowerSupply(sPowerSupply);
				
				// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				// Verify Default Battery Standby load value
				verifyBatteryStandby(expectedDefaultBatteryStandby,false,PanelType);
				
				// Verify Default Alarm load value
				verifyAlarmLoad(expectedDefaultAlarmLoad,false,PanelType);
				
				
				
				for(int j=5; j<=7; j++)
				{
					sChangePowerSupply = ((Range)Excel_Utilities.ExcelRange.Cells[j,13]).Value.ToString();
					expectedBatteryStandby = ((Range)Excel_Utilities.ExcelRange.Cells[j,14]).Value.ToString();
					expectedAlarmLoad = ((Range)Excel_Utilities.ExcelRange.Cells[j,15]).Value.ToString();
					
					// Click on Site node
			Common_Functions.ClickOnNavigationTreeItem("Site");
			
				
				// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
					//Click on Device
					Devices_Functions.SelectRowUsingLabelName(sLabel);
					
					//Change Power Supply
					Devices_Functions.ChangePowerSupply(sChangePowerSupply);
					
					// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
					
					// Verify Default Battery Standby load value
					verifyBatteryStandby(expectedBatteryStandby,false,PanelType);
					
					// Verify Default Alarm load value
					verifyAlarmLoad(expectedAlarmLoad,false,PanelType);
				}
				
				//Close opened excel sheet
				Excel_Utilities.CloseExcel();
				
			}
			
		}
		/*****************************************************************************************************************
		 * Function Name:VerifyNormalLoadandAlarmLoadPropertyOnAdditionDeletionOfDevicesInPLXOrXLMLoop
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner:Purvi Bhasin
		 * Last Update :4/2/2019  Alpesh Dhakad - 30/07/2019 - Updated test scripts as per new build and xpaths
		 * Alpesh Dhakad - 08/08/2019 - Updated script
		 * Alpesh Dhakad - 21/08/2019 - Updated with new navigation tree method, xpath and devices gallery
		 * Alpesh Dhakad - 08/09/2019 - Updated scripts - removed last line after delete panel
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyNormalLoadandAlarmLoadPropertyOnAdditionDeletionOfDevicesInPLXOrXLMLoop(string sFileName,string sAddPanelSheet, string sAddDeviceSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;

			// Declared string type
			string PanelName, PanelNode,RowNumber,RowNumberForAlarm,CPUType,PanelType,BatterStandby,AlarmLoad,ChangePanelLED,LEDBatterStandby,LEDAlarmLoad,ModelNumber,sType;
			int PanelLED;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				BatterStandby = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				AlarmLoad = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				RowNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				RowNumberForAlarm = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				ChangePanelLED = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				LEDBatterStandby = ((Range)Excel_Utilities.ExcelRange.Cells[i,9]).Value.ToString();
				LEDAlarmLoad = ((Range)Excel_Utilities.ExcelRange.Cells[i,10]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,11]).Value.ToString();
				
				int.TryParse(ChangePanelLED, out PanelLED);
				
				// Add panels using test data in excel sheet
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				Report.Log(ReportLevel.Info, "Panel "+PanelName+" added successfully");
				
				
				// Click on Expander node
				Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
				
				// Click on Loop Card node
				Common_Functions.ClickOnNavigationTreeExpander(PanelType);
				
				// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				
				Excel_Utilities.CloseExcel();
				
				Excel_Utilities.OpenExcelFile(sFileName,sAddDeviceSheet);
				
				// Count number of rows in excel and store it in rows variable
				int rows2= Excel_Utilities.ExcelRange.Rows.Count;
				
				for(int j=2; j<=rows2; j++)
				{
					ModelNumber = ((Range)Excel_Utilities.ExcelRange.Cells[j,1]).Value.ToString();
					Report.Log(ReportLevel.Info, "Device "+ModelNumber+" added successfully");
					sType = ((Range)Excel_Utilities.ExcelRange.Cells[j,2]).Value.ToString();
					
					Devices_Functions.AddDevicesfromGallery(ModelNumber,sType);
					
				}
				
				Excel_Utilities.CloseExcel();
				
				Excel_Utilities.OpenExcelFile(sFileName,sAddPanelSheet);
				
				//Verify Battery Standby
				verifyBatteryStandbyAccToRow(BatterStandby,RowNumber,PanelType);
				
				//Verify Alarm Load
				verifyAlarmLoadAccToRow(AlarmLoad,RowNumberForAlarm,PanelType);
				
				// Click on Site node
				Common_Functions.ClickOnNavigationTreeItem("Site");
				
				// Click on Panel node
				Common_Functions.ClickOnNavigationTreeItem(PanelNode);
				
				
				//Change Panel LED
				Panel_Functions.changePanelLED(PanelLED);
				
				// Click on Loop A node
				Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
				
				// Verify Default Battery Standby load value
				verifyBatteryStandbyAccToRow(LEDBatterStandby,RowNumber,PanelType);
				
				// Verify Default Alarm load value
				verifyAlarmLoadAccToRow(LEDAlarmLoad,RowNumberForAlarm,PanelType);
				
				// Click on Site node
				Common_Functions.ClickOnNavigationTreeItem("Site");
				
				// Delete added Panel
				Panel_Functions.DeletePanel(1,PanelNode,1);
				
			}
			Excel_Utilities.CloseExcel();
		}
		
		/*****************************************************************************************************************
		 * Function Name: verifySystemLoadValue
		 * Function Details:
		 * Parameter/Arguments:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 4/2/2019
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void verifySystemLoadValue(string sSystemLoadValue)
		{
			sPsuV = sSystemLoadValue;
			repo.FormMe.SystemLoad.Click();
			string sActualLoadValue = repo.FormMe.SystemLoad.TextValue;
			
			Report.Log(ReportLevel.Info,"System Load value is"+sActualLoadValue);
			
			if(sSystemLoadValue.Equals(sActualLoadValue))
			{
				Report.Log(ReportLevel.Success,"System Load value is displayed "+sActualLoadValue+"correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"System Load value is displayed "+sActualLoadValue+"instad of"+sSystemLoadValue);
			}
		}
		
		/********************************************************************************************************************************************
		 * Function Name: verifySystemLoadValueOnChangingPSU
		 * Function Details:
		 * Parameter/Arguments:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 4/2/2019 Alpesh Dhakad - 23/08/2019 - Updated with new navigation tree method, xpath and devices gallery
		 ********************************************************************************************************************************************/
		[UserCodeMethod]
		public static void verifySystemLoadValueOnChangingPSU(string sFileName,string sAddPanelSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared string type
			string PanelName,PanelNode,CPUType,sRowNumber,PanelType,PSUType,expectedSystemLoad,DefaultSystemLoad;
			int rowNumber;
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sRowNumber = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				PSUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				expectedSystemLoad = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				DefaultSystemLoad = ((Range)Excel_Utilities.ExcelRange.Cells[i,8]).Value.ToString();
				
				int.TryParse(sRowNumber, out rowNumber);
				
				// Add panels using test data in excel sheet
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				
				// Click on Expander node
					Common_Functions.ClickOnNavigationTreeExpander(PanelNode);
					
					// Click on Loop Card node
					Common_Functions.ClickOnNavigationTreeExpander(PanelType);
					
					// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
				//Click on Physical Layout Tab
				repo.ProfileConsys1.tab_PhysicalLayout.Click();
				
				// Verify max System Load load value
				verifySystemLoadValue(DefaultSystemLoad);
				
				// Click on Site node
			Common_Functions.ClickOnNavigationTreeItem("Site");
			
				
				// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem(PanelNode);
			
				
				//Change PSU
				Panel_Functions.ChangePSUType(PSUType);
				
				// Click on Loop A node
					Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
					
				//Click on Physical Layout Tab
				repo.ProfileConsys1.tab_PhysicalLayout.Click();
				
				// Verify max System Load load value
				verifySystemLoadValue(expectedSystemLoad);
				
				// Click on Site node
			Common_Functions.ClickOnNavigationTreeItem("Site");
			
				// Delete panel using PanelNode details from excel sheet
				Panel_Functions.DeletePanel(1,PanelNode,1);
				
			}
			//Close opened excel sheet
			Excel_Utilities.CloseExcel();
			
		}

	}
	


}

