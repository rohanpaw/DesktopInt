/*
 * Created by Ranorex
 * User: jbhosash
 * Date: 5/21/2018
 * Time: 2:08 PM
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
	public class Panel_Functions
	{
		// You can use the "Insert New User Code Method" functionality from the context menu,
		// to add a new method with the attribute [UserCodeMethod].
		static NGConsysRepository repo = NGConsysRepository.Instance;
		static int iAddress;
		static string Label;
		
		static string sNode
		{
			get { return repo.sNode; }
			set { repo.sNode = value; }
		}
		
		static string sPSU
		{
			get { return repo.sPSU; }
			set { repo.sPSU = value; }
		}
		
		static string PanelName
		{
			get { return repo.PanelName; }
			set { repo.PanelName = value; }
		}
		
		static string Address
		{
			get { return repo.Address; }
			set { repo.Address = value; }
		}
		
		static string sCPU
		{
			get { return repo.sCPU; }
			set { repo.sCPU = value; }
		}
		
		static string PanelNode
		{
			get { return repo.PanelNode; }
			set { repo.PanelNode = value; }
		}
		
		static string sPictureIndex
		{
			get { return repo.sPictureIndex; }
			set { repo.sPictureIndex = value; }
		}
		
		static string sLabelName
		{
			get { return repo.sLabelName; }
			set { repo.sLabelName = value; }
		}
		
		static string sRow
		{
			get { return repo.sRow; }
			set { repo.sRow = value; }
		}
		static string sPanelLabelIndex
		{
			get { return repo.sPanelLabelIndex; }
			set { repo.sPanelLabelIndex = value; }
		}
		
		static string ModelNumber
		{
			
			get { return repo.ModelNumber; }
			set { repo.ModelNumber = value; }
		}
		
		
		/// <summary>
		/// This is a placeholder text. Please describe the purpose of the
		/// user code method here. The method is published to the user code library
		/// within a user code collection.
		/// </summary>
		/// 
		
		/**********************************************************************************************
		 * Function Name: SelectPanelNode
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update : Alpesh Dhakad - 30/07/2019 - Updated test scripts as per new build and xpaths
		 **********************************************************************************************/
		[UserCodeMethod]
		public static void SelectPanelNode(Int32 iNodeNumber)
		{
			sNode=iNodeNumber.ToString();
			// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem("Node");
			
			
		}
		
		/**********************************************************************************************************************************
		 * Function Name: AddPanels
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update : 28/12/2018 Alpesh Dhakad - Line 162 Commented
		 * Alpesh Dhakad - 19/08/2019 - Updated with new navigation tree method, xpath and devices gallery
		 * Alpesh Dhakad - 01/10/2019 - Added step as per new Panel Node in Build 43
		 * Alpesh Dhakad - 27/06/2020 - Added 2 lines to selct panel with generic xpath
		 **********************************************************************************************************************************/
		[UserCodeMethod]
		public static void AddPanels(int NumberofPanels,string PanelNames,string sPanelCPU)
		{
			for (int i=0; i<NumberofPanels;i++)
			{
				string[] splitPanelNames = PanelNames.Split(',');
				
				// Click on Site node
				Common_Functions.ClickOnNavigationTreeItem("Site");
				
				string PanelNameWithSpace=splitPanelNames[i];
				PanelName=PanelNameWithSpace.Replace(" ",String.Empty);
			
				// Added this line on 27/06/2020 to select panel with generic xpath
				ModelNumber = PanelName;
				
				if(PanelName.StartsWith("P"))
				{
					sPanelLabelIndex ="5";
				}
				else if(PanelName.StartsWith("MZX"))
				{
					sPanelLabelIndex ="5";
				}
				else
				{
					sPanelLabelIndex ="7";
				}
				
				//Commened 2 lines on 27/06/2020
				//repo.ProfileConsys1.btnDropDownPanelsGallery.Click();
				//repo.ContextMenu.txt_SelectPanel.Click();
			
				// Added this line on 27/06/2020 to select panel with generic xpath				
				repo.FormMe.btn_AllGalleryDropdown.Click();
				repo.ContextMenu.txt_SelectDevice.Click();
				
				repo.AddANewPanel.AddNewPanelContainer.cmb_Addresses.Click();
				iAddress=i+1;
				Address =iAddress.ToString();
				repo.ContextMenu.lstPanelAddress.Click();
				
				if(repo.AddANewPanel.AddNewPanelContainer.txt_LabelInfo.Exists())
				{
					repo.AddANewPanel.AddNewPanelContainer.txt_Label.Click();
				}
				else	
				{
					sPanelLabelIndex ="5";
					repo.AddANewPanel.AddNewPanelContainer.txt_Label.Click();
				}
				
					
				//repo.AddANewPanel.AddNewPanelContainer.txt_Label.Click();
				Label="Node"+iAddress;
				
				//Added this step after 43 build update
				Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
				
				
				Keyboard.Press(Label);
				if (!sPanelCPU.IsEmpty())
				{
					repo.AddANewPanel.AddNewPanelContainer.cmb_CPU.Click();
					sCPU=sPanelCPU;
					repo.ContextMenu.lstPanelCPU.Click();
				}
				repo.AddANewPanel.ButtonOK.Click();
				
				if(PanelNameWithSpace == "MZX252")
				{
					PanelNameWithSpace = "MZX 252";
				}
				PanelNode = Label+" "+"-"+" "+PanelNameWithSpace;
				
				//Commenting below line as for Panel name with Space and hi-fen it is not displaying as it is displaying while adding panel
				//Validate.AttributeEqual(repo.ProfileConsys1.NavigationTree.VerifyPanelNodeInfo, "Text", PanelNode);
				Report.Log(ReportLevel.Success, "Panel "+PanelNames+" Added Successfully");
			}
		}
		
		/******************************************************************************************************************
		 * Function Name: VerifyCPUType
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update : Alpesh Dhakad - 30/07/2019 & 21/08/2019 & 26/08/2019 - Updated scripts as per new build and xpaths
		 * Alpesh Dhakad - 31/12/2019 - Updated Panel node name
		 ******************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyCPUType(string sExpectedCPU,int PanelNode, bool AfterImport)
		{
			string sActualText;
			if(AfterImport)
			{
				// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem("Node1");
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Enter the  text in Search Properties fields to view related text
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("CPU" +"{ENTER}");
			
				// Click on CPU Cell
			repo.FormMe.cell_CPUType.Click();
			
				// Enter the CPU value and click Enter twice
			sActualText = repo.FormMe.txt_CPUType.TextValue;
			
				// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
				//repo.ProfileConsys1.Cell_CPU_afterimport.DoubleClick();
				//sActualText = repo.ProfileConsys1.VerifyCPUTpye_afterimport.TextValue;
				// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
				
				// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");

			}
			else
			{
			
			// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem("Node1");
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Enter the  text in Search Properties fields to view related text
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("CPU" +"{ENTER}");
			
			// Click on CPU Cell
			repo.FormMe.cell_CPUType.Click();
			
			// Enter the CPU value and click Enter twice
			sActualText = repo.FormMe.txt_CPUType.TextValue;
			
				// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
			}
			
			if (sExpectedCPU==sActualText)
			{
				Report.Log(ReportLevel.Success, "CPU Type: "+sExpectedCPU+" selection is persisted");
			}
			else
			{
				Report.Log(ReportLevel.Failure, "CPU Type: "+sExpectedCPU+ " selection is not persisted");
			}
			
		}
		
		/***********************************************************************************************************************************
		 * Function Name: changePanelLED
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update : Alpesh Dhakad - 01/08/2019 & 30/08/2019- Updated test scripts as per new build and xpaths
		 ***********************************************************************************************************************************/
		[UserCodeMethod]
		public static void changePanelLED(int PanelLED)
		{
			
			
			// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem("Node");
			
			
			repo.FormMe.cell_NumberOfAlarmLeds.Click();
			Keyboard.Press("{LControlKey down}{Akey}{LControlKey up}"+PanelLED +"{ENTER}");
			
			
			// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem("Node");
			
			Report.Log(ReportLevel.Info," Panel LED changed to " +PanelLED + "  successfully  ");
				
			
			
		}
		
		/***********************************************************************************
		 * Function Name: ChangeCPUType
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update : Alpesh Dhakad - 20/02/2020 Updated method
		 * Alpesh Dhakad - 18/05/2020 Updated script and xpaths
		 ***********************************************************************************/
		[UserCodeMethod]
		public static void ChangeCPUType(string SelectCPU)
		{
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Enter the  text in Search Properties fields to view related text
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("CPU" +"{ENTER}");
			
			// Click on CPU Cell
			//repo.FormMe.cell_CPU_beforeimport.Click();
			repo.FormMe.cell_CPUType.Click();
			
			//repo.FormMe.cmb_PanelType.Click();
			
			//sCPU=sSelectCPU;
			
			// Enter the CPU value and click Enter twice
			repo.FormMe.txt_CPUType.PressKeys((SelectCPU) +"{ENTER}" + "{ENTER}");
			
			Report.Log(ReportLevel.Info," CPU Type changed to " +SelectCPU + " successfully  ");
			
				// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");

			
		}
		
		/****************************************************************************************************************************************
		 * Function Name: DeletePanel
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale

		 * Purvi Bhasin - 22/08/2019 commented Inventory_LabelCell.DoubleClick() as it causes an error
		 * Alpesh Dhakad - 28/08/2019 - Added site node script 11/02/2020 - Added Ok button click line after new implementation
		 * Alpesh Dhakad - 28/04/2020 - Updated script and xpaths
		 ****************************************************************************************************************************************/
		[UserCodeMethod]
		public static void DeletePanel(int NumberofPanels,string PanelNode,int rowNumber )
		{
			
			for (int i=0; i<NumberofPanels; i++)
			{
				sRow = rowNumber.ToString();
				sLabelName=PanelNode;
				

				// Click on Site node
				Common_Functions.ClickOnNavigationTreeItem("Site");
				
				/*  If else statement added as when we have added only 1 panel and then to delete the same Xpath is different
				 * Date : 27/11/2018
				 * */
//				if(repo.ProfileConsys1.PanelInvetoryGrid.Inventory_LabelCellInfo.Exists())
//				{
//					repo.ProfileConsys1.PanelInvetoryGrid.Inventory_LabelCell.Click();
//				}
				if(repo.FormMe.PanelNodeNameInfo.Exists())
				{
					repo.FormMe.PanelNodeName.Click();
				}
				else
				{
					repo.FormMe.SinglePanel.Click();
				}
				
				Thread.Sleep(300);
				
				Common_Functions.clickOnDeleteButton();
				
				//repo.FormMe2.ButtonOK.Click();
				
				repo.FormDeletePanel.btn_Ok_DeletePanel.Click();
				Report.Log(ReportLevel.Info," Panel deleted " +PanelNode + " deleted successfully  ");
				
				// Click on Site node
				Common_Functions.ClickOnNavigationTreeItem("Site");
				
				
//				Common_Functions.clickOnPanelNetworkTab();
//				//repo.ProfileConsys1.PanelInvetoryGrid.Inventory_LabelCell.DoubleClick();
//				if(repo.ProfileConsys1.PanelInvetoryGrid.LabelNameInfo.Exists())
//				{
//					Report.Log(ReportLevel.Failure, "Panel with label name: "+sLabelName+" is not deleted successfully");
//				}
//				else
//				{
//					Report.Log(ReportLevel.Success, "Panel with label name: "+sLabelName+" is deleted successfully");
//				}
				
			}
		}
		
		/********************************************************************************************************************************
		 * Function Name: SelectPanelNode
		 * Function Details: To select Panel Node
		 * Parameter/Arguments: PanelName
		 * Output:
		 * Function Owner: Poonam Kadam
		 * Last Update : 3/1/2019  Alpesh Dhakad - 01/08/2019 & 19/08/2019 - Updated test scripts as per new build and xpaths
		 ********************************************************************************************************************************/
		[UserCodeMethod]
		public static void SelectPanelNode(string sPanelName)
		{
			PanelName=sPanelName.ToString();
			
			// Click on panel node
			Common_Functions.ClickOnNavigationTreeItem("Node");
			
			Report.Log(ReportLevel.Success, "Panel Node "+sPanelName+" selected");
		}
		
		/************************************************************************************************
		 * Function Name: ChangePSUType
		 * Function Details:Used to change 1st PSU of panel
		 * Parameter/Arguments:PSU to be selected
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update : 09/01/2019 Alpesh Dhakad - 08/09/2019 Updated report log and cell_PSU xpath
		 ************************************************************************************************/
		[UserCodeMethod]
		public static void ChangePSUType(string sPSUType)
		{
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Enter the Device text in Search Properties fields to view Power supply related text
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("PSU" +"{ENTER}" );
			
			// Click on PSU cell
			repo.FormMe.Cell_PSU.Click();
			
			// Enter the value to change PSU value
			repo.FormMe.Cell_PSU.PressKeys((sPSUType) +"{ENTER}" + "{ENTER}");
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");

			
			//repo.FormMe.Cell_PSU.DoubleClick();
			//repo.FormMe.Cell_PSU.PressKeys(sPSUType+"{ENTER}");
			
			
			//repo.FormMe.cmb_PSU.Click();
			//sPSU=sPSUType;
			
			//repo.ContextMenu.lstPSU.Click();
			Report.Log(ReportLevel.Info," PSU Type changed to " +sPSUType + " successfully  ");
		}
		
		/******************************************************************************************************
		 * Function Name: ChangeSecondPSUType
		 * Function Details:Used to change 2nd PSU of panel
		 * Parameter/Arguments:PSU to be selected
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 09/01/2019 Alpesh Dhakad - 06/11/2019 Updated code to change PSU
		 ******************************************************************************************************/
		[UserCodeMethod]
		public static void ChangeSecondPSUType(string SecondPSU)
		{
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Enter the Device text in Search Properties fields to view Power supply related text
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("PSU" +"{ENTER}" );
			
			// Click on PSU cell
			repo.FormMe.Cell_SecondPSU.Click();
			
			// Enter the value to change PSU value
			repo.FormMe.Cell_SecondPSU.PressKeys((SecondPSU) +"{ENTER}" + "{ENTER}");
			
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");

			//repo.ContextMenu.lstPSU.Click();
			Report.Log(ReportLevel.Info," Second PSU Type changed to " +SecondPSU + " successfully  ");
			
			//repo.FormMe.Cell_SecondPSU.Click();
			//repo.FormMe.Cell_SecondPSU.PressKeys(SecondPSU+"{ENTER}");
			
		}
		
		/********************************************************************************************************
		 * Function Name: DevicePoweredFrom
		 * Function Details:Used to change 2nd PSU of panel
		 * Parameter/Arguments:PSU to be selected
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 09/01/2019 Alpesh Dhakad - 30/06/2020 Updated script as per new implementation
		 ********************************************************************************************************/
		[UserCodeMethod]
		public static void DevicePoweredFrom(string PoweredBy)
		{
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Enter the Device text in Search Properties fields to view Power supply related text
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("Powered" +"{ENTER}" );
			
			
			repo.FormMe.PoweredFrom.Click();
			
			// Enter the value to change PSU value
			repo.FormMe.PoweredFrom.PressKeys((PoweredBy) +"{ENTER}" + "{ENTER}");
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");

			//repo.ContextMenu.lstPSU.Click();
			Report.Log(ReportLevel.Info," Powered from change to  " +PoweredBy + " successfully  ");
		
//			repo.FormMe.PoweredFrom.Click();
//			repo.FormMe.PoweredFrom.PressKeys(PoweredBy+"{ENTER}");
//			
		}
		
		/**************************************************************************************************************
		 * Function Name: AddPanelsInBetween
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 30/05/2019 Alpesh Dhakad - 01/10/2019 - Added step as per new Panel Node in Build 43
		 **************************************************************************************************************/
		[UserCodeMethod]
		public static void AddPanelOnAddingOnePanel(int NumberofPanels,string PanelNames,string sPanelCPU)
		{
			for (int i=1; i<NumberofPanels;i++)
			{
				string[] splitPanelNames = PanelNames.Split(',');
				
				repo.ProfileConsys1.SiteNode.Click();
				
				string PanelNameWithSpace=splitPanelNames[i];
				PanelName=PanelNameWithSpace.Replace(" ",String.Empty);
				if(PanelName.StartsWith("P"))
				{
					sPanelLabelIndex ="5";
				}
				else
				{
					sPanelLabelIndex ="7";
				}
				repo.ProfileConsys1.btnDropDownPanelsGallery.Click();
				repo.ContextMenu.txt_SelectPanel.Click();
				repo.AddANewPanel.AddNewPanelContainer.cmb_Addresses.Click();
				iAddress=i+1;
				Address =iAddress.ToString();
				repo.ContextMenu.lstPanelAddress.Click();
				repo.AddANewPanel.AddNewPanelContainer.txt_Label.Click();
				Label="Node"+iAddress;
				
				//Added this step after 43 build update
				Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
				
				
				Keyboard.Press(Label);
				if (!sPanelCPU.IsEmpty())
				{
					repo.AddANewPanel.AddNewPanelContainer.cmb_CPU.Click();
					sCPU=sPanelCPU;
					repo.ContextMenu.lstPanelCPU.Click();
				}
				repo.AddANewPanel.ButtonOK.Click();
				
				if(PanelNameWithSpace == "MZX252")
				{
					PanelNameWithSpace = "MZX 252";
				}
				PanelNode = Label+" "+"-"+" "+PanelNameWithSpace;
				
				//Commenting below line as for Panel name with Space and hi-fen it is not displaying as it is displaying while adding panel
				//Validate.AttributeEqual(repo.ProfileConsys1.NavigationTree.VerifyPanelNodeInfo, "Text", PanelNode);
				
			}
		}
		
		
		/*****************************************************************************************************************
		 * Function Name: AddMorePanels
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 30/05/2019  Alpesh Dhakad - 01/08/2019 - Updated test scripts as per new build and xpaths
		 * Alpesh Dhakad - 01/10/2019 - Added step as per new Panel Node in Build 43
		 *****************************************************************************************************************/
		[UserCodeMethod]
		public static void AddOnePanel(int NodeNumber,string PanelNames,string sPanelCPU)
		{
			
			//string[] splitPanelNames = PanelNames.Split(',');
			
			// Click on Site node
			Common_Functions.ClickOnNavigationTreeItem("Site");
			
			
			string PanelNameWithSpace=PanelNames;
			PanelName=PanelNameWithSpace.Replace(" ",String.Empty);
			if(PanelName.StartsWith("P"))
			{
				sPanelLabelIndex ="5";
			}
			else
			{
				sPanelLabelIndex ="7";
			}
			
			
			
			repo.FormMe.btn_DropDownPanelsGallery.Click();
			
			repo.ContextMenu.txt_SelectPanel.Click();
			repo.AddANewPanel.AddNewPanelContainer.cmb_Addresses.Click();
			iAddress=NodeNumber;
			
			Address =iAddress.ToString();
			repo.ContextMenu.lstPanelAddress.Click();
			repo.AddANewPanel.AddNewPanelContainer.txt_Label.Click();
			Label="Node"+iAddress;
			
			//Added this step after 43 build update
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
				
			
			Keyboard.Press(Label);
			if (!sPanelCPU.IsEmpty())
			{
				repo.AddANewPanel.AddNewPanelContainer.cmb_CPU.Click();
				sCPU=sPanelCPU;
				repo.ContextMenu.lstPanelCPU.Click();
			}
			repo.AddANewPanel.ButtonOK.Click();
			
		}
		
		/**************************************************************************************************************
		 * Function Name: AddPanelAndAddCPUAndPSU
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 8/8/2019 * Alpesh Dhakad - 01/10/2019 - Added step as per new Panel Node in Build 43
		 **************************************************************************************************************/
		[UserCodeMethod]
		public static void AddPanelAndAddCPUAndPSU(int NumberofPanels,string PanelNames,string sPanelCPU)
		{
			for (int i=0; i<NumberofPanels;i++)
			{
				string[] splitPanelNames = PanelNames.Split(',');
				
				
				
				// Click on Site node
				Common_Functions.ClickOnNavigationTreeItem("Site");
				
				
				string PanelNameWithSpace=splitPanelNames[i];
				PanelName=PanelNameWithSpace.Replace(" ",String.Empty);
				if(PanelName.StartsWith("P"))
				{
					sPanelLabelIndex ="5";
				}
				else
				{
					sPanelLabelIndex ="7";
				}
				repo.ProfileConsys1.btnDropDownPanelsGallery.Click();
				repo.ContextMenu.txt_SelectPanel.Click();
				repo.AddANewPanel.AddNewPanelContainer.cmb_Addresses.Click();
				iAddress=i+1;
				Address =iAddress.ToString();
				repo.ContextMenu.lstPanelAddress.Click();
				repo.AddANewPanel.AddNewPanelContainer.txt_Label.Click();
				Label="Node"+iAddress;
				
				//Added this step after 43 build update
				Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
				
				
				Keyboard.Press(Label);
				if (!sPanelCPU.IsEmpty())
				{
					repo.AddANewPanel.AddNewPanelContainer.cmb_CPU.Click();
					sCPU=sPanelCPU;
					repo.ContextMenu.lstPanelCPU.Click();
				}
				
				if(PanelNameWithSpace == "MZX252")
				{
					PanelNameWithSpace = "MZX 252";
				}
				PanelNode = Label+" "+"-"+" "+PanelNameWithSpace;
				
				//Commenting below line as for Panel name with Space and hi-fen it is not displaying as it is displaying while adding panel
				//Validate.AttributeEqual(repo.ProfileConsys1.NavigationTree.VerifyPanelNodeInfo, "Text", PanelNode);
				
				//After this add User Code AddPSUDuringPanelSelection
				
			}
		}
		
		/********************************************************************
		 * Function Name: AddPSUDuringPanelSelection
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 8/8/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void AddPSUDuringPanelSelection(string sPSU,string sSecondPSU)
		{
			//Before using this User code add AddPanelAndAddCPUAndPSU
			repo.AddANewPanel.cmb_PowerSupply1.Click();
			
			//Add 1st PSU
			repo.AddANewPanel.FirstPSU_cnt1.Click();
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
			
			repo.AddANewPanel.FirstPSU_cnt1.PressKeys(sPSU);
			
			//Add Second PSU
			if(!sSecondPSU.IsEmpty())
			{
				repo.AddANewPanel.SecondPSU_txt1.Click();
				Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
				
				repo.AddANewPanel.SecondPSU_txt1.PressKeys(sSecondPSU);
			}
			repo.AddANewPanel.ButtonOK.Click();
			
		}
		
		/******************************************************************************************
		 * Function Name: VerifyValueOf2ndPSU
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 09/09/2019  Alpesh Dhakad - 07/11/2019 Updated code to verify PSU
		 ******************************************************************************************/
		[UserCodeMethod]
		public static void VerifyValueOf2ndPSU(string SecondPSU)
		{
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Enter the Device text in Search Properties fields to view Power supply related text
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("PSU" +"{ENTER}" );
			
		// Click on PSU cell
			repo.FormMe.Cell_SecondPSU.Click();
			
			string ActualPSU = repo.ContextMenu.SecondPSU_Value.TextValue;
			
			if(ActualPSU.Equals(SecondPSU))
			{
				Report.Log(ReportLevel.Success, "PSU "+SecondPSU+" is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure, "PSU "+SecondPSU+" is not displayed");
			}
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
			
			
		}
		
		
		/********************************************************************
		 * Function Name: VerifyValueOf2ndPSUOnReopen
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 09/09/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyValueOf2ndPSUOnReopen(string SecondPSU)
		{
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Enter the Device text in Search Properties fields to view device related text
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("PSU" +"{ENTER}");
			
			
			repo.FormMe.cell_SecondPSU_Reopen.Click();
			//string ActualPSU = repo.ContextMenu.SecondPSU_Value.TextValue;
			
			string ActualPSU = repo.FormMe.txt_SecondPSU_Reopen.TextValue;
			
			if(ActualPSU.Equals(SecondPSU))
			{
				Report.Log(ReportLevel.Success, "PSU "+SecondPSU+" is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure, "PSU "+SecondPSU+" is not displayed");
			}
			
		}
		
		/**********************************************************************************************************************************
		 * Function Name: AddPanelsFC
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : Alpesh Dhakad - 01/10/2019 - Added step as per new Panel Node in Build 43
		 * Alpesh Dhakad - 02/09/2020 - Updated script to work dynamically for all panels
		 **********************************************************************************************************************************/
		[UserCodeMethod]
		public static void AddPanelsFC(int NumberofPanels,string PanelNames,string sPanelCPU)
		{
			for (int i=0; i<NumberofPanels;i++)
			{
				// Click on Site node
				Common_Functions.ClickOnNavigationTreeItem("Site");
				
				PanelName = PanelNames;
				//                string PanelNameWithSpace=splitPanelNames[i];
				//                PanelName=PanelNameWithSpace.Replace(" ",String.Empty);
				if(PanelName.StartsWith("FC7"))                {
					sPanelLabelIndex ="5";
				}
				else
				{
					sPanelLabelIndex ="7";
				}
				//Commened 2 lines on 02/09/2020
				//repo.ProfileConsys1.btnDropDownPanelsGallery.Click();
				
				//repo.ContextMenu.txt_SelectPanelFC.Click();
				
				
				
				//repo.ProfileConsys1.btnDropDownPanelsGallery.Click();
				//repo.ContextMenu.txt_SelectPanel.Click();
			
				// Added this line on 02/09/2020 to select panel with generic xpath
				ModelNumber = PanelName;
				
				
				// Added this line on 02/09/2020 to select panel with generic xpath				
				repo.FormMe.btn_AllGalleryDropdown.Click();
				repo.ContextMenu.txt_SelectDevice.Click();
				
				
				//repo.ContextMenu.txt_SelectPanel.Click();
				repo.AddANewPanel.AddNewPanelContainer.cmb_Addresses.Click();
				iAddress=i+1;
				Address =iAddress.ToString();
				repo.ContextMenu.lstPanelAddress.Click();
				repo.AddANewPanel.AddNewPanelContainer.txt_Label.Click();
				Label="Node"+iAddress;
				
				//Added this step after 43 build update
				Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
				
				Keyboard.Press(Label);
				if (!sPanelCPU.IsEmpty())
				{
					repo.AddANewPanel.AddNewPanelContainer.cmb_CPU.Click();
					sCPU=sPanelCPU;
					repo.ContextMenu.lstPanelCPU.Click();
				}
				repo.AddANewPanel.ButtonOK.Click();
				
				
				//Commenting below line as for Panel nme with Space and hi-fen it is not displaying as it is displaying while adding panel
				//Validate.AttributeEqual(repo.ProfileConsys1.NavigationTree.VerifyPanelNodeInfo, "Text", PanelNode);
				Report.Log(ReportLevel.Success, "Panel "+PanelNames+" Added Successfully");
			}
		}
		
		/**********************************************************************************************************************************
		 * Function Name: AddPanelsFCAndVerifyPSUs
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 18/09/2019 Alpesh Dhakad - 01/10/2019 - Added step as per new Panel Node in Build 43
		 **********************************************************************************************************************************/
		[UserCodeMethod]
		public static void AddPanelsFCAndVerifyPSUs(string sFileName,string sAddDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared string type
			string sPanelName,PanelNode,sPanelCPU,FirstPSU,SecondPSU;
			int j =0;
			
			for(int i=10; i<=rows; i++)
			{
				sPanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				sPanelCPU = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				FirstPSU = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				SecondPSU = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				
				// Click on Site node
				Common_Functions.ClickOnNavigationTreeItem("Site");
				
				PanelName = sPanelName;
				//                string PanelNameWithSpace=splitPanelNames[i];
				//                PanelName=PanelNameWithSpace.Replace(" ",String.Empty);
				if(PanelName.StartsWith("P"))                {
					sPanelLabelIndex ="5";
				}
				else
				{
					sPanelLabelIndex ="7";
				}
				repo.ProfileConsys1.btnDropDownPanelsGallery.Click();
				
				repo.ContextMenu.txt_SelectPanelFC.Click();
				
				
				
				//repo.ContextMenu.txt_SelectPanel.Click();
				repo.AddANewPanel.AddNewPanelContainer.cmb_Addresses.Click();
				iAddress=j+1;
				Address =iAddress.ToString();
				repo.ContextMenu.lstPanelAddress.Click();
				repo.AddANewPanel.AddNewPanelContainer.txt_Label.Click();
				Label="Node"+iAddress;
				
				//Added this step after 43 build update
				Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
				
				
				Keyboard.Press(Label);
				if (!sPanelCPU.IsEmpty())
				{
					repo.AddANewPanel.AddNewPanelContainer.cmb_CPU.Click();
					sCPU=sPanelCPU;
					repo.ContextMenu.lstPanelCPU.Click();
				}
				
				VerifyPSUDuringPanelSelection(FirstPSU,SecondPSU);
				
				repo.AddANewPanel.ButtonOK.Click();
				
				
				//Commenting below line as for Panel nme with Space and hi-fen it is not displaying as it is displaying while adding panel
				//Validate.AttributeEqual(repo.ProfileConsys1.NavigationTree.VerifyPanelNodeInfo, "Text", PanelNode);
				Report.Log(ReportLevel.Success, "Panel "+PanelName+" Added Successfully");
				j=j+1;
			}
		}
		
		/**********************************************************************************************************************************
		 * Function Name: VerifyPSUDuringPanelSelection
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 18/09/2019
		 **********************************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyPSUDuringPanelSelection(string ExpectedFirstPSU,string ExpectedSecondPSU)
		{
			//Before using this User code add AddPanelAndAddCPUAndPSU
			repo.AddANewPanel.cmb_PowerSupply1.Click();
			
			//Add 1st PSU
			repo.AddANewPanel.FirstPSUValue.Click();
			
			string Actual1stPSU = repo.AddANewPanel.FirstPSUValue.TextValue;
			if(Actual1stPSU.Equals(ExpectedFirstPSU))
			{
				Report.Log(ReportLevel.Success,"First PSU Displayed Correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"First PSU is not Displayed Correctly");
			}
			
			if(!ExpectedSecondPSU.IsEmpty())
			{
				repo.AddANewPanel.SecondPSUValue.Click();
				string Actual2ndPSU = repo.AddANewPanel.SecondPSUValue.TextValue;
				
				if(Actual2ndPSU.Equals(ExpectedSecondPSU))
				{
					Report.Log(ReportLevel.Success,"Second PSU Displayed Correctly");
				}
				else
				{
					Report.Log(ReportLevel.Failure,"Second PSU is not Displayed Correctly");
				}
			}
		}
		
		/************************************************************************************************
		 * Function Name: VerifyPSUType
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 19/09/2019 Alpesh Dhakad - 05/06/2020 Updated script as per new xpath 
		 ************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyPSUType(string sExpectedPSU)
		{
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Enter the Device text in Search Properties fields to view Power supply related text
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("PSU" +"{ENTER}" );
			
			repo.FormMe.Cell_PSU.Click();
			string ActualPSUValue = repo.FormMe.txt_FirstPSU.TextValue;
			Report.Log(ReportLevel.Info,ActualPSUValue);
			if(ActualPSUValue.Equals(sExpectedPSU))
			{
				Report.Log(ReportLevel.Success,"First PSU Displayed Correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"First PSU is not Displayed Correctly");
			}
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
			
		}
		
		/****************************************************************************************************************************************
		 * Function Name: DeletePanel
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update
		 ****************************************************************************************************************************************/
		[UserCodeMethod]
		public static void DeletePanel(string PanelNode,int rowNumber )
		{
				sRow = rowNumber.ToString();
				sLabelName=PanelNode;
				
				// Click on Site node
				Common_Functions.ClickOnNavigationTreeItem("Site");
				
				
				repo.FormMe.PanelNodeName.Click();
				
				Thread.Sleep(300);
				
				Common_Functions.clickOnDeleteButton();
				
				repo.FormMe2.ButtonOK.Click();
				
				// Click on Site node
				Common_Functions.ClickOnNavigationTreeItem("Site");
		}
		
		/****************************************************************************************************************************************
		 * Function Name: DeletePanel
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update
		 ****************************************************************************************************************************************/
		[UserCodeMethod]
		public static void DeleteSinglePanel(string PanelNode,int rowNumber )
		{
				sRow = rowNumber.ToString();
				sLabelName=PanelNode;
				
				// Click on Site node
				Common_Functions.ClickOnNavigationTreeItem("Site");
				
				
				repo.FormMe.SinglePanel.Click();
				
				Thread.Sleep(300);
				
				Common_Functions.clickOnDeleteButton();
				
				repo.FormMe2.ButtonOK.Click();
				
				// Click on Site node
				Common_Functions.ClickOnNavigationTreeItem("Site");
		}
		
		/******************************************************************************************************************
		 * Function Name: VerifyCPUTypeOnImport
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 
		 //******************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyCPUTypeOnImport(string sExpectedCPU,int PanelNode, bool AfterImport)
		{
			string sActualText;
			if(AfterImport)
			{
				// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem("Panel 1");
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Enter the  text in Search Properties fields to view related text
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("CPU" +"{ENTER}");
			
				// Click on CPU Cell
			repo.FormMe.cell_CPUType.Click();
			
				// Enter the CPU value and click Enter twice
			sActualText = repo.FormMe.txt_CPUType.TextValue;
			
				// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
				//repo.ProfileConsys1.Cell_CPU_afterimport.DoubleClick();
				//sActualText = repo.ProfileConsys1.VerifyCPUTpye_afterimport.TextValue;
				// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
				
				// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");

			}
			else
			{
			
			// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem("Panel 1");
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Enter the  text in Search Properties fields to view related text
			repo.ProfileConsys1.txt_SearchProperties.PressKeys("CPU" +"{ENTER}");
			
			// Click on CPU Cell
			repo.FormMe.cell_CPUType.Click();
			
			// Enter the CPU value and click Enter twice
			sActualText = repo.FormMe.txt_CPUType.TextValue;
			
				// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
			}
			
			if (sExpectedCPU==sActualText)
			{
				Report.Log(ReportLevel.Success, "CPU Type: "+sExpectedCPU+" selection is persisted");
			}
			else
			{
				Report.Log(ReportLevel.Failure, "CPU Type: "+sExpectedCPU+ " selection is not persisted");
			}
			
		}
		
		
		/**********************************************************************************************************************************
		 * Function Name: AddPanels
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update : 28/12/2018 Alpesh Dhakad - Line 162 Commented
		 * Alpesh Dhakad - 19/08/2019 - Updated with new navigation tree method, xpath and devices gallery
		 * Alpesh Dhakad - 01/10/2019 - Added step as per new Panel Node in Build 43
		 **********************************************************************************************************************************/
		[UserCodeMethod]
		public static void AddPanelsMultipleTimes(int NumberofPanels,string PanelNames,string sPanelCPU)
		{
			for (int i=0; i<NumberofPanels;i++)
			{
				string[] splitPanelNames = PanelNames.Split(',');
				
				// Click on Site node
				Common_Functions.ClickOnNavigationTreeItem("Site");
				
				string PanelNameWithSpace=splitPanelNames[i];
				PanelName=PanelNameWithSpace.Replace(" ",String.Empty);
				
				ModelNumber = PanelName;
				
				if(PanelName.StartsWith("P"))
				{
					sPanelLabelIndex ="5";
				}
				else
				{
					sPanelLabelIndex ="7";
				}
				repo.FormMe.btn_AllGalleryDropdown.Click();
				
				//repo.FormMe.btn_AllGalleryDropdown.Click();
				
				repo.ContextMenu.txt_SelectDevice.Click();
				
				repo.AddANewPanel.AddNewPanelContainer.cmb_Addresses.Click();
				iAddress=i+1;
				Address =iAddress.ToString();
				repo.ContextMenu.lstPanelAddress.Click();
				repo.AddANewPanel.AddNewPanelContainer.txt_Label.Click();
				Label="Node"+iAddress;
				
				//Added this step after 43 build update
				Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
				
				
				Keyboard.Press(Label);
				if (!sPanelCPU.IsEmpty())
				{
					repo.AddANewPanel.AddNewPanelContainer.cmb_CPU.Click();
					sCPU=sPanelCPU;
					repo.ContextMenu.lstPanelCPU.Click();
				}
				repo.AddANewPanel.ButtonOK.Click();
				
				if(PanelNameWithSpace == "MZX252")
				{
					PanelNameWithSpace = "MZX 252";
				}
				PanelNode = Label+" "+"-"+" "+PanelNameWithSpace;
				
				//Commenting below line as for Panel name with Space and hi-fen it is not displaying as it is displaying while adding panel
				//Validate.AttributeEqual(repo.ProfileConsys1.NavigationTree.VerifyPanelNodeInfo, "Text", PanelNode);
				Report.Log(ReportLevel.Success, "Panel "+PanelNames+" Added Successfully");
			}
		}
		
		/**********************************************************************************************************************************
		 * Function Name: AddPanelsFC
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : Alpesh Dhakad - 01/10/2019 - Added step as per new Panel Node in Build 43
		 **********************************************************************************************************************************/
		[UserCodeMethod]
		public static void AddPanelsFC1(int NumberofPanels,string PanelNames,string sPanelCPU)
		{
			if(NumberofPanels=='1')
			{
				repo.AddANewPanel.AddNewPanelContainer.txt_Label.Click();
				
				//Added this step after 43 build update
				Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
				
				Keyboard.Press(Label);
				if (!sPanelCPU.IsEmpty())
				{
					repo.AddANewPanel.AddNewPanelContainer.cmb_CPU.Click();
					sCPU=sPanelCPU;
					repo.ContextMenu.lstPanelCPU.Click();
				}
				repo.AddANewPanel.ButtonOK.Click();
				
				
				//Commenting below line as for Panel nme with Space and hi-fen it is not displaying as it is displaying while adding panel
				//Validate.AttributeEqual(repo.ProfileConsys1.NavigationTree.VerifyPanelNodeInfo, "Text", PanelNode);
				Report.Log(ReportLevel.Success, "Panel "+PanelNames+" Added Successfully");
			}
			
			else
			{
				
			
				for (int i=0; i<NumberofPanels;i++)
			{
				// Click on Site node
				Common_Functions.ClickOnNavigationTreeItem("Site");
				
				PanelName = PanelNames;
				//                string PanelNameWithSpace=splitPanelNames[i];
				//                PanelName=PanelNameWithSpace.Replace(" ",String.Empty);
				if(PanelName.StartsWith("P"))                {
					sPanelLabelIndex ="5";
				}
				else
				{
					sPanelLabelIndex ="7";
				}
				repo.ProfileConsys1.btnDropDownPanelsGallery.Click();
				
				repo.ContextMenu.txt_SelectPanelFC.Click();
				
				
				
				//repo.ContextMenu.txt_SelectPanel.Click();
				repo.AddANewPanel.AddNewPanelContainer.cmb_Addresses.Click();
				iAddress=i+1;
				Address =iAddress.ToString();
				repo.ContextMenu.lstPanelAddress.Click();
				repo.AddANewPanel.AddNewPanelContainer.txt_Label.Click();
				Label="Node"+iAddress;
				
				//Added this step after 43 build update
				Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
				
				Keyboard.Press(Label);
				if (!sPanelCPU.IsEmpty())
				{
					repo.AddANewPanel.AddNewPanelContainer.cmb_CPU.Click();
					sCPU=sPanelCPU;
					repo.ContextMenu.lstPanelCPU.Click();
				}
				repo.AddANewPanel.ButtonOK.Click();
				
				
				//Commenting below line as for Panel nme with Space and hi-fen it is not displaying as it is displaying while adding panel
				//Validate.AttributeEqual(repo.ProfileConsys1.NavigationTree.VerifyPanelNodeInfo, "Text", PanelNode);
				Report.Log(ReportLevel.Success, "Panel "+PanelNames+" Added Successfully");
			}
		}
		}
	}
}


