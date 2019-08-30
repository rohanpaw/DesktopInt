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
using System.Text;
using System.IO;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Threading;
using WinForms = System.Windows.Forms;

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
			repo.FormMe.PanelNode1.Click();
			
		}
		
		/**********************************************************************************************************************************
		 * Function Name: AddPanels
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update : 28/12/2018 Alpesh Dhakad - Line 162 Commented
		 * Alpesh Dhakad - 19/08/2019 - Updated with new navigation tree method, xpath and devices gallery 
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
		 ******************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyCPUType(string sExpectedCPU,int PanelNode, bool AfterImport)
		{
			string sActualText;
		
			
			// Click on Panel node
			Common_Functions.ClickOnNavigationTreeItem("Node");
			
			if(AfterImport)
			{
				repo.ProfileConsys1.Cell_CPU_afterimport.DoubleClick();
				sActualText = repo.ProfileConsys1.VerifyCPUTpye_afterimport.TextValue;
			}
			else
			{
				repo.ProfileConsys1.Cell_CPU_beforeimport.DoubleClick();
				sActualText = repo.ProfileConsys1.VerifyCPUTpye_beforeimport.TextValue;
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
		
		/********************************************************************
		 * Function Name: changePanelLED
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update : Alpesh Dhakad - 01/08/2019 - Updated test scripts as per new build and xpaths
		 ********************************************************************/
		[UserCodeMethod]
		public static void changePanelLED(int PanelLED)
		{
			
			repo.FormMe.NodeExpander1.Click();
			
			repo.ProfileConsys1.cell_NumberOfAlarmLeds.Click();
			Keyboard.Press("{LControlKey down}{Akey}{LControlKey up}"+PanelLED +"{ENTER}");
			
			repo.FormMe.NodeExpander1.Click();
			
		}
		
		/********************************************************************
		 * Function Name: ChangeCPUType
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void ChangeCPUType(string sSelectCPU)
		{
			repo.ProfileConsys1.Cell_CPU_beforeimport.DoubleClick();
			repo.ProfileConsys1.cmb_PanelType.Click();
			sCPU=sSelectCPU;
			
			repo.ContextMenu.lstPanelCPU.Click();
			
			Report.Log(ReportLevel.Info," CPU Type changed to " +sSelectCPU + " successfully  ");
		}
		
		/****************************************************************************************************************************************
		 * Function Name: DeletePanel
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale

		 * Purvi Bhasin - 22/08/2019 commented Inventory_LabelCell.DoubleClick() as it causes an error
		 * Alpesh Dhakad - 28/08/2019 - Added site node script 
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
				if(repo.ProfileConsys1.PanelInvetoryGrid.Inventory_LabelCellInfo.Exists())
				{
					repo.ProfileConsys1.PanelInvetoryGrid.Inventory_LabelCell.Click();
				}
				else
				{
					repo.FormMe.SinglePanel.Click();
				}
				
				repo.ProfileConsys1.btn_Delete.Click();
				
				// Click on Site node
				Common_Functions.ClickOnNavigationTreeItem("Site");
				
				
				repo.FormMe.tab_Panel_Network.Click();
				//repo.ProfileConsys1.PanelInvetoryGrid.Inventory_LabelCell.DoubleClick();
				if(repo.ProfileConsys1.PanelInvetoryGrid.LabelNameInfo.Exists())
				{
					Report.Log(ReportLevel.Failure, "Panel with label name: "+sLabelName+" is not deleted successfully");
				}
				else
				{
					Report.Log(ReportLevel.Success, "Panel with label name: "+sLabelName+" is deleted successfully");
				}
				
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
		
		/********************************************************************
		 * Function Name: ChangePSUType
		 * Function Details:Used to change 1st PSU of panel
		 * Parameter/Arguments:PSU to be selected
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update : 09/01/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void ChangePSUType(string sPSUType)
		{
			repo.FormMe.Cell_PSU.DoubleClick();
			repo.FormMe.Cell_PSU.PressKeys(sPSUType+"{ENTER}");
			//repo.FormMe.cmb_PSU.Click();
			//sPSU=sPSUType;
			
			//repo.ContextMenu.lstPSU.Click();
		}
		
		/********************************************************************
		 * Function Name: ChangeSecondPSUType
		 * Function Details:Used to change 2nd PSU of panel
		 * Parameter/Arguments:PSU to be selected
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 09/01/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void ChangeSecondPSUType(string SecondPSU)
		{
			repo.FormMe.Cell_SecondPSU.DoubleClick();
			repo.FormMe.Cell_SecondPSU.PressKeys(SecondPSU+"{ENTER}");
			
		}
		
		/********************************************************************
		 * Function Name: DevicePoweredFrom
		 * Function Details:Used to change 2nd PSU of panel
		 * Parameter/Arguments:PSU to be selected
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 09/01/2019
		 ********************************************************************/
		public static void DevicePoweredFrom(string PoweredBy)
		{
			repo.FormMe.PoweredFrom.Click();
			repo.FormMe.PoweredFrom.PressKeys(PoweredBy+"{ENTER}");
			
		}
		
		/********************************************************************
		 * Function Name: AddPanelsInBetween
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 30/05/2019
		 ********************************************************************/
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
		
		
		/********************************************************************
		 * Function Name: AddMorePanels
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 30/05/2019  Alpesh Dhakad - 01/08/2019 - Updated test scripts as per new build and xpaths
		 ********************************************************************/
		[UserCodeMethod]
		public static void AddOnePanel(int NodeNumber,string PanelNames,string sPanelCPU)
		{
			
			//string[] splitPanelNames = PanelNames.Split(',');
			
			repo.FormMe.SiteNode1.Click();
			
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
			Keyboard.Press(Label);
			if (!sPanelCPU.IsEmpty())
			{
				repo.AddANewPanel.AddNewPanelContainer.cmb_CPU.Click();
				sCPU=sPanelCPU;
				repo.ContextMenu.lstPanelCPU.Click();
			}
			repo.AddANewPanel.ButtonOK.Click();
			
		}
		
		/********************************************************************
		 * Function Name: AddPanelAndAddCPUAndPSU
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 8/8/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void AddPanelAndAddCPUAndPSU(int NumberofPanels,string PanelNames,string sPanelCPU)
		{
			for (int i=0; i<NumberofPanels;i++)
			{
				string[] splitPanelNames = PanelNames.Split(',');
				
				
				
				repo.FormMe.SiteNode1.Click();
				
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
		
		/********************************************************************
		 * Function Name: VerifyValueOf2ndPSU
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 09/09/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyValueOf2ndPSU(string SecondPSU)
		{
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
			repo.FormMe.cell_SecondPSU_Reopen.Click();
			string ActualPSU = repo.ContextMenu.SecondPSU_Value.TextValue;
			
			if(ActualPSU.Equals(SecondPSU))
			{
				Report.Log(ReportLevel.Success, "PSU "+SecondPSU+" is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure, "PSU "+SecondPSU+" is not displayed");
			}
			
		}
		
			
	}
}


