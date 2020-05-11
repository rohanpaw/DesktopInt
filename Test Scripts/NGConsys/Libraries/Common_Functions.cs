/*
 * Created by Ranorex
 * User: jbhosash
 * Date: 5/22/2018
 * Time: 3:20 PM
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
using System.IO;
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
	public class Common_Functions
	{
		//Create instance of repository to access repository items
		static NGConsysRepository repo = NGConsysRepository.Instance;
		
		static string sExpanderName
		{
			get { return repo.sExpanderName; }
			set { repo.sExpanderName = value; }
		}
		
		static string sTreeItem
		{
			get { return repo.sTreeItem; }
			set { repo.sTreeItem = value; }
		}
		
		static string sProjectName
		{
			get { return repo.sProjectName; }
			set { repo.sProjectName = value; }
		}
		
		static string sRow
		{
			get { return repo.sRow; }
			set { repo.sRow = value; }
		}
		
		static string sButtonName
		{
			get { return repo.sButtonName; }
			set { repo.sButtonName = value; }
		}
		
		/********************************************************************
		 * Function Name: GetDirPath
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static string GetDirPath()
		{
			string dirPath = Directory.GetCurrentDirectory();
			
			string[] splitPath = dirPath.Split('\\');
			
			string actualDirPath = string.Empty;
			for (int i = 0; i < splitPath.Length && !actualDirPath.Contains("consys-uiauto"); i++)
			{
				actualDirPath = actualDirPath + splitPath[i] + "\\";
			}
			
			return actualDirPath;
		}
		
		/********************************************************************
		 * Function Name: SaveProject
		 * Function Details: To save project
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update : 15/09/2019 - Alpesh Dhakad Updated btn_Save Xpath
		 ********************************************************************/
		[UserCodeMethod]
		public static void SaveProject(string sProjectName)
		{
			repo.FormMe.File.Click();
			
			repo.FormMe.Save.Click();
			
			repo.ProjectChangeDescription.txt_Desc.Click();
					Keyboard.Press("Automation....");
					repo.ProjectChangeDescription.btn_OK.Click();
			
			
			
			
//			if (repo.FormMe.btn_Save.Enabled)
//			{
//				repo.FormMe.btn_Save.Click();
//				
//				if(repo.ProjectChangeDescription.btn_OK.Visible)
//				{
//					repo.ProjectChangeDescription.txt_Desc.Click();
//					Keyboard.Press("Automation....");
//					repo.ProjectChangeDescription.btn_OK.Click();
//				}
				
				if(repo.SaveConfirmationWindow.ButtonSave.Visible)
				{
					
					string actualDirPath= Common_Functions.GetDirPath();
					string sSaveProjectDirPath = actualDirPath+ "NGDesigner Saved Projects";
					repo.SaveConfirmationWindow.Save_Open_Window.Click();
					sProjectName= sSaveProjectDirPath + "\\"+ sProjectName;
					repo.SaveConfirmationWindow.txt_Path.PressKeys(sProjectName);
					//repo.SaveConfirmationWindow.txt_Path.PressKeys("{Return}");
					
					//   	repo.SaveConfirmationWindow.txt_ProjectName.Click();
					//	repo.SaveConfirmationWindow.txt_ProjectName.PressKeys(sProjectName);
					repo.SaveConfirmationWindow.ButtonSave.Click();
					
				}
//			}
//			else
//			{
//				Report.Log(ReportLevel.Failure, "Save button is not enabled");
//			}
//			}catch(Exception e){
//				Report.Log(ReportLevel.Info, "Exception occured"+e.Message);
//			}
		}
		
		
		/********************************************************************
		 * Function Name: ReopenProject
		 * Function Details: To reopen project
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void ReopenProject(string sProjectName)
		{
			repo.ProfileConsys1.File.Click();
			Delay.Duration(1000);
			Keyboard.Press("{LControlKey down}{Okey}{LControlKey up}");
			Delay.Duration(1000);
			repo.Open.btn_Open.Click();
			//repo.ProfileConsys1.txt_Open.Click();
//			string actualDirPath= Common_Functions.GetDirPath();
//			string sSaveProjectDirPath = actualDirPath+ "NGDesigner Saved Projects";
//			repo.Open.PreviousLocations.Click();
//
//			repo.Open.txt_EnterPath.PressKeys(sSaveProjectDirPath);
//			repo.Open.txt_EnterPath.PressKeys("{Return}");
			
			repo.Open.txt_EnterProjectName.Click();
			repo.Open.txt_EnterProjectName.PressKeys(sProjectName);
			repo.Open.btn_Open.Click();
		}
		
		/********************************************************************
		 * Function Name: Application_Close
		 * Function Details: To close application
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update : 30/11/2018 by Devendra Kulkarni
		 ********************************************************************/
		[UserCodeMethod]
		public static void Application_Close(bool Save, bool SaveConfirmation, string sProjectName)
		{
			//repo.ProfileConsys1.btn_Close.Click();
			
			repo.FormMe.btn_Close1.Click();
			//repo.FormMe.btn_Close.Click();
			
			
			if (Save)
			{
				if(repo.SaveConfirmationWindow.SelfInfo.Exists())
				{
					repo.SaveConfirmationWindow.btnYes_SaveConfirmation.Click();
					Report.Log(ReportLevel.Success, "Save confirmation asked");
					
					if(repo.ProjectChangeDescription.SelfInfo.Exists())
					{
						repo.ProjectChangeDescription.txt_Desc.Click();
						Keyboard.Press("Automation....");
						repo.ProjectChangeDescription.btn_OK.Click();
					}
					
					if(repo.SaveConfirmationWindow.SelfInfo.Exists())
					{
						
//						string actualDirPath= Common_Functions.GetDirPath();
//						Console.WriteLine("PAth:" + actualDirPath);
//						string sSaveProjectDirPath = actualDirPath+ "NGDesigner Saved Projects";
//						repo.SaveConfirmationWindow.Btn_PreviousLocations.Click();
//						repo.SaveConfirmationWindow.txt_Path.PressKeys("{Return}");
//						repo.SaveConfirmationWindow.txt_Path.PressKeys(sSaveProjectDirPath);
//						repo.SaveConfirmationWindow.txt_Path.PressKeys("{Return}");
						
						repo.SaveConfirmationWindow.txt_ProjectName.Click();
						repo.SaveConfirmationWindow.txt_ProjectName.PressKeys(sProjectName);
						repo.SaveConfirmationWindow.ButtonSave.Click();
						
					}
				}
				else
				{
					Report.Log(ReportLevel.Failure, "Save confirmation not asked");
				}
				
			}
			
			else
			{
				if(SaveConfirmation)
				{
					if(repo.SaveConfirmationWindow.btnNo_SaveConfirmationInfo.Exists())
					{
						repo.SaveConfirmationWindow.btnNo_SaveConfirmation.Click();
						Report.Log(ReportLevel.Success, "Save confirmation asked");
						
					}
					else
					{
						Report.Log(ReportLevel.Failure, "Save confirmation not asked");
					}
				}
				else
				{
					Report.Log(ReportLevel.Success, "Save confirmation not asked");
				}
				
			}
			
		}

		/********************************************************************
		 * Function Name: Import_MxDesignerProject
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void Import_MxDesignerProject(string sProjectName)
		{
			repo.ProfileConsys1.File.Click();
			repo.ProfileConsys1.txt_Import.Click();
			repo.ProfileConsys1.txt_DesignerDataFile.Click();
			string actualDirPath= Common_Functions.GetDirPath();
			string sSaveProjectDirPath = actualDirPath+ "MxDesigner Saved Projects";
			repo.Open.File_Open_Window.Click();
			repo.Open.txt_EnterPath.PressKeys(sSaveProjectDirPath);
			repo.Open.txt_EnterPath.PressKeys("{Return}");
			
			repo.Open.txt_EnterProjectName.Click();
			repo.Open.txt_EnterProjectName.PressKeys(sProjectName);
			repo.Open.btn_Open.Click();
		}
		
		
		/********************************************************************
		 * Function Name: CloseProject
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void CloseProject(bool SaveConfirmation, string sProjectName)
		{
			repo.ProfileConsys1.File.Click();
			repo.ProfileConsys1.txt_Close.Click();
			if (SaveConfirmation)
			{
				if(repo.SaveConfirmationWindow.SelfInfo.Exists())
				{
					repo.SaveConfirmationWindow.btnYes_SaveConfirmation.Click();
					Report.Log(ReportLevel.Success, "Save confirmation asked");
					
					if(repo.ProjectChangeDescription.SelfInfo.Exists())
					{
						repo.ProjectChangeDescription.txt_Desc.Click();
						Keyboard.Press("Automation....");
						repo.ProjectChangeDescription.btn_OK.Click();
					}
					
					if(repo.SaveConfirmationWindow.SelfInfo.Exists())
					{
						
						string actualDirPath= Common_Functions.GetDirPath();
						string sSaveProjectDirPath = actualDirPath+ "NGDesigner Saved Projects";
						repo.SaveConfirmationWindow.Save_Open_Window.Click();
						repo.SaveConfirmationWindow.txt_Path.PressKeys(sSaveProjectDirPath);
						repo.SaveConfirmationWindow.txt_Path.PressKeys("{Return}");
						
						repo.SaveConfirmationWindow.txt_ProjectName.Click();
						repo.SaveConfirmationWindow.txt_ProjectName.PressKeys(sProjectName);
						repo.SaveConfirmationWindow.ButtonSave.Click();
						
					}
				}
				else
				{
					Report.Log(ReportLevel.Failure, "Save confirmation not asked");
				}
				
			}
			
			else
			{
				if(repo.SaveConfirmationWindow.SelfInfo.Exists())
				{
					repo.SaveConfirmationWindow.btnNo_SaveConfirmation.Click();
					Report.Log(ReportLevel.Success, "Save confirmation asked");
					
				}
				else
				{
					Report.Log(ReportLevel.Failure, "Save confirmation not asked");
				}
				
			}
			
		}
		
		/********************************************************************
		 * Function Name: SaveProjectFromFileOption
		 * Function Details: To save project from File->Save
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Poonam Kadam
		 * Last Update :04/06/19
		 ********************************************************************/
		[UserCodeMethod]
		public static void SaveProjectFromFileOption(string sProjectName)
		{
			if (repo.ProfileConsys1.File.Enabled)
			{
				repo.ProfileConsys1.File.Click();
				repo.FormMe.Save.Click();
				
				if(repo.ProjectChangeDescription.btn_OK.Visible)
				{
					repo.ProjectChangeDescription.txt_Desc.Click();
					Keyboard.Press("Automation....");
					repo.ProjectChangeDescription.btn_OK.Click();
				}
				
				if(repo.SaveConfirmationWindow.ButtonSave.Visible)
				{
					
					string actualDirPath= Common_Functions.GetDirPath();
					string sSaveProjectDirPath = actualDirPath+ "NGDesigner Saved Projects";
					repo.SaveConfirmationWindow.Save_Open_Window.Click();
					sProjectName= sSaveProjectDirPath + "\\"+ sProjectName;
					repo.SaveConfirmationWindow.txt_Path.PressKeys(sProjectName);
					//repo.SaveConfirmationWindow.txt_Path.PressKeys("{Return}");
					
					//   	repo.SaveConfirmationWindow.txt_ProjectName.Click();
					//	repo.SaveConfirmationWindow.txt_ProjectName.PressKeys(sProjectName);
					repo.SaveConfirmationWindow.ButtonSave.Click();
					
				}
			}
			else
			{
				Report.Log(ReportLevel.Failure, "Save button is not enabled");
			}
		}
		
		
		/****************************************************************************************************
		 * Function Name: ClickOnNavigationTreeExpanderButton
		 * Function Details: To click on navigation tree item expander button
		 * Parameter/Arguments: Tree item Expander button name
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 13/08/2019
		 ****************************************************************************************************/
		[UserCodeMethod]
		public static void ClickOnNavigationTreeExpander(string ExpanderName)
		{
			sExpanderName = ExpanderName;
			repo.FormMe.btn_NavigationTreeExpander.Click();
			Report.Log(ReportLevel.Info," Tree Item with ExpanderName name " +ExpanderName + " is displayed and clicked successfully ");
		}
		
		
		/****************************************************************************************************
		 * Function Name: ClickOnNavigationTreeItem
		 * Function Details: To click on navigation tree item
		 * Parameter/Arguments: Tree item name
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 13/08/2019
		 ****************************************************************************************************/
		[UserCodeMethod]
		public static void ClickOnNavigationTreeItem(string TreeItemName)
		{
			sTreeItem = TreeItemName;

			repo.FormMe.NavigationTreeItem.Click();
			Report.Log(ReportLevel.Info," Tree Item name " +TreeItemName + " is displayed and clicked successfully  ");
			
			
			
		}
		
		/****************************************************************************************************
		 * Function Name: VerifyNavigationTreeItemText
		 * Function Details: To verify navigation tree item text
		 * Parameter/Arguments: Tree Item name
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 13/08/2019
		 ****************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyNavigationTreeItemText(string TreeItemName)
		{
			sTreeItem = TreeItemName;
			string ActualText = repo.FormMe.txt_NavigationTreeItem.TextValue;
			
			if(ActualText.Equals(TreeItemName))
			{
				Report.Log(ReportLevel.Success,"Tree Item " +ActualText+ " text is as displayed as expected");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Tree Item text is displayed as " +ActualText+ "instead of " +TreeItemName);
			}
		}
		
		
		
		/****************************************************************************************************
		 * Function Name: VerifyAndClickNavigationTreeItemText
		 * Function Details: To verify navigation tree item text and then click on it
		 * Parameter/Arguments: Tree Item name
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 13/08/2019
		 ****************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyAndClickNavigationTreeItemText(string TreeItemName)
		{
			sTreeItem = TreeItemName;
			string ActualText = repo.FormMe.txt_NavigationTreeItem.TextValue;
			
			
			if(ActualText.Equals(TreeItemName))
			{
				Report.Log(ReportLevel.Success,"Tree Item " +ActualText+ " text is as displayed as expected");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Tree Item text is displayed as " +ActualText+ " instead of " +TreeItemName);
			}
			repo.FormMe.txt_NavigationTreeItem.Click();
		}
		
		
		/****************************************************************************************************
		 * Function Name: VerifyNavigationTreeItemText
		 * Function Details: To verify navigation tree Item and verify text is visible or not
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 13/08/2019
		 ****************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyNavigationTreeItem(string TreeItemName, string Visibility)
		{
			sTreeItem = TreeItemName;
			
			bool sVisibility = Convert.ToBoolean(Visibility);
			if(sVisibility)
			{
				if(repo.FormMe.NavigationTreeItemInfo.Exists())
				{
					Report.Log(ReportLevel.Success, "Tree Item " +TreeItemName+ " text is displayed");
				}
				else
				{
					Report.Log(ReportLevel.Failure, "Tree Item " +TreeItemName+ " text is not displayed");
				}
			}
			else
			{
				if(repo.FormMe.NavigationTreeItemInfo.Exists())
				{
					Report.Log(ReportLevel.Failure, "Tree Item " +TreeItemName+ " text is displayed");
				}
				else
				{
					Report.Log(ReportLevel.Success, "Tree Item " +TreeItemName+ " text is not displayed ");
				}
				
			}
			
			
		}
		
		/********************************************************************
		 * Function Name: VerifyProjectName
		 * Function Details: To verify if project name for a new project and a saved project
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 19/8/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyProjectName(string ExpectedProjectName)
		{
			
			sProjectName = ExpectedProjectName;
			repo.FormMe.Project_Name.Click();
			//string ActualProjectName = repo.FormMe.Project_Name.TextValue;
			//Report.Log(ReportLevel.Info,"The name of the project displayed is "+ActualProjectName);
			//ActualProjectName.Equals(ExpectedProjectName)
			if(repo.FormMe.Project_NameInfo.Exists())
			{
				Report.Log(ReportLevel.Success, "Project Name " +ExpectedProjectName+ " is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure, "Project Name " +ExpectedProjectName+ " text is not displayed correctly");
			}
		}
		
		
		/********************************************************************
		 * Function Name: VerifyProjectNameForDifferentPanels
		 * Function Details: To verify if project name for a new project and a saved project for different Panels
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 20/8/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyProjectNameForDifferentPanels(string sFileName,string sAddPanelSheet, string sAddDevicesSheet)
		{
			//Open excel sheet and read it values,
			Excel_Utilities.OpenExcelFile(sFileName,sAddPanelSheet);
			
			// Count number of rows in excel and store it in rows variable
			int rows= Excel_Utilities.ExcelRange.Rows.Count;
			
			// Declared string type
			string PanelName,PanelNode,CPUType,PanelType,sExpectedProjectNameBefore,sSaveName,sExpectedProjectNameAfter,sModelName,sType;
			
			
			// For loop to iterate on data present in excel
			for(int i=8; i<=rows; i++)
			{
				PanelName =  ((Range)Excel_Utilities.ExcelRange.Cells[i,1]).Value.ToString();
				PanelNode = ((Range)Excel_Utilities.ExcelRange.Cells[i,2]).Value.ToString();
				CPUType = ((Range)Excel_Utilities.ExcelRange.Cells[i,3]).Value.ToString();
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[i,4]).Value.ToString();
				sExpectedProjectNameBefore = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				sSaveName = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				sExpectedProjectNameAfter = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				
				
				
				// Add panels using test data in excel sheet
				Panel_Functions.AddPanels(1,PanelName,CPUType);
				
				//Verify Project Name
				VerifyProjectName(sExpectedProjectNameBefore);
				
				//Save Project
				SaveProject(sSaveName);
				
				//Reopen the project
				ReopenProject(sSaveName);
				
				//Verify saved project name
				VerifyProjectName(sExpectedProjectNameAfter);
				
				//Close Panel sheet and open device sheet
				Excel_Utilities.CloseExcel();
				Excel_Utilities.OpenExcelFile(sFileName,sAddDevicesSheet);
				
				// Count number of rows in excel and store it in rows variable
				int rows2= Excel_Utilities.ExcelRange.Rows.Count;
				
				for(int j=2;j<=rows2;j++)
				{
					sModelName = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
					sType = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
					
					//Expand Panel Node
					ClickOnNavigationTreeExpander("Node");
					
					//Expand PFI/FIM Loop card
					//ClickOnNavigationTreeExpander(PanelType);
					
					//Click on Loop A
					ClickOnNavigationTreeItem("Built-in Loop-A");
					
					//Add Devices
					Devices_Functions.AddDevicesfromGallery(sModelName,sType);
					
				}
				
				//Save Project
				repo.ProfileConsys1.btn_Save.Click();
				if(repo.ProjectChangeDescription.btn_OK.Visible)
				{
					repo.ProjectChangeDescription.txt_Desc.Click();
					Keyboard.Press("Automation....");
					repo.ProjectChangeDescription.btn_OK.Click();
				}
				else
				{
					Report.Log(ReportLevel.Failure,"Description not asked");
				}
				
				//Reopen the project
				ReopenProject(sSaveName);
				
				//Verify saved project name
				VerifyProjectName(sExpectedProjectNameAfter);
			}
			
			Devices_Functions.CreateProject("United Kingdom",2);
		}
		
		/********************************************************************
		 * Function Name: VerifyElementVisibilityInNavigationTree
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 22/08/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyElementVisibilityInNavigationTree(bool sExists, string TreeItemName)
		{
			sTreeItem = TreeItemName;
			if(sExists)
			{

				if(repo.FormMe.NavigationTreeItemInfo.Exists())
				{
					Report.Log(ReportLevel.Success, "Element "+TreeItemName+" is displayed correctly");
				}
				else
				{
					Report.Log(ReportLevel.Failure, "Element "+TreeItemName+" is not displayed");
				}
			}
			else
			{
				if(repo.FormMe.NavigationTreeItemInfo.Exists())
				{
					Report.Log(ReportLevel.Failure, "Element "+TreeItemName+" is getting displayed");
				}
				else
				{
					Report.Log(ReportLevel.Success, "Element "+TreeItemName+" is not displayed as expected");
				}
			}
		}
		
		/********************************************************************
		 * Function Name: ApplicationCloseUsingCloseInFile
		 * Function Details: To close application
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 22/08/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void ApplicationCloseUsingCloseInFile(bool Save, bool SaveConfirmation, string sProjectName)
		{
			//repo.ProfileConsys1.btn_Close.Click();
			repo.ProfileConsys1.File.Click();
			repo.FormMe.Close_In_File.Click();
			//repo.FormMe.btn_Close.Click();
			
			
			if (Save)
			{
				if(repo.SaveConfirmationWindow.SelfInfo.Exists())
				{
					repo.SaveConfirmationWindow.btnYes_SaveConfirmation.Click();
					Report.Log(ReportLevel.Success, "Save confirmation asked");
					
					if(repo.ProjectChangeDescription.SelfInfo.Exists())
					{
						repo.ProjectChangeDescription.txt_Desc.Click();
						Keyboard.Press("Automation....");
						repo.ProjectChangeDescription.btn_OK.Click();
					}
					
					if(repo.SaveConfirmationWindow.SelfInfo.Exists())
					{
						
//						string actualDirPath= Common_Functions.GetDirPath();
//						Console.WriteLine("PAth:" + actualDirPath);
//						string sSaveProjectDirPath = actualDirPath+ "NGDesigner Saved Projects";
//						repo.SaveConfirmationWindow.Btn_PreviousLocations.Click();
//						repo.SaveConfirmationWindow.txt_Path.PressKeys("{Return}");
//						repo.SaveConfirmationWindow.txt_Path.PressKeys(sSaveProjectDirPath);
//						repo.SaveConfirmationWindow.txt_Path.PressKeys("{Return}");
						
						repo.SaveConfirmationWindow.txt_ProjectName.Click();
						repo.SaveConfirmationWindow.txt_ProjectName.PressKeys(sProjectName);
						repo.SaveConfirmationWindow.ButtonSave.Click();
						
					}
				}
				else
				{
					Report.Log(ReportLevel.Failure, "Save confirmation not asked");
				}
				
			}
			
			else
			{
				if(SaveConfirmation)
				{
					if(repo.SaveConfirmationWindow.btnNo_SaveConfirmationInfo.Exists())
					{
						repo.SaveConfirmationWindow.btnNo_SaveConfirmation.Click();
						Report.Log(ReportLevel.Success, "Save confirmation asked");
						
					}
					else
					{
						Report.Log(ReportLevel.Failure, "Save confirmation not asked");
					}
				}
				else
				{
					Report.Log(ReportLevel.Success, "Save confirmation not asked");
				}
				
			}
			
		}
		
		/********************************************************************
		 * Function Name: CreateProjectFCParameters
		 * Function Details: Enter Project Name, Client Name etc. 
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 10/09/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void CreateProjectFCParameters(string RowParameter, string sValue)
		{
			sRow = RowParameter;//starts with 8 for verify then 9,10 and so on
			repo.CreateNewProject.CreateNewProjectContainer.CreateProject_FC_Parameter.Click();
			repo.CreateNewProject.CreateNewProjectContainer.CreateProject_FC_Parameter.PressKeys(sValue);
				
		}
		
		/********************************************************************
		 * Function Name: ReopenProjectForDifferentProjectType
		 * Function Details: To reopen project for different project typ eg jpl
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update :10	/09/2019
		 ********************************************************************/
		[UserCodeMethod]
		public static void ReopenProjectForDifferentProjectType(string ProjectName,string sProjectType)
		{
			sProjectName = sProjectType; 
			repo.ProfileConsys1.File.Click();
			Delay.Duration(1000);
			Keyboard.Press("{LControlKey down}{Okey}{LControlKey up}");
			Delay.Duration(1000);
			//repo.ProfileConsys1.txt_Open.Click();
//			string actualDirPath= Common_Functions.GetDirPath();
//			string sSaveProjectDirPath = actualDirPath+ "NGDesigner Saved Projects";
//			repo.Open.PreviousLocations.Click();
//
//			repo.Open.txt_EnterPath.PressKeys(sSaveProjectDirPath);
//			repo.Open.txt_EnterPath.PressKeys("{Return}");
			
			repo.Open.txt_EnterProjectName.Click();
			repo.Open.txt_EnterProjectName.PressKeys(ProjectName);
			repo.Open.FileType_Expander.Click();
			repo.List1000.ProjectType.Click();
			repo.Open.btn_Open.Click();
		}
		
		/********************************************************************
		 * Function Name: changeConfiguratonToUIA
		 * Function Details: Change Configuration to UIA to detect objects
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :28/01/2020
		 ********************************************************************/
		
		[UserCodeMethod]
		public static void changeConfiguratonToUIA()
		{
			
						
			Ranorex.Plugin.WpfConfiguration.WpfApplicationTrees = Ranorex.Plugin.WpfTreeSelection.WpfImprovedOnly;

			Ranorex.Plugin.WpfConfiguration.WpfApplicationTrees = Ranorex.Plugin.WpfTreeSelection.UiaOnly;
		}
		
		/********************************************************************
		 * Function Name: ReopenProjectForDifferentProjectType
		 * Function Details: Change Configuration to WpfImprovedOnly (Default)
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :28/01/2020
		 ********************************************************************/
		[UserCodeMethod]
		public static void changeConfiguratonToWPF()
		{
			
			Ranorex.Plugin.WpfConfiguration.WpfApplicationTrees = Ranorex.Plugin.WpfTreeSelection.UiaOnly;

			Ranorex.Plugin.WpfConfiguration.WpfApplicationTrees = Ranorex.Plugin.WpfTreeSelection.WpfImprovedOnly;
		}
		
		/********************************************************************
		 * Function Name: maximizeApplication
		 * Function Details: To maximize application if it is not expanded
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :30/01/2020
		 ********************************************************************/
		
		[UserCodeMethod]
		public static void maximizeApplication()
		{
			if(repo.FormMe.btn_MaximizeInfo.Exists())
			{
				repo.FormMe.btn_Maximize.Click();
			}
		}
		
		
		/********************************************************************
		 * Function Name: clickOnButton
		 * Function Details: 
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :3/02/2020
		 ********************************************************************/
		
		[UserCodeMethod]
		public static void clickOnButton(string ButtonName)
		{
		//	changeConfiguratonToUIA();
			
			sButtonName = ButtonName; 
			repo.FormMe.btnNameMainGallery.Click();
			
		//	changeConfiguratonToWPF();
		}
		
		
		/********************************************************************
		 * Function Name: verifyButtonState
		 * Function Details: 
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :3/02/2020
		 ********************************************************************/
		
		[UserCodeMethod]
		public static void verifyButtonState(bool isReadOnly,string ButtonName)
		{
			Common_Functions.changeConfiguratonToUIA();
			
			
			
			sButtonName = ButtonName; 
			
			repo.FormMe.btnNameMainGallery.Click();
			bool result = repo.FormMe.btnNameMainGallery.Enabled;
			
			if(result==isReadOnly)
			{
				Report.Log(ReportLevel.Success,"Delete button displaying is as expected");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Delete button displaying is not as expected");
			}
			
			Common_Functions.changeConfiguratonToWPF();
		}
		
		
		/********************************************************************
		 * Function Name: verifyButtonState
		 * Function Details: 
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :3/02/2020
		 ********************************************************************/
		
		[UserCodeMethod]
		public static void verifyDeleteButtonState(bool isReadOnly)
		{
			//Common_Functions.changeConfiguratonToUIA();
			
			
			
			//sButtonName = ButtonName; 
			
			
			bool result = repo.FormMe.btn_Delete.Enabled;
			
			if(result==isReadOnly)
			{
				Report.Log(ReportLevel.Success,"Delete button displaying is as expected");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Delete button displaying is not as expected");
			}
			
			//Common_Functions.changeConfiguratonToWPF();
		}
		
		/********************************************************************
		 * Function Name: clickOnInventoryTab
		 * Function Details: 
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :11/02/2020
		 ********************************************************************/
		
		[UserCodeMethod]
		public static void clickOnInventoryTab()
		{
			repo.FormMe.tab_Inventory.Click();
			Report.Log(ReportLevel.Info," Inventory tab clicked successfully  ");
		}
		
		/********************************************************************
		 * Function Name: clickOnPanelAccessoriesTab
		 * Function Details: 
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :11/02/2020
		 ********************************************************************/
		
		[UserCodeMethod]
		public static void clickOnPanelAccessoriesTab()
		{
			repo.FormMe.tab_PanelAccessories.Click();
			Report.Log(ReportLevel.Info," Panel Accessories tab clicked successfully  ");
		}
		
		/********************************************************************
		 * Function Name: clickOnSiteAccessoriesTab
		 * Function Details: 
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :11/02/2020
		 ********************************************************************/
		
		[UserCodeMethod]
		public static void clickOnSiteAccessoriesTab()
		{
			repo.FormMe.tab_SiteAccessories.Click();
			Report.Log(ReportLevel.Info," Site Accessories tab clicked successfully  ");
		}
	
		/********************************************************************
		 * Function Name: clickOnShoppingListTab
		 * Function Details: 
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :11/02/2020
		 ********************************************************************/
		
		[UserCodeMethod]
		public static void clickOnShoppingListTab()
		{
			repo.FormMe.tab_ShoppingList.Click();
			Report.Log(ReportLevel.Info," Shopping list tab clicked successfully  ");
		}
		
		
		/********************************************************************
		 * Function Name: clickOnPanelNetworkTab
		 * Function Details: 
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :11/02/2020
		 ********************************************************************/
		
		[UserCodeMethod]
		public static void clickOnPanelNetworkTab()
		{
			repo.FormMe.tab_Panel_Network.Click();
			Report.Log(ReportLevel.Info,"Panel Network tab clicked successfully  ");
		}
		
		/********************************************************************
		 * Function Name: clickOnPointsTab
		 * Function Details: 
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :11/02/2020
		 ********************************************************************/
		
		[UserCodeMethod]
		public static void clickOnPointsTab()
		{
			repo.ProfileConsys1.tab_Points.Click();
			Report.Log(ReportLevel.Info," Points tab clicked successfully  ");
		}
		
		/********************************************************************
		 * Function Name: clickOnPhysicalLayoutTab
		 * Function Details: 
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :11/02/2020
		 ********************************************************************/
		
		[UserCodeMethod]
		public static void clickOnPhysicalLayoutTab()
		{
			repo.ProfileConsys1.tab_PhysicalLayout.Click();
			Report.Log(ReportLevel.Info," Physical Layout clicked successfully  ");
		}
		
		/********************************************************************
		 * Function Name: clickOnDeleteButton
		 * Function Details: 
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :06/02/2020
		 ********************************************************************/
		
		[UserCodeMethod]
		public static void clickOnDeleteButton()
		{
			repo.FormMe.btn_Delete.Click();
			Report.Log(ReportLevel.Info," Delete button clicked successfully  ");
		}
		
		/********************************************************************
		 * Function Name: clickOnCutButton
		 * Function Details: 
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :13/02/2020
		 ********************************************************************/
		
		[UserCodeMethod]
		public static void clickOnCutButton()
		{
			repo.FormMe.btn_Cut.Click();
			Report.Log(ReportLevel.Info," Cut button clicked successfully  ");
		}
		
		/********************************************************************
		 * Function Name: clickOnPasteButton
		 * Function Details: 
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :13/02/2020
		 ********************************************************************/
		
		[UserCodeMethod]
		public static void clickOnPasteButton()
		{
			repo.FormMe.btn_Paste.Click();
			Report.Log(ReportLevel.Info," Paste button clicked successfully  ");
		}
		
		
		/********************************************************************
		 * Function Name: clickOnCopyButton
		 * Function Details: 
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :13/02/2020
		 ********************************************************************/
		
		[UserCodeMethod]
		public static void clickOnCopyButton()
		{
			repo.FormMe.btn_Copy.Click();
			Report.Log(ReportLevel.Info," Copy button clicked successfully  ");
		}
	

/****************************************************************************************************
		 * Function Name: ClickOnNavigationTreeItem
		 * Function Details: To click on navigation tree item
		 * Parameter/Arguments: Tree item name
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 13/08/2019
		 ****************************************************************************************************/
		[UserCodeMethod]
		public static void ClickOnTreeItem()
		{
//			sTreeItem = TreeItemName;
//
//			repo.FormMe.NavigationTreeItem.Click();
//			Report.Log(ReportLevel.Info," Tree Item name " +TreeItemName + " is displayed and clicked successfully  ");
//			
			//IList<Ranorex.Unknown>
			Ranorex.Container test = repo.FormMe.test_AutomationID;
		}
		
		/********************************************************************
		 * Function Name: clickOnPanelCalculationsTab
		 * Function Details: Click on panel calculations tab
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :30/04/2020
		 ********************************************************************/
		[UserCodeMethod]
		public static void clickOnPanelCalculationsTab()
		{
			repo.FormMe.tab_PanelCalculations.Click();
			Report.Log(ReportLevel.Info," PanelCalculations tab clicked successfully  ");
		}
		
		
		/********************************************************************
		 * Function Name: clickOnPasteWithPointsButton
		 * Function Details: 
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :11/05/2020
		 ********************************************************************/
		
		[UserCodeMethod]
		public static void clickOnPasteWithPointsButton()
		{
			repo.FormMe.btn_PasteWithPoints.Click();
			Report.Log(ReportLevel.Info," Paste with points button clicked successfully  ");
		}
		
	}
}



