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
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void SaveProject(string sProjectName)
		{
			if (repo.ProfileConsys1.btn_Save.Enabled)
			{
				repo.ProfileConsys1.btn_Save.Click();
				
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
			repo.ProfileConsys1.btn_Close.Click();
			
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
		
		
	}
}


