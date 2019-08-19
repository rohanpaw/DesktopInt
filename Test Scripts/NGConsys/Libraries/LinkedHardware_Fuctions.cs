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
        
        /********************************************************************
		 * Function Name: VerifyLinkedDevicesGetAddedInLoop
		 * Function Details: Add a device and its child till Max Limit and check if linked devices get added in the Loop
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 9/08/2019 
		 ********************************************************************/
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
				PanelType = ((Range)Excel_Utilities.ExcelRange.Cells[i,5]).Value.ToString();
				
				//Click Panel Node
				repo.FormMe.PanelNode1.Click();
				
				//Add parent Device
				Devices_Functions.AddDevicesfromPanelNodeGallery(ParentDeviceName,ParentDeviceType,PanelType);
				
				//Select row of parent Device
				Devices_Functions.SelectRowUsingLabelName(ParentLabel);
				
				
				if(!ChildDeviceName.IsEmpty())
				{
					//Select Child Device
					Devices_Functions.AddDevicesfromPanelNodeGallery(ChildDeviceName,ChildDeviceType,PanelType);
				}
				
			}
			//Verify Linked Devices are added in Loop A
			
			//Expand Panel Node
			repo.FormMe.NodeExpander1.Click();
			
			//Expand Loop Card
			repo.FormMe.LoopExpander1.Click();
			
			//Click on Loop A
			repo.FormMe.Loop_A1.Click();
			
			//Click Points Tab
			
			for(int i=8; i<=10; i++)
			{
				expectedLabel1 = ((Range)Excel_Utilities.ExcelRange.Cells[i,6]).Value.ToString();
				expectedLabel2 = ((Range)Excel_Utilities.ExcelRange.Cells[i,7]).Value.ToString();
				
				Devices_Functions.VerifyDeviceUsingLabelName(expectedLabel1);
				Devices_Functions.VerifyDeviceUsingLabelName(expectedLabel1);
			}
				
		}
    }
}
