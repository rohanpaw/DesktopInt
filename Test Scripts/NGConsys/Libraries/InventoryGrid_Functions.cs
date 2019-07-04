/*
 * Created by Ranorex
 * User: jbhosash
 * Date: 5/21/2018
 * Time: 6:25 PM
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

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace TestProject.Libraries
{
	/// <summary>
	/// Ranorex user code collection. A collection is used to publish user code methods to the user code library.
	/// </summary>
	[UserCodeCollection]
	public class InventoryGrid_Functions
	{
		
		//Create instance of repository to access repository items
		static NGConsysRepository repo = NGConsysRepository.Instance;
		
		//variables
		static string sRowIndex
		{
			get { return repo.sRowIndex; }
			set { repo.sRowIndex = value; }
		}
		
		static string sSKU
		{
			get { return repo.sSKU; }
			set { repo.sSKU = value; }
		}
		
		static string sColumnIndex
		{
			get { return repo.sColumnIndex; }
			set { repo.sColumnIndex = value; }
		}
		
		
		/// <summary>
		/// This is function used to Verify Inventory Grid
		/// iStartRowIndex: Verification start row
		/// iEndRowIndex: Verification start row
		/// sDesc: Model number of item to verify
		/// </summary>
		/// 
		
		/********************************************************************
		 * Function Name: VerifyInventoryGrid
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyInventoryGrid(int iStartRowIndex, int iEndRowIndex, string sSKUofItem)
		{
			
			for(int i=iStartRowIndex; i<=iEndRowIndex; i++)
			{
				sRowIndex=i.ToString();
				sSKU=sSKUofItem;
				Validate.AttributeEqual(repo.ProfileConsys1.PanelInvetoryGrid.txt_SKUInfo, "Text", sSKU);
			}
		}
		
		/// <summary>
		/// This is function used to select any row of inventory
		/// iRowNumber: Row number
		/// sItemName: Model number of item
		/// </summary>
		/// 
		
		
		/********************************************************************
		 * Function Name: SelectRow
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void SelectRow(int iRowNumber,string sItemName,string sSKUofItem)
		{
			sRowIndex=iRowNumber.ToString();
			
			sSKU=sSKUofItem;
			repo.ProfileConsys1.PanelInvetoryGrid.txt_SKU.Click();
		}

		/// <summary>
		/// This is function used to delete item from inventory grid
		/// iRowNumber: Row number
		/// sItemName: Model number of item
		/// </summary>
		/// 
		
		
		/********************************************************************
		 * Function Name: DeleteItemfromInventory
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void DeleteItemfromInventory(int iRowNumber, string sItemName,string sSKUofItem)
		{
			sSKU=sSKUofItem;
			SelectRow(iRowNumber,sItemName,sSKU);
			repo.ProfileConsys1.btn_Delete.Click();
			
		}
		
		/// <summary>
		/// This is function used to verify item is not present in inventory grid
		/// iRowNumber: Row number
		/// sItemName: Model number of item
		/// </summary>
		/// 
		
		/********************************************************************
		 * Function Name: VerifyRowNotExist
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyRowNotExist(int iRowNumber,string sItemName,string sSKUofItem)
		{
			sRowIndex=iRowNumber.ToString();
			
			sSKU=sSKUofItem;
			if(repo.ProfileConsys1.PanelInvetoryGrid.txt_SKUInfo.Exists())
			{
				Report.Log(ReportLevel.Failure,"Item: "+sItemName+" not deleted successfully");
			}
			else
			{
				Report.Log(ReportLevel.Success,"Item: "+sItemName+" not deleted successfully");
			}
			
		}
		
		/// <summary>
		/// This is function used to verify item is not present in inventory grid
		/// iRowNumber: Row number
		/// sItemName: Model number of item
		/// </summary>
		/// 
		
		/********************************************************************
		 * Function Name: VerifyRowExist
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Shweta Bhosale
		 * Last Update :
		 ********************************************************************/
		[UserCodeMethod]
		public static void VerifyRowExist(int iRowNumber,string sItemName,string sSKUofItem)
		{
			sRowIndex=iRowNumber.ToString();
			sSKU=sSKUofItem;
			if(repo.ProfileConsys1.PanelInvetoryGrid.txt_SKUInfo.Exists())
			{
				Report.Log(ReportLevel.Success,"Item: "+sItemName+" device added successfully");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Item: "+sItemName+" device not added successfully");
			}
			
		}
		
		/***********************************************************************************************************
		 * Function Name: verifyInventoryGridProperties
		 * Function Details: To verify inventory grid properties for a device
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Poonam Kadam
		 * Last Update : 05/04/2019  Updated on 02/07/2019 by Alpesh Dhakad : Updated log reports as readable format
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyInventoryGridProperties(string ExpectedInventoryGridRowCount, string ExpectedInventoryGridColumn, string ExpectedDeviceProperty)
		{
			// Retrieve Point grid rows and column count
			sColumnIndex = getColumnNumberForInventoryGrid(ExpectedInventoryGridColumn);
			sRowIndex = ExpectedInventoryGridRowCount;
			
			Report.Log(ReportLevel.Success,"Column and row values are set as " +sColumnIndex+  " and " +sRowIndex+ " respectively ");
			repo.FormMe.txt_InventoryGridDeviceProperty.Click();
			string ActualInventoryGridProperty = repo.FormMe.txt_InventoryGridDeviceProperty.TextValue;
			
			// Compare Actual and Expected Point Grid Row count
			Report.Log(ReportLevel.Success,"Actual " +ActualInventoryGridProperty+ "  Expected " +ExpectedDeviceProperty);
			
			if(ActualInventoryGridProperty.Equals(ExpectedDeviceProperty))
			{
				Report.Log(ReportLevel.Success,"Device property is verified in Inventory grid");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Device property is not verified not in Inventory grid");
			}
			
		}
		
		/*************************************************************************************************************
		 * Function Name: getColumnNumberForInventoryGrid
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Poonam Kadam
		 * Last Update : 05/04/2019  Updated on 02/07/2019 by Alpesh Dhakad : Updated log reports as readable format
		 **************************************************************************************************************/
		[UserCodeMethod]
		public static string getColumnNumberForInventoryGrid(string columnName)
		{
			string columnNumber="";
			switch (columnName)
			{
				case "SKU":
					columnNumber="1";
					Report.Log(ReportLevel.Success,"Column number is set as " +columnNumber);
					break;
					
				case "Model":
					columnNumber="2";
					Report.Log(ReportLevel.Success,"Column number is set as " +columnNumber);
					break;
					
				case "Label":
					columnNumber="3";
					Report.Log(ReportLevel.Success,"Column number is set as " +columnNumber);
					break;
					
				case "Address":
					columnNumber="4";
					Report.Log(ReportLevel.Success,"Column number is set as " +columnNumber);
					break;
					
				case "Slot Address":
					columnNumber="5";
					Report.Log(ReportLevel.Success,"Column number is set as " +columnNumber);
					break;
					
				case "Connection":
					columnNumber="6";
					Report.Log(ReportLevel.Success,"Column number is set as " +columnNumber);
					break;
					
			}
			return columnNumber;
			
		}
		
		/***********************************************************************************************************
		 * Function Name: editDeviceLabel
		 * Function Details: To edit device label in inventory grid
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Poonam Kadam
		 * Last Update : 05/04/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void editDeviceLabel(string ExpectedInventoryGridRowCount, string ExpectedInventoryGridColumn, string sNewLabel)
		{
			// Retrieve Point grid rows and column count
			sColumnIndex = getColumnNumberForInventoryGrid(ExpectedInventoryGridColumn);
			sRowIndex = ExpectedInventoryGridRowCount;
			
			Report.Log(ReportLevel.Success,"Row an column values are set as" +sColumnIndex  +sRowIndex);
			
			//Modifying the label
			repo.FormMe.txt_InventoryGridDeviceProperty.Click();
			repo.FormMe.txt_InventoryGridDeviceProperty.PressKeys("{LControlKey down}{Akey}{Delete}{LControlKey up}");
			repo.FormMe.txt_InventoryGridDeviceProperty.PressKeys(sNewLabel +"{ENTER}" );
			Report.Log(ReportLevel.Success,"Label is editied to " +sNewLabel);
			
		}
		
		/***********************************************************************************************************
		 * Function Name: verifyInventoryDeviceProperty
		 * Function Details: Verify device property from Inventory properties section
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Poonam Kadam
		 * Last Update : 05/04/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyInventoryDeviceProperty(string sPropertyLabel, string sExpectedValue)
		{
			
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Search for the Label property
			repo.ProfileConsys1.txt_SearchProperties.PressKeys(sPropertyLabel +"{ENTER}" );
			
			// Click on Label property cell
			repo.FormMe.txt_InventoryProperty.Click();
			
			// Get the text value of property
			repo.FormMe.txt_InventoryProperty.Click();
			string actualValue = repo.FormMe.txt_InventoryProperty.TextValue;
			
			Report.Log(ReportLevel.Success,"Actual: "+actualValue+" Expected"+sExpectedValue);
			//Comparing expected and actual Device Sensitivity value
			if(actualValue.Equals(sExpectedValue))
			{
				Report.Log(ReportLevel.Success,"Property value of " +sPropertyLabel + " is displayed correctly");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Property value of " +sPropertyLabel+ " is not displayed correctly");
			}
			// Click on SearchProperties text field
			repo.ProfileConsys1.txt_SearchProperties.Click();
			
			// Select the text in SearchProperties text field and delete it
			Keyboard.Press("{LControlKey down}{Akey}{Delete}{LControlKey up}");
		}
		
		
		/***********************************************************************************************************
		 * Function Name: verifyModelFilterListText
		 * Function Details: To verify Model column filter lists in inventory grid
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Poonam Kadam
		 * Last Update : 05/31/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyModelFilterListText(string sListSize, string sListText)
		{
			Report.Log(ReportLevel.Success,"sColumnIndex " +sColumnIndex);
			string ActualFilterList = "";
			int size=Convert.ToInt32(sListSize);
			string[] strArray = sListText.Split(',');
			for(int i=0;i<size;i++)
			{
				sColumnIndex=(i+1).ToString();
				ActualFilterList=repo.ContextMenu.ModelFilterList.TextValue;
				// Compare Actual and Expected Point Grid Row count
				Report.Log(ReportLevel.Success,"Actual "+ActualFilterList+"  Expected "+strArray[i]);
				
				if(ActualFilterList.Equals(strArray[i]))
				{
					Report.Log(ReportLevel.Success,"Model filter list verified");
				}
				else
				{
					Report.Log(ReportLevel.Failure,"Model filter list not verified");
				}
				
			}
			
		}
		
		
		/***********************************************************************************************************
		 * Function Name: selectModelFilterListText
		 * Function Details: To select Model column filter lists
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Poonam Kadam
		 * Last Update : 05/31/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void selectModelFilterListText(string sColumnNo)
		{
			sColumnIndex=sColumnNo;
			repo.ContextMenu.ModelFilterList.Click();
		}
		
		
	}}