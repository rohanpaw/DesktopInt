/*
 * Created by Ranorex
 * User: jdhakaa
 * Date: 6/6/2019
 * Time: 12:18 PM
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
	public class Export_Functions
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
		 * Function Name: ExportAndGenerateShoppingListInExcelFormat
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 06/06/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void ExportAndGenerateShoppingListInExcelFormat()
		{
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
		}
		
		
		/***********************************************************************************************************
		 * Function Name: CloseShoppingListExcel
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 06/06/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void CloseShoppingListExcel()
		{
			// Click to close excel sheet
			repo.ShoppingListCompatibilityModeE.btn_CloseExcel.Click();
			
			// Click on close button
			repo.PrintPreview.btn_CloseB.Click();
			
		}
		
		
		/***********************************************************************************************************
		 * Function Name: verifyShoppingListDevicesTextForCell14
		 * Function Details: To verify shopping list devices via clicking on its row
		 * Parameter/Arguments: sFileName,sDeviceSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 06/06/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyShoppingListDevicesTextForCell14(string sExpectedText)
		{
			repo.ShoppingListCompatibilityModeE.CellF14.Click();
			string actualText = repo.ShoppingListCompatibilityModeE.CellF14.Text;
			
			if(actualText.Equals(sExpectedText))
			{
				Report.Log(ReportLevel.Success,"Model name " +actualText+ " is displayed successfully");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Model name" +sExpectedText+ " is not displayed correctly instead " +actualText+  "is displayed " );
			}
		}
		
		/***********************************************************************************************************
		 * Function Name: verifyShoppingListDevicesTextForCell17
		 * Function Details: To verify shopping list devices via clicking on its row
		 * Parameter/Arguments: sFileName,sDeviceSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 06/06/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyShoppingListDevicesTextForCell17(string sExpectedText)
		{
			repo.ShoppingListCompatibilityModeE.CellF17.Click();
			string actualText = repo.ShoppingListCompatibilityModeE.CellF17.Text;
			
			if(actualText.Equals(sExpectedText))
			{
				Report.Log(ReportLevel.Success,"Model name " +actualText+ " is displayed successfully");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Model name" +sExpectedText+ " is not displayed correctly instead " +actualText+  "is displayed " );
			}
		}
		
		/***********************************************************************************************************
		 * Function Name: verifyShoppingListDevicesTextForCell21
		 * Function Details: To verify shopping list devices via clicking on its row
		 * Parameter/Arguments: sFileName,sDeviceSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 06/06/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyShoppingListDevicesTextForCell21(string sExpectedText)
		{
			repo.ShoppingListCompatibilityModeE.CellF21.Click();
			string actualText = repo.ShoppingListCompatibilityModeE.CellF21.Text;
			
			if(actualText.Equals(sExpectedText))
			{
				Report.Log(ReportLevel.Success,"Model name " +actualText+ " is displayed successfully");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Model name" +sExpectedText+ " is not displayed correctly instead " +actualText+  "is displayed " );
			}
		}
		
		/***********************************************************************************************************
		 * Function Name: verifyShoppingListDevicesTextForCell24
		 * Function Details: To verify shopping list devices via clicking on its row
		 * Parameter/Arguments: sFileName,sDeviceSheet
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 06/06/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void verifyShoppingListDevicesTextForCell24(string sExpectedText)
		{
			repo.ShoppingListCompatibilityModeE.CellF24.Click();
			string actualText = repo.ShoppingListCompatibilityModeE.CellF24.Text;
			
			if(actualText.Equals(sExpectedText))
			{
				Report.Log(ReportLevel.Success,"Model name " +actualText+ " is displayed successfully");
			}
			else
			{
				Report.Log(ReportLevel.Failure,"Model name" +sExpectedText+ " is not displayed correctly instead " +actualText+  "is displayed " );
			}
		}
		
	}
}
