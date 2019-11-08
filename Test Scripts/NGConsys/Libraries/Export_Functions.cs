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
		
		static string sValue
		{
			get { return repo.sSKU; }
			set { repo.sSKU = value; }
		}
		
		/***********************************************************************************************************
		 * Function Name: ExportAndGenerateShoppingListInExcelFormat
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 06/06/2019 17/07/2019 - Alpesh Dhakad - Updated code
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
			
			//*****************17/07/2019 - Alpesh Dhakad - Updated code ***********************************
			// Click on OK Button of export document
			Export_Functions.validateAndClickOKButtonOnExportDocument();
			
//				// Click on OK button of export document
//				repo.ExportDocument.ButtonOK.Click();
//				Delay.Milliseconds(200);
//
//				// Click on OK button of export document again
//				repo.ExportDocument.ButtonOK.Click();
			
			//*****************17/07/2019 - Alpesh Dhakad - Updated code ***********************************
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
		

		/***********************************************************************************************************
		 * Function Name: validateAndClickOKButtonOnExportDocument
		 * Function Details: To validate export document window and then click on Ok after verification
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update :  17/07/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void validateAndClickOKButtonOnExportDocument()
		{
			try{
//			//repo.ExportDocument.PARTDragWidget.Click();
//
//			if(repo.FormMe2.ButtonOKInfo.Exists())
//			{
//				repo.FormMe2.ButtonOK.Click();
//				Delay.Milliseconds(200);
//
//				//repo.FormMe2.ButtonOK.Click(); updated Purvi-23/08/2019
//			}
//			else
//			{
//
			repo.ExportDocument.ButtonOK.Click();
			Delay.Milliseconds(200);
			
			repo.ExportDocument.ButtonOK.Click();
			//}
			}catch(Exception e)
			{
				Report.Log(ReportLevel.Info,"Exception occured. Ok button is not displayed"+e.Message);
			}
		}
		
		/***********************************************************************************************************
		 * Function Name: SearchDeviceInExportUsingSKUOrDescription
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 09/08/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void SearchDeviceInExportUsingSKUOrDescription(string sValue,bool sExist)
		{
			
			// Click on Export button
			repo.FormMe.Export2ndTime.Click();
			Delay.Milliseconds(200);
			
			//Click on Search button
			repo.PrintPreview.SearchExport1.Click();
			
			//Click on Search bar
			repo.PrintPreview.SearchBox_Export1.Click();
			
			//Enter the required Device's SKU no
			Keyboard.Press(sValue +"{ENTER}");
			
			if(sExist)
			{
				
				string ActualValue = repo.PrintPreview.txt_ExportResult.TextValue;
				
				repo.PrintPreview.txt_ExportResult.Click();
				
				if(ActualValue.Equals(sValue))
				{
					Report.Log(ReportLevel.Success,"Device with SKU "+sValue+" is displayed correctly");
				}
				else
				{
					Report.Log(ReportLevel.Failure,"Device with SKU "+sValue+" is not getting displayed");
				}
			}
			else
			{
				if(repo.PrintPreview.NoMatches_InExportInfo.Exists())
				{
					Report.Log(ReportLevel.Success,"Device with SKU "+sValue+" is not getting displayed");
				}
				else
				{
					Report.Log(ReportLevel.Failure,"Device with SKU "+sValue+" is getting displayed");
				}
			}
			repo.PrintPreview.btn_CloseB.Click();
			Delay.Milliseconds(200);
		}
		
		/***********************************************************************************************************
		 * Function Name:
		 * Function Details:
		 * Parameter/Arguments: The output file already exists. Click OK to overwrite.
		 * Output:
		 * Function Owner: Alpesh Dhakad
		 * Last Update : 26/08/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyOverwriteMessageAndClickButton(string sExpectedText)
		{
			if(repo.Export.Msg_OverwriteInfo.Exists())
			{
				
				string actualText = repo.Export.Msg_Overwrite.TextValue;
				
				if(actualText.Equals(sExpectedText))
				{
					Report.Log(ReportLevel.Info,"Overwrite message " +actualText+ " is displayed");
					repo.Export.ButtonOK.Click();
				}
				else
				{
					Report.Log(ReportLevel.Info,"Overwrite message is not displayed");
				}
			}
		}
		
		/***********************************************************************************************************
		 * Function Name: VerifyOverwriteConfirmationForExcel
		 * Function Details:
		 * Parameter/Arguments:
		 * Output:
		 * Function Owner: Purvi Bhasin
		 * Last Update : 28/08/2019
		 ************************************************************************************************************/
		[UserCodeMethod]
		public static void VerifyOverwriteConfirmationForExcel()
		{
			if(repo.FormMe2.Export_OK_OverwriteInfo.Exists())
			{
				repo.FormMe2.Export_OK_Overwrite.Click();
			}
			else
			{
				//Delay.Milliseconds(100);
				validateAndClickOKButtonOnExportDocument();
			}
			
		}
	}
}

