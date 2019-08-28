﻿///////////////////////////////////////////////////////////////////////////////
//
// This file was automatically generated by RANOREX.
// DO NOT MODIFY THIS FILE! It is regenerated by the designer.
// All your modifications will be lost!
// http://www.ranorex.com
//
///////////////////////////////////////////////////////////////////////////////

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
using Ranorex.Core.Repository;

namespace TestProject.Recording_Modules
{
#pragma warning disable 0436 //(CS0436) The type 'type' in 'assembly' conflicts with the imported type 'type2' in 'assembly'. Using the type defined in 'assembly'.
    /// <summary>
    ///The Verify_Deletion_Of_PxD_PSC_800_From_Shopping_List recording.
    /// </summary>
    [TestModule("d67b5e7a-b58a-4f13-b496-83226254abae", ModuleType.Recording, 1)]
    public partial class Verify_Deletion_Of_PxD_PSC_800_From_Shopping_List : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Verify_Deletion_Of_PxD_PSC_800_From_Shopping_List instance = new Verify_Deletion_Of_PxD_PSC_800_From_Shopping_List();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_Deletion_Of_PxD_PSC_800_From_Shopping_List()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_Deletion_Of_PxD_PSC_800_From_Shopping_List Instance
        {
            get { return instance; }
        }

#region Variables

#endregion

        /// <summary>
        /// Starts the replay of the static recording <see cref="Instance"/>.
        /// </summary>
        [System.CodeDom.Compiler.GeneratedCode("Ranorex", "8.3")]
        public static void Start()
        {
            TestModuleRunner.Run(Instance);
        }

        /// <summary>
        /// Performs the playback of actions in this recording.
        /// </summary>
        /// <remarks>You should not call this method directly, instead pass the module
        /// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
        /// that will in turn invoke this method.</remarks>
        [System.CodeDom.Compiler.GeneratedCode("Ranorex", "8.3")]
        void ITestModule.Run()
        {
            Mouse.DefaultMoveTime = 300;
            Keyboard.DefaultKeyPressTime = 100;
            Delay.SpeedFactor = 1.00;

            Init();

            Libraries.Panel_Functions.AddPanels(ValueConverter.ArgumentFromString<int>("NumberofPanels", "1"), "Pro32xD", "");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeExpander("Node");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeExpander("Backplane");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Other Slot Cards");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromMainProcessorGallery("PCH800", "Slot Cards", "PFI");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromMainProcessorGallery("PCH800", "Slot Cards", "PFI");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeExpander("Backplane");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Site");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.ShoppingList' at Center.", repo.FormMe.ShoppingListInfo, new RecordItemIndex(8));
            repo.FormMe.ShoppingList.Click();
            Delay.Milliseconds(200);
            
            Libraries.Common_Functions.ClickOnNavigationTreeExpander("Backplane  2/3");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Other Slot Cards");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromMainProcessorGallery("PCH800", "Slot Cards", "PFI");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeExpander("Backplane  2/3");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeExpander("Backplane  3/3");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Other Slot Cards");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Site");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.ShoppingList' at Center.", repo.FormMe.ShoppingListInfo, new RecordItemIndex(16));
            repo.FormMe.ShoppingList.Click();
            Delay.Milliseconds(200);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Other Slot Cards");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Site");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.ShoppingList' at Center.", repo.FormMe.ShoppingListInfo, new RecordItemIndex(19));
            repo.FormMe.ShoppingList.Click();
            Delay.Milliseconds(200);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Other Slot Cards");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelName("PCH800-3");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Site");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.ShoppingList' at Center.", repo.FormMe.ShoppingListInfo, new RecordItemIndex(23));
            repo.FormMe.ShoppingList.Click();
            Delay.Milliseconds(200);
            
            Libraries.Devices_Functions.verifyShoppingList(ValueConverter.ArgumentFromString<int>("ShoppingListDeviceCount", "6"));
            Delay.Milliseconds(0);
            
            //Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.Export' at Center.", repo.FormMe.ExportInfo, new RecordItemIndex(25));
            //repo.FormMe.Export.Click();
            //Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.Export2ndTime' at Center.", repo.FormMe.Export2ndTimeInfo, new RecordItemIndex(26));
            repo.FormMe.Export2ndTime.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'PrintPreview.PARTMaximize' at Center.", repo.PrintPreview.PARTMaximizeInfo, new RecordItemIndex(27));
            repo.PrintPreview.PARTMaximize.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'PrintPreview.ExportDropdown' at Center.", repo.PrintPreview.ExportDropdownInfo, new RecordItemIndex(28));
            repo.PrintPreview.ExportDropdown.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ExportDocument.ExcelFormat' at Center.", repo.ExportDocument.ExcelFormatInfo, new RecordItemIndex(29));
            repo.ExportDocument.ExcelFormat.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Delay", "Waiting for 5s.", new RecordItemIndex(30));
            Delay.Duration(5000, false);
            
            Report.Log(ReportLevel.Info, "Set value", "Setting attribute Text to 'Xls' on item 'ExportDocument.ExcelFormat'.", repo.ExportDocument.ExcelFormatInfo, new RecordItemIndex(31));
            repo.ExportDocument.ExcelFormat.Element.SetAttributeValue("Text", "Xls");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Delay", "Waiting for 5s.", new RecordItemIndex(32));
            Delay.Duration(5000, false);
            
            //Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ExportDocument.ButtonOK' at Center.", repo.ExportDocument.ButtonOKInfo, new RecordItemIndex(33));
            //repo.ExportDocument.ButtonOK.Click();
            //Delay.Milliseconds(200);
            
            //Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ExportDocument.ButtonOK' at Center.", repo.ExportDocument.ButtonOKInfo, new RecordItemIndex(34));
            //repo.ExportDocument.ButtonOK.Click();
            //Delay.Milliseconds(200);
            
            Libraries.Export_Functions.validateAndClickOKButtonOnExportDocument();
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ExportDocument.ButtonOK' at Center.", repo.ExportDocument.ButtonOKInfo, new RecordItemIndex(36));
            repo.ExportDocument.ButtonOK.Click();
            Delay.Milliseconds(200);
            
            Libraries.Export_Functions.VerifyOverwriteMessageAndClickButton("The output file already exists. Click OK to overwrite.");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ShoppingListCompatibilityModeE.Cell18' at Center.", repo.ShoppingListCompatibilityModeE.Cell18Info, new RecordItemIndex(38));
            repo.ShoppingListCompatibilityModeE.Cell18.Click();
            Delay.Milliseconds(200);
            
            Libraries.Devices_Functions.verifyShoppingListDevicesTextForPxD("PxD");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ShoppingListCompatibilityModeE.CellF22' at Center.", repo.ShoppingListCompatibilityModeE.CellF22Info, new RecordItemIndex(40));
            repo.ShoppingListCompatibilityModeE.CellF22.Click();
            Delay.Milliseconds(200);
            
            Libraries.Devices_Functions.verifyShoppingListDevicesTextForPSC("PSC800");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ShoppingListCompatibilityModeE.btn_CloseExcel' at Center.", repo.ShoppingListCompatibilityModeE.btn_CloseExcelInfo, new RecordItemIndex(42));
            repo.ShoppingListCompatibilityModeE.btn_CloseExcel.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'PrintPreview.btn_Close' at Center.", repo.PrintPreview.btn_CloseInfo, new RecordItemIndex(43));
            repo.PrintPreview.btn_Close.Click();
            Delay.Milliseconds(200);
            
            //Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.Header_ProfileConsys' at Center.", repo.FormMe.Header_ProfileConsysInfo, new RecordItemIndex(44));
            //repo.FormMe.Header_ProfileConsys.Click();
            //Delay.Milliseconds(200);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Other Slot Cards");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelName("PCH800-3");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ProfileConsys1.btn_Delete' at Center.", repo.ProfileConsys1.btn_DeleteInfo, new RecordItemIndex(47));
            repo.ProfileConsys1.btn_Delete.Click();
            Delay.Milliseconds(200);
            
            Libraries.Common_Functions.ClickOnNavigationTreeExpander("Backplane  3/3");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeExpander("Backplane  2/3");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Other Slot Cards");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelName("PCH800-2");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ProfileConsys1.btn_Delete' at Center.", repo.ProfileConsys1.btn_DeleteInfo, new RecordItemIndex(52));
            repo.ProfileConsys1.btn_Delete.Click();
            Delay.Milliseconds(200);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Site");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.ShoppingList' at Center.", repo.FormMe.ShoppingListInfo, new RecordItemIndex(54));
            repo.FormMe.ShoppingList.Click();
            Delay.Milliseconds(200);
            
            Libraries.Devices_Functions.verifyShoppingList(ValueConverter.ArgumentFromString<int>("ShoppingListDeviceCount", "4"));
            Delay.Milliseconds(0);
            
            //Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.Export' at Center.", repo.FormMe.ExportInfo, new RecordItemIndex(56));
            //repo.FormMe.Export.Click();
            //Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.Export2ndTime' at Center.", repo.FormMe.Export2ndTimeInfo, new RecordItemIndex(57));
            repo.FormMe.Export2ndTime.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'PrintPreview.PARTMaximize' at Center.", repo.PrintPreview.PARTMaximizeInfo, new RecordItemIndex(58));
            repo.PrintPreview.PARTMaximize.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'PrintPreview.ExportDropdown' at Center.", repo.PrintPreview.ExportDropdownInfo, new RecordItemIndex(59));
            repo.PrintPreview.ExportDropdown.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ExportDocument.ExcelFormat' at Center.", repo.ExportDocument.ExcelFormatInfo, new RecordItemIndex(60));
            repo.ExportDocument.ExcelFormat.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Delay", "Waiting for 5s.", new RecordItemIndex(61));
            Delay.Duration(5000, false);
            
            Report.Log(ReportLevel.Info, "Set value", "Setting attribute Text to 'Xls' on item 'ExportDocument.ExcelFormat'.", repo.ExportDocument.ExcelFormatInfo, new RecordItemIndex(62));
            repo.ExportDocument.ExcelFormat.Element.SetAttributeValue("Text", "Xls");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Delay", "Waiting for 5s.", new RecordItemIndex(63));
            Delay.Duration(5000, false);
            
            Libraries.Export_Functions.validateAndClickOKButtonOnExportDocument();
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ExportDocument.ButtonOK' at Center.", repo.ExportDocument.ButtonOKInfo, new RecordItemIndex(65));
            repo.ExportDocument.ButtonOK.Click();
            Delay.Milliseconds(200);
            
            Libraries.Export_Functions.VerifyOverwriteMessageAndClickButton("The output file already exists. Click OK to overwrite.");
            Delay.Milliseconds(0);
            
            //Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ExportDocument.ButtonOK' at Center.", repo.ExportDocument.ButtonOKInfo, new RecordItemIndex(67));
            //repo.ExportDocument.ButtonOK.Click();
            //Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ShoppingListCompatibilityModeE.Cell18' at Center.", repo.ShoppingListCompatibilityModeE.Cell18Info, new RecordItemIndex(68));
            repo.ShoppingListCompatibilityModeE.Cell18.Click();
            Delay.Milliseconds(200);
            
            Libraries.Devices_Functions.verifyShoppingListDevicesTextForPxD("");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ShoppingListCompatibilityModeE.CellF22' at Center.", repo.ShoppingListCompatibilityModeE.CellF22Info, new RecordItemIndex(70));
            repo.ShoppingListCompatibilityModeE.CellF22.Click();
            Delay.Milliseconds(200);
            
            Libraries.Devices_Functions.verifyShoppingListDevicesTextForPSC("");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ShoppingListCompatibilityModeE.btn_CloseExcel' at Center.", repo.ShoppingListCompatibilityModeE.btn_CloseExcelInfo, new RecordItemIndex(72));
            repo.ShoppingListCompatibilityModeE.btn_CloseExcel.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'PrintPreview.btn_Close' at Center.", repo.PrintPreview.btn_CloseInfo, new RecordItemIndex(73));
            repo.PrintPreview.btn_Close.Click();
            Delay.Milliseconds(200);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
