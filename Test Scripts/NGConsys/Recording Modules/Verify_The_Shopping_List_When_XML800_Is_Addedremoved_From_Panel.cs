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
    ///The Verify_The_Shopping_List_When_XML800_Is_Addedremoved_From_Panel recording.
    /// </summary>
    [TestModule("7259ae3a-2387-405c-a951-459eb5e20671", ModuleType.Recording, 1)]
    public partial class Verify_The_Shopping_List_When_XML800_Is_Addedremoved_From_Panel : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Verify_The_Shopping_List_When_XML800_Is_Addedremoved_From_Panel instance = new Verify_The_Shopping_List_When_XML800_Is_Addedremoved_From_Panel();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_The_Shopping_List_When_XML800_Is_Addedremoved_From_Panel()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_The_Shopping_List_When_XML800_Is_Addedremoved_From_Panel Instance
        {
            get { return instance; }
        }

#region Variables

#endregion

        /// <summary>
        /// Starts the replay of the static recording <see cref="Instance"/>.
        /// </summary>
        [System.CodeDom.Compiler.GeneratedCode("Ranorex", global::Ranorex.Core.Constants.CodeGenVersion)]
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
        [System.CodeDom.Compiler.GeneratedCode("Ranorex", global::Ranorex.Core.Constants.CodeGenVersion)]
        void ITestModule.Run()
        {
            Mouse.DefaultMoveTime = 300;
            Keyboard.DefaultKeyPressTime = 20;
            Delay.SpeedFactor = 1.00;

            Init();

            Libraries.Panel_Functions.AddPanels(ValueConverter.ArgumentFromString<int>("NumberofPanels", "1"), "MZX252", "");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Node");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromPanelNodeGallery("XLM800", "Loops", "PFI");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Node");
            Delay.Milliseconds(0);
            
            Libraries.Panel_Functions.VerifyValueOf2ndPSU("PMM840");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Node");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Site");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.clickOnShoppingListTab();
            Delay.Milliseconds(0);
            
            //Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.tab_ShoppingList' at Center.", repo.FormMe.tab_ShoppingListInfo, new RecordItemIndex(8));
            //repo.FormMe.tab_ShoppingList.Click();
            //Delay.Milliseconds(200);
            
            Libraries.Export_Functions.SearchDeviceInExportUsingSKUOrDescription("557.202.073", ValueConverter.ArgumentFromString<bool>("sExist", "True"));
            Delay.Milliseconds(0);
            
            Libraries.Export_Functions.SearchDeviceInExportUsingSKUOrDescription("557.202.612", ValueConverter.ArgumentFromString<bool>("sExist", "True"));
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Node");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.clickOnInventoryTab();
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelNameFromInventoryTab("XLM800-C");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.clickOnDeleteButton();
            Delay.Milliseconds(0);
            
            //Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ProfileConsys1.btn_Delete' at Center.", repo.ProfileConsys1.btn_DeleteInfo, new RecordItemIndex(15));
            //repo.ProfileConsys1.btn_Delete.Click();
            //Delay.Milliseconds(200);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Node");
            Delay.Milliseconds(0);
            
            //Libraries.Panel_Functions.VerifyValueOf2ndPSU("None");
            //Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Site");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.clickOnShoppingListTab();
            Delay.Milliseconds(0);
            
            //Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.tab_ShoppingList' at Center.", repo.FormMe.tab_ShoppingListInfo, new RecordItemIndex(20));
            //repo.FormMe.tab_ShoppingList.Click();
            //Delay.Milliseconds(200);
            
            Libraries.Export_Functions.SearchDeviceInExportUsingSKUOrDescription("557.202.073", ValueConverter.ArgumentFromString<bool>("sExist", "False"));
            Delay.Milliseconds(0);
            
            Libraries.Export_Functions.SearchDeviceInExportUsingSKUOrDescription("557.202.612", ValueConverter.ArgumentFromString<bool>("sExist", "False"));
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
