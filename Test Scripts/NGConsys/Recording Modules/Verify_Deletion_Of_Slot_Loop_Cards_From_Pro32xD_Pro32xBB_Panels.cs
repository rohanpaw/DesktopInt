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
    ///The Verify_Deletion_Of_Slot_Loop_Cards_From_Pro32xD_Pro32xBB_Panels recording.
    /// </summary>
    [TestModule("f2137f57-f0f8-4489-bd0c-01380e88918a", ModuleType.Recording, 1)]
    public partial class Verify_Deletion_Of_Slot_Loop_Cards_From_Pro32xD_Pro32xBB_Panels : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Verify_Deletion_Of_Slot_Loop_Cards_From_Pro32xD_Pro32xBB_Panels instance = new Verify_Deletion_Of_Slot_Loop_Cards_From_Pro32xD_Pro32xBB_Panels();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_Deletion_Of_Slot_Loop_Cards_From_Pro32xD_Pro32xBB_Panels()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_Deletion_Of_Slot_Loop_Cards_From_Pro32xD_Pro32xBB_Panels Instance
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
            Keyboard.DefaultKeyPressTime = 100;
            Delay.SpeedFactor = 1.00;

            Init();

            Libraries.Panel_Functions.AddPanels(ValueConverter.ArgumentFromString<int>("NumberofPanels", "1"), "Pro32xBB", "");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeExpander("Node");
            Delay.Milliseconds(0);
            
            //Libraries.Common_Functions.ClickOnNavigationTreeExpander("Backplane");
            //Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Other Slot Cards");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromMainProcessorGallery("PCH800 5.0A", "Slot Cards", "PFI");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromMainProcessorGallery("PCH800 5.0A", "Slot Cards", "PFI");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Other Slot Cards");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromMainProcessorGallery("PCH800 5.0A", "Slot Cards", "PFI");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Other Slot Cards");
            Delay.Milliseconds(0);
            
            //Libraries.Common_Functions.ClickOnNavigationTreeExpander("Node");
            //Delay.Milliseconds(0);
            
            //Libraries.Devices_Functions.VerifyDeleteButton(ValueConverter.ArgumentFromString<bool>("isReadOnly", "False"));
            //Delay.Milliseconds(0);
            
            try {
                //Libraries.Common_Functions.verifyDeleteButtonState(ValueConverter.ArgumentFromString<bool>("isReadOnly", "False"));
                //Delay.Milliseconds(0);
            } catch(Exception ex) { Report.Log(ReportLevel.Warn, "Module", "(Optional Action) " + ex.Message, new RecordItemIndex(11)); }
            
            //Libraries.Common_Functions.ClickOnNavigationTreeExpander("Node");
            //Delay.Milliseconds(0);
            
            //Libraries.Common_Functions.VerifyNavigationTreeItemText("Node1");
            //Delay.Milliseconds(0);
            
            //Libraries.Devices_Functions.VerifyDeleteButton(ValueConverter.ArgumentFromString<bool>("isReadOnly", "False"));
            //Delay.Milliseconds(0);
            
            //Libraries.Common_Functions.verifyDeleteButtonState(ValueConverter.ArgumentFromString<bool>("isReadOnly", "False"));
            //Delay.Milliseconds(0);
            
            //Libraries.Common_Functions.ClickOnNavigationTreeExpander("Backplane");
            //Delay.Milliseconds(0);
            
            //Libraries.Common_Functions.ClickOnNavigationTreeExpander("Backplane");
            //Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Other Slot Cards");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelNameFromInventoryTab("PCH800 5.0A-3");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.verifyDeleteButtonState(ValueConverter.ArgumentFromString<bool>("isReadOnly", "True"));
            Delay.Milliseconds(0);
            
            //Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ProfileConsys1.btn_Delete' at Center.", repo.ProfileConsys1.btn_DeleteInfo, new RecordItemIndex(21));
            //repo.ProfileConsys1.btn_Delete.Click();
            //Delay.Milliseconds(200);
            
            //Libraries.Common_Functions.verifyButtonState(ValueConverter.ArgumentFromString<bool>("isReadOnly", "True"), "Delete");
            //Delay.Milliseconds(0);
            
            Libraries.Common_Functions.clickOnDeleteButton();
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Site");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.clickOnShoppingListTab();
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyShoppingList(ValueConverter.ArgumentFromString<int>("ShoppingListDeviceCount", "3"));
            Delay.Milliseconds(0);
            
            Libraries.Export_Functions.SearchDeviceInExportUsingSKUOrDescription("PxD", ValueConverter.ArgumentFromString<bool>("sExist", "False"));
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
