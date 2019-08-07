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
    ///The Verify_Enabling_Or_Disabling_Of_Paste_Button_For_Copy_Paste_Functionality_For_Audix recording.
    /// </summary>
    [TestModule("973b545f-daf4-47a9-b82e-d0a7b6dbfca9", ModuleType.Recording, 1)]
    public partial class Verify_Enabling_Or_Disabling_Of_Paste_Button_For_Copy_Paste_Functionality_For_Audix : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Verify_Enabling_Or_Disabling_Of_Paste_Button_For_Copy_Paste_Functionality_For_Audix instance = new Verify_Enabling_Or_Disabling_Of_Paste_Button_For_Copy_Paste_Functionality_For_Audix();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_Enabling_Or_Disabling_Of_Paste_Button_For_Copy_Paste_Functionality_For_Audix()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_Enabling_Or_Disabling_Of_Paste_Button_For_Copy_Paste_Functionality_For_Audix Instance
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
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.NodeExpander1' at Center.", repo.FormMe.NodeExpander1Info, new RecordItemIndex(1));
            repo.FormMe.NodeExpander1.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left DoubleClick item 'FormMe.MainProcessorNode1' at Center.", repo.FormMe.MainProcessorNode1Info, new RecordItemIndex(2));
            repo.FormMe.MainProcessorNode1.DoubleClick();
            Delay.Milliseconds(200);
            
            Libraries.Devices_Functions.AddDevicesfromMainProcessorGallery("MPM800", "Miscellaneous", "PFI");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.tab_Inventory' at Center.", repo.FormMe.tab_InventoryInfo, new RecordItemIndex(4));
            repo.FormMe.tab_Inventory.Click();
            Delay.Milliseconds(200);
            
            Libraries.InventoryGrid_Functions.SelectRowUsingDevicePropertyForMainProcessorGallery("2", "Label", "MPM800-1");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromMainProcessorGallery("Audix 8", "Attached Functionality", "PFI");
            Delay.Milliseconds(0);
            
            Libraries.InventoryGrid_Functions.SelectRowUsingDevicePropertyForMainProcessorGallery("3", "Label", "Audix 8");
            Delay.Milliseconds(0);
            
            Libraries.Gallery_Functions.VerifyCopyButton(ValueConverter.ArgumentFromString<bool>("isEnabled", "False"));
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyPasteButtonDisabled();
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.LoopExpander1' at Center.", repo.FormMe.LoopExpander1Info, new RecordItemIndex(10));
            repo.FormMe.LoopExpander1.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.Loop_A1' at Center.", repo.FormMe.Loop_A1Info, new RecordItemIndex(11));
            repo.FormMe.Loop_A1.Click();
            Delay.Milliseconds(200);
            
            Libraries.Devices_Functions.SelectRowUsingLabelName("SIO 800 - 1");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.btn_Copy1' at Center.", repo.FormMe.btn_Copy1Info, new RecordItemIndex(13));
            repo.FormMe.btn_Copy1.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.Paste' at Center.", repo.FormMe.PasteInfo, new RecordItemIndex(14));
            repo.FormMe.Paste.Click();
            Delay.Milliseconds(200);
            
            Libraries.Devices_Functions.VerifyDeviceExists(ValueConverter.ArgumentFromString<bool>("sExists", "True"), "SIO 800 - 3");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelName("RIM 800 - 2");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.btn_Copy1' at Center.", repo.FormMe.btn_Copy1Info, new RecordItemIndex(17));
            repo.FormMe.btn_Copy1.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.Paste' at Center.", repo.FormMe.PasteInfo, new RecordItemIndex(18));
            repo.FormMe.Paste.Click();
            Delay.Milliseconds(200);
            
            Libraries.Devices_Functions.VerifyDeviceExists(ValueConverter.ArgumentFromString<bool>("sExists", "True"), "RIM 800 - 4");
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
