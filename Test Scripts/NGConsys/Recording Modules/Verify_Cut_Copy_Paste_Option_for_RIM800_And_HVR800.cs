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
    ///The Verify_Cut_Copy_Paste_Option_for_RIM800_And_HVR800 recording.
    /// </summary>
    [TestModule("d06aa5da-a769-490f-ac17-1b54645a0b7b", ModuleType.Recording, 1)]
    public partial class Verify_Cut_Copy_Paste_Option_for_RIM800_And_HVR800 : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Verify_Cut_Copy_Paste_Option_for_RIM800_And_HVR800 instance = new Verify_Cut_Copy_Paste_Option_for_RIM800_And_HVR800();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_Cut_Copy_Paste_Option_for_RIM800_And_HVR800()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_Cut_Copy_Paste_Option_for_RIM800_And_HVR800 Instance
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

            Libraries.Panel_Functions.AddPanels(ValueConverter.ArgumentFromString<int>("NumberofPanels", "1"), "Pro32xD", "");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeExpander("Node");
            Delay.Milliseconds(0);
            
            //Libraries.Common_Functions.ClickOnNavigationTreeExpander("PFI");
            //Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("RIM 800", "Ancillary");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelNameForOneRow("RIM 800 - 1");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("HVR800", "Other");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelName("RIM 800 - 1");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.clickOnCopyButton();
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.clickOnPasteButton();
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelName("RIM 800 - 2");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelName("HVR800 - 1");
            Delay.Milliseconds(0);
            
            //Libraries.Gallery_Functions.VerifyCopyButton(ValueConverter.ArgumentFromString<bool>("isEnabled", "False"));
            //Delay.Milliseconds(0);
            
            try {
                //Libraries.Gallery_Functions.VerifyCopyButtonState(ValueConverter.ArgumentFromString<bool>("isEnabled", "False"));
                //Delay.Milliseconds(0);
            } catch(Exception ex) { Report.Log(ReportLevel.Warn, "Module", "(Optional Action) " + ex.Message, new RecordItemIndex(13)); }
            
            Libraries.Common_Functions.SaveProject("54241");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.Application_Close(ValueConverter.ArgumentFromString<bool>("Save", "False"), ValueConverter.ArgumentFromString<bool>("SaveConfirmation", "False"), "");
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
