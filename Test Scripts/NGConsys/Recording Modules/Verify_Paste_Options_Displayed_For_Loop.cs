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
    ///The Verify_Paste_Options_Displayed_For_Loop recording.
    /// </summary>
    [TestModule("a2460558-217d-4451-affa-e0ba9edbdb7f", ModuleType.Recording, 1)]
    public partial class Verify_Paste_Options_Displayed_For_Loop : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Verify_Paste_Options_Displayed_For_Loop instance = new Verify_Paste_Options_Displayed_For_Loop();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_Paste_Options_Displayed_For_Loop()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_Paste_Options_Displayed_For_Loop Instance
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
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Node");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.clickOnPanelAccessoriesTab();
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicefromPanelAccessoriesGallery("IOB800", "Accessories");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeExpander("Node");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.clickOnInventoryTab();
            Delay.Milliseconds(0);
            
            Libraries.InventoryGrid_Functions.verifyInventoryGridProperties("7", "Label", "IOB800-1");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Node");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromMainProcessorGallery("PLX800", "Loops", "PFI");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.RightClickOnSelectedInventoryGridRow("6");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.clickContextMenuOptionOnRightClick("Copy");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyPasteButtonEnabled();
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.RightClickOnSelectedInventoryGridRow("6");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyContextMenuOptionOnRightClickEnabledOrDisabled("Paste");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeExpander("PFI");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyPasteButtonDisabled();
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("801 CH", "Detectors");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("801 CH", "Detectors");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.RightClickOnSelectedRow("1");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyContextMenuOptionOnRightClickEnabledOrDisabled("Paste");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.clickContextMenuOptionOnRightClick("Copy");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectPointsGridRow("1");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyPasteButtonEnabled();
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.RightClickOnSelectedRow("1");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyContextMenuOptionOnRightClickEnabledOrDisabled("Paste");
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
