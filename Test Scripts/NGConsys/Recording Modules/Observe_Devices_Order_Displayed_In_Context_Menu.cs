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
    ///The Observe_Devices_Order_Displayed_In_Context_Menu recording.
    /// </summary>
    [TestModule("660de8e4-d434-4c63-85b7-ce3df90b906c", ModuleType.Recording, 1)]
    public partial class Observe_Devices_Order_Displayed_In_Context_Menu : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Observe_Devices_Order_Displayed_In_Context_Menu instance = new Observe_Devices_Order_Displayed_In_Context_Menu();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Observe_Devices_Order_Displayed_In_Context_Menu()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Observe_Devices_Order_Displayed_In_Context_Menu Instance
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
            
            Libraries.Common_Functions.ClickOnNavigationTreeExpander("Node");
            Delay.Milliseconds(0);
            
            //Libraries.Common_Functions.ClickOnNavigationTreeExpander("PFI");
            //Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("801 CH", "Detectors");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("801 CH", "Detectors");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.RightClickOnSelectedRow("1");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyContextMenuOptionTextOnRightClickInPointsGrid("Detectors");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyContextMenuOptionTextOnRightClickInPointsGrid("Call points");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyContextMenuOptionTextOnRightClickInPointsGrid("Sounders/Beacons");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyContextMenuOptionTextOnRightClickInPointsGrid("Ancillary");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyContextMenuOptionTextOnRightClickInPointsGrid("Ancillary Conventional");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyContextMenuOptionTextOnRightClickInPointsGrid("Ancillary Specific");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyContextMenuOptionTextOnRightClickInPointsGrid("Other");
            Delay.Milliseconds(0);
            
            Libraries.Panel_Functions.DeletePanel(ValueConverter.ArgumentFromString<int>("NumberofPanels", "1"), "Node1", ValueConverter.ArgumentFromString<int>("rowNumber", "1"));
            Delay.Milliseconds(0);
            
            Libraries.Panel_Functions.AddPanels(ValueConverter.ArgumentFromString<int>("NumberofPanels", "1"), "MX 4000", "");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeExpander("Node");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeExpander("FIM");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("801 CH", "Detectors");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("801 CH", "Detectors");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.RightClickOnSelectedRow("1");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyContextMenuOptionTextOnRightClickInPointsGrid("Detectors");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyContextMenuOptionTextOnRightClickInPointsGrid("Call points");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyContextMenuOptionTextOnRightClickInPointsGrid("Sounders/Beacons");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyContextMenuOptionTextOnRightClickInPointsGrid("Ancillary");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyContextMenuOptionTextOnRightClickInPointsGrid("Ancillary Conventional");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyContextMenuOptionTextOnRightClickInPointsGrid("Ancillary Specific");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyContextMenuOptionTextOnRightClickInPointsGrid("Other");
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
