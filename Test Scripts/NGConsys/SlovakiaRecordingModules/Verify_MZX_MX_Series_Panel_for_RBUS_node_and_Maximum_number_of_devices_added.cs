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

namespace TestProject.SlovakiaRecordingModules
{
#pragma warning disable 0436 //(CS0436) The type 'type' in 'assembly' conflicts with the imported type 'type2' in 'assembly'. Using the type defined in 'assembly'.
    /// <summary>
    ///The Verify_MZX_MX_Series_Panel_for_RBUS_node_and_Maximum_number_of_devices_added recording.
    /// </summary>
    [TestModule("a9d7744d-af38-4a1e-ba36-7f6c1fddf88a", ModuleType.Recording, 1)]
    public partial class Verify_MZX_MX_Series_Panel_for_RBUS_node_and_Maximum_number_of_devices_added : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Verify_MZX_MX_Series_Panel_for_RBUS_node_and_Maximum_number_of_devices_added instance = new Verify_MZX_MX_Series_Panel_for_RBUS_node_and_Maximum_number_of_devices_added();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_MZX_MX_Series_Panel_for_RBUS_node_and_Maximum_number_of_devices_added()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_MZX_MX_Series_Panel_for_RBUS_node_and_Maximum_number_of_devices_added Instance
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

            Libraries.Panel_Functions.AddPanels(ValueConverter.ArgumentFromString<int>("NumberofPanels", "1"), "MZX250", "");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeExpander("Node");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("RBus");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelNameFromInventoryTab("Front Door");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelNameFromInventoryTab("Zonal Alarm/Fault Display 32");
            Delay.Milliseconds(0);
            
            // Repeaters
            Libraries.Gallery_Functions.verifyGalleryListItems("Gallery_RBus_Repeaters_MZX_Panels", "Slovakia", "MXR");
            Delay.Milliseconds(0);
            
            // Miscellaneous
            Libraries.Gallery_Functions.verifyGalleryListItems("Gallery_Miscellaneous_MZX_Panels", "Slovakia", "PR1D2");
            Delay.Milliseconds(0);
            
            // Printers
            Libraries.Gallery_Functions.verifyGalleryListItems("Gallery_RBus_Printers_MZX_Panels", "Slovakia", "LCD800");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("MXR", "");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("MXR-F", "");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("MXR", "");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("MXR-F", "");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("MXR", "");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("MXR-F", "");
            Delay.Milliseconds(0);
            
            //Libraries.Devices_Functions.AddDevicesfromGallery("MXR-F", "");
            //Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("MPM800", "");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("MPM800", "");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("MPM800", "");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("MPM800", "");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("MPM800", "");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("MPM800", "");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("MPM800", "");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("MPM800", "");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("MPM800", "");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("XIOM", "");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.VerifyNavigationTreeItemText("RBus (16 of 16)");
            Delay.Milliseconds(0);
            
            //Libraries.Gallery_Functions.verifyGalleryExistsWithDropdown(ValueConverter.ArgumentFromString<bool>("GalleryVisibility", "False"), "MXR");
            //Delay.Milliseconds(0);
            
            //Libraries.Gallery_Functions.verifyGalleryExistsWithDropdown(ValueConverter.ArgumentFromString<bool>("GalleryVisibility", "False"), "MPM800");
            //Delay.Milliseconds(0);
            
            //Libraries.Gallery_Functions.verifyGalleryExistsWithDropdown(ValueConverter.ArgumentFromString<bool>("GalleryVisibility", "False"), "IOB800(x2)");
            //Delay.Milliseconds(0);
            
            Libraries.Panel_Functions.DeletePanel(ValueConverter.ArgumentFromString<int>("NumberofPanels", "1"), "Node1", ValueConverter.ArgumentFromString<int>("rowNumber", "1"));
            Delay.Milliseconds(0);
            
            Libraries.Panel_Functions.AddPanels(ValueConverter.ArgumentFromString<int>("NumberofPanels", "1"), "ZX1", "");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeExpander("Node");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("RBus");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyInventoryLabelNameExist(ValueConverter.ArgumentFromString<bool>("Visibility", "True"), "Front Door");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyInventoryLabelNameExist(ValueConverter.ArgumentFromString<bool>("Visibility", "False"), "Zonal Alarm/Fault Display 32");
            Delay.Milliseconds(0);
            
            // Repeaters
            Libraries.Gallery_Functions.verifyGalleryListItems("Gallery_RBus_Repeaters_MZX_Panels", "Slovakia", "MXR");
            Delay.Milliseconds(0);
            
            // Miscellaneous
            Libraries.Gallery_Functions.verifyGalleryListItems("Gallery_Miscellaneous_ZX_BB_Panels", "Slovakia", "MPM800");
            Delay.Milliseconds(0);
            
            // Printers
            Libraries.Gallery_Functions.verifyGalleryListItems("Gallery_RBus_Printers_MZX_Panels", "Slovakia", "LCD800");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("MXR", "");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("MXR-F", "");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("MXR", "");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("MXR-F", "");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("MXR", "");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("MXR-F", "");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("MPM800", "");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("MPM800", "");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("MPM800", "");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("MPM800", "");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("MPM800", "");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("MPM800", "");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("MPM800", "");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("MPM800", "");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("MPM800", "");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("COM820 I/O Module", "");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.VerifyNavigationTreeItemText("RBus (16 of 16)");
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
