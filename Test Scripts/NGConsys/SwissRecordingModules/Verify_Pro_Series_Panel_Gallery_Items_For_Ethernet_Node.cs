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

namespace TestProject.SwissRecordingModules
{
#pragma warning disable 0436 //(CS0436) The type 'type' in 'assembly' conflicts with the imported type 'type2' in 'assembly'. Using the type defined in 'assembly'.
    /// <summary>
    ///The Verify_Pro_Series_Panel_Gallery_Items_For_Ethernet_Node recording.
    /// </summary>
    [TestModule("044d954c-42e8-4fe6-bd98-a61e181102f6", ModuleType.Recording, 1)]
    public partial class Verify_Pro_Series_Panel_Gallery_Items_For_Ethernet_Node : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Verify_Pro_Series_Panel_Gallery_Items_For_Ethernet_Node instance = new Verify_Pro_Series_Panel_Gallery_Items_For_Ethernet_Node();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_Pro_Series_Panel_Gallery_Items_For_Ethernet_Node()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_Pro_Series_Panel_Gallery_Items_For_Ethernet_Node Instance
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

            Libraries.Panel_Functions.AddPanels(ValueConverter.ArgumentFromString<int>("NumberofPanels", "1"), "Pro815D", "");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeExpander("Node");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Ethernet");
            Delay.Milliseconds(0);
            
            // Repeaters
            Libraries.Gallery_Functions.verifyGalleryListItems("Gallery_Ethernet_ProPanels", "Swiss", "PR8AS");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("PR1DS", "");
            Delay.Milliseconds(0);
            
            // Attached Functionality
            Libraries.Gallery_Functions.verifyGalleryListItems("Gallery_Ethernet_AttachedFunctionality_ProPanels", "Swiss", "PZ8DS");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelNameFromInventoryTab("PR1DS-102");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingModelNameFromInventoryTab("PR1DS");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingSKUFromInventoryTab("557.200.801");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("PR1DS", "Repeaters");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("PR1DSCH\r\n", "Repeaters");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("PR8AS", "Repeaters");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("PR8ASCH", "Repeaters");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("PR1DS", "Repeaters");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("PR1DS", "Repeaters");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.VerifyNavigationTreeItemText("Ethernet (8 of 8)");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelNameFromInventoryTab("PR1DS-108");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("Two PZ4DS", "Attached Functionality");
            Delay.Milliseconds(0);
            
            Libraries.Gallery_Functions.verifyGalleryExistsWithDropdown(ValueConverter.ArgumentFromString<bool>("GalleryVisibility", "False"), "PR1DS");
            Delay.Milliseconds(0);
            
            Libraries.Gallery_Functions.verifyGalleryExistsWithDropdown(ValueConverter.ArgumentFromString<bool>("GalleryVisibility", "False"), "Two PZ4DS");
            Delay.Milliseconds(0);
            
            Libraries.Panel_Functions.DeletePanel(ValueConverter.ArgumentFromString<int>("NumberofPanels", "1"), "Node1", ValueConverter.ArgumentFromString<int>("rowNumber", "1"));
            Delay.Milliseconds(0);
            
            // PRO LITE PANEL TEST CASE
            Report.Log(ReportLevel.Info, "Section", "PRO LITE PANEL TEST CASE", new RecordItemIndex(21));
            
            Libraries.Panel_Functions.AddPanelsMT(ValueConverter.ArgumentFromString<int>("NumberofPanels", "1"), "Pro415D Lite", "");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeExpander("Node");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Ethernet");
            Delay.Milliseconds(0);
            
            Libraries.Gallery_Functions.verifyGalleryListItems("Gallery_Ethernet_ProPanels", "Swiss", "PR8AS");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("PR1DS", "");
            Delay.Milliseconds(0);
            
            Libraries.Gallery_Functions.verifyGalleryListItems("Gallery_Ethernet_AttachedFunctionality_ProPanels", "Swiss", "PZ8DS");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelNameFromInventoryTab("PR1DS-102");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingModelNameFromInventoryTab("PR1DS");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingSKUFromInventoryTab("557.200.801");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("PR1DS", "Repeaters");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("PR1DS", "Repeaters");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("PR8AS", "Repeaters");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("PR1DS", "Repeaters");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("PR1DS", "Repeaters");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("PR1DS", "Repeaters");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.VerifyNavigationTreeItemText("Ethernet (8 of 8)");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelNameFromInventoryTab("PR1DS-108");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("Two PZ4DS", "Attached Functionality");
            Delay.Milliseconds(0);
            
            Libraries.Gallery_Functions.verifyGalleryExistsWithDropdown(ValueConverter.ArgumentFromString<bool>("GalleryVisibility", "False"), "PR1DS");
            Delay.Milliseconds(0);
            
            Libraries.Gallery_Functions.verifyGalleryExistsWithDropdown(ValueConverter.ArgumentFromString<bool>("GalleryVisibility", "False"), "Two PZ4DS");
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
