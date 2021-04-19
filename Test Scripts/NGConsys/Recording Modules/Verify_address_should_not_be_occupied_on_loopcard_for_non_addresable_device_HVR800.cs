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
    ///The Verify_address_should_not_be_occupied_on_loopcard_for_non_addresable_device_HVR800 recording.
    /// </summary>
    [TestModule("496de603-0d69-4fc6-b09f-6903beac02eb", ModuleType.Recording, 1)]
    public partial class Verify_address_should_not_be_occupied_on_loopcard_for_non_addresable_device_HVR800 : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Verify_address_should_not_be_occupied_on_loopcard_for_non_addresable_device_HVR800 instance = new Verify_address_should_not_be_occupied_on_loopcard_for_non_addresable_device_HVR800();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_address_should_not_be_occupied_on_loopcard_for_non_addresable_device_HVR800()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_address_should_not_be_occupied_on_loopcard_for_non_addresable_device_HVR800 Instance
        {
            get { return instance; }
        }

#region Variables

        /// <summary>
        /// Gets or sets the value of variable sRow.
        /// </summary>
        [TestVariable("a610866c-7085-4bec-8c50-ccf9db7fc18b")]
        public string sRow
        {
            get { return repo.sRow; }
            set { repo.sRow = value; }
        }

        /// <summary>
        /// Gets or sets the value of variable sColumn.
        /// </summary>
        [TestVariable("5b381f08-b8c2-4ac0-bd6e-0ea0a8bde220")]
        public string sColumn
        {
            get { return repo.sColumn; }
            set { repo.sColumn = value; }
        }

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
            
            Libraries.Devices_Functions.VerifyGalleryItem("Other", "HVR800", "Enabled");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelNameForOneRow("RIM 800 - 1");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("HVR800", "Other");
            Delay.Milliseconds(0);
            
            //Libraries.Devices_Functions.verifyBlankDeviceAddress("2", "Address");
            //Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromMultiplePointWizard("RIM 800", ValueConverter.ArgumentFromString<int>("DeviceQty", "124"));
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Delay", "Waiting for 5s.", new RecordItemIndex(11));
            Delay.Duration(5000, false);
            
            Libraries.Common_Functions.clickOnPointsTab();
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelName("RIM 800 - 2");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("HVR800", "Other");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.SaveProject("54250");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.Application_Close(ValueConverter.ArgumentFromString<bool>("Save", "False"), ValueConverter.ArgumentFromString<bool>("SaveConfirmation", "False"), "");
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
