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
    ///The Verify_the_properties_of_Audix_Interface recording.
    /// </summary>
    [TestModule("e5591827-cfa7-4c05-9578-6ea1d74f4404", ModuleType.Recording, 1)]
    public partial class Verify_the_properties_of_Audix_Interface : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Verify_the_properties_of_Audix_Interface instance = new Verify_the_properties_of_Audix_Interface();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_the_properties_of_Audix_Interface()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_the_properties_of_Audix_Interface Instance
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
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Main");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromMainProcessorGallery("MPM800", "Miscellaneous", "PFI");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.tab_Inventory' at Center.", repo.FormMe.tab_InventoryInfo, new RecordItemIndex(4));
            repo.FormMe.tab_Inventory.Click();
            Delay.Milliseconds(200);
            
            Libraries.InventoryGrid_Functions.SelectRowUsingDevicePropertyForMainProcessorGallery("2", "Label", "MPM800-1");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromMainProcessorGallery("Audix 8", "Attached Functionality", "PFI");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.tab_Inventory' at Center.", repo.FormMe.tab_InventoryInfo, new RecordItemIndex(7));
            repo.FormMe.tab_Inventory.Click();
            Delay.Milliseconds(200);
            
            Libraries.InventoryGrid_Functions.SelectRowUsingDevicePropertyForMainProcessorGallery("3", "Label", "Audix 8");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.VerifyModelInSearchProperties("XIOM");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.VerifyProductCodeTextRowInSearchProperties();
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyFunctionality("XIOM configured to communicate with an 8 zone Audix Voice Alarm System");
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
