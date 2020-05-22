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
    ///The Verify_AC_Units_Calculation_For_Loops_With_IS_Devices recording.
    /// </summary>
    [TestModule("55ec67c7-3815-4ead-bddd-dedeba62e4cd", ModuleType.Recording, 1)]
    public partial class Verify_AC_Units_Calculation_For_Loops_With_IS_Devices : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Verify_AC_Units_Calculation_For_Loops_With_IS_Devices instance = new Verify_AC_Units_Calculation_For_Loops_With_IS_Devices();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_AC_Units_Calculation_For_Loops_With_IS_Devices()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_AC_Units_Calculation_For_Loops_With_IS_Devices Instance
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

            Libraries.Panel_Functions.AddPanels(ValueConverter.ArgumentFromString<int>("NumberofPanels", "1"), "MX 4000", "");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeExpander("Node");
            Delay.Milliseconds(0);
            
            //Libraries.Common_Functions.ClickOnNavigationTreeExpander("FIM");
            //Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.EnableISDevices();
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("Exi800", "Other");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.clickOnPanelCalculationsTab();
            Delay.Milliseconds(0);
            
            //Libraries.AC_Functions.verifyMaxACUnitsValue("20.39999");
            //Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyMaxLoadingDetailsValue("20.39999", "Signal (AC Units)");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("801 CHEx IS", "Detectors");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelName("Exi800");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("801 HEx IS", "Detectors");
            Delay.Milliseconds(0);
            
            //Libraries.AC_Functions.verifyMaxACUnitsValue("3.6");
            //Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyMaxLoadingDetailsValue("3.6", "Signal (AC Units)");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.clickOnPropertiesTab();
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.ChangeCableLength(ValueConverter.ArgumentFromString<int>("fchangeCableLength", "500"));
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
            Delay.Milliseconds(0);
            
            //Libraries.AC_Functions.verifyMaxACUnitsValue("53.6");
            //Delay.Milliseconds(0);
            
            Libraries.Common_Functions.clickOnPanelCalculationsTab();
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyMaxLoadingDetailsValue("53.6", "Signal (AC Units)");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.clickOnPropertiesTab();
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.ChangeCableCapacitance(ValueConverter.ArgumentFromString<int>("fchangeCableCapacitance", "90"), "Exi800");
            Delay.Milliseconds(0);
            
            //Libraries.Common_Functions.ClickOnNavigationTreeExpander("Node");
            //Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
            Delay.Milliseconds(0);
            
            //Libraries.AC_Functions.verifyMaxACUnitsValue("104");
            //Delay.Milliseconds(0);
            
            Libraries.Common_Functions.clickOnPanelCalculationsTab();
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyMaxLoadingDetailsValue("104", "Signal (AC Units)");
            Delay.Milliseconds(0);
            
            Libraries.AC_Functions.VerifyACUnitIndicationWithISDevices();
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromMultiplePointWizard("DDM 800 Loop", ValueConverter.ArgumentFromString<int>("DeviceQty", "28"));
            Delay.Milliseconds(0);
            
            Libraries.AC_Functions.VerifyACUnitIndicationWithISDevices();
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromMultiplePointWizard("DDM 800 Loop", ValueConverter.ArgumentFromString<int>("DeviceQty", "2"));
            Delay.Milliseconds(0);
            
            Libraries.AC_Functions.VerifyACUnitIndicationWithISDevices();
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
