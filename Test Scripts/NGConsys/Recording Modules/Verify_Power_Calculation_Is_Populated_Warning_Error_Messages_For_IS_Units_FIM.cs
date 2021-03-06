///////////////////////////////////////////////////////////////////////////////
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
    ///The Verify_Power_Calculation_Is_Populated_Warning_Error_Messages_For_IS_Units_FIM recording.
    /// </summary>
    [TestModule("d3a6a84c-4624-49db-869b-8e558f7f5783", ModuleType.Recording, 1)]
    public partial class Verify_Power_Calculation_Is_Populated_Warning_Error_Messages_For_IS_Units_FIM : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Verify_Power_Calculation_Is_Populated_Warning_Error_Messages_For_IS_Units_FIM instance = new Verify_Power_Calculation_Is_Populated_Warning_Error_Messages_For_IS_Units_FIM();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_Power_Calculation_Is_Populated_Warning_Error_Messages_For_IS_Units_FIM()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_Power_Calculation_Is_Populated_Warning_Error_Messages_For_IS_Units_FIM Instance
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
            
            Libraries.Common_Functions.ClickOnNavigationTreeExpander("FIM");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
            Delay.Milliseconds(0);
            
            Libraries.PSULoad_Functions.verifyPowerCalculationsForISUnitsAndACUnits("TC_52210_11_12_13_Verify_Power_Calculation_Is_Populated_Where_Error_And_Warning_Messages_Are_Displayed_FIM", "IS_Units");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectPointsGridRow("1");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("S271i+\r\n", "Detectors");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("FV 411 F", "Detectors");
            Delay.Milliseconds(0);
            
            Libraries.PSULoad_Functions.verifyPowerCalculationsForISUnits("FIM");
            Delay.Milliseconds(0);
            
            Libraries.PSULoad_Functions.verifyPowerCalculationsText("AC value has reached 100% for panel Node1-MX 4000,Intrinsicly safe units value has reached 95% for panel Node1-MX 4000");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectPointsGridRow("1");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("801 CHEx IS", "Detectors");
            Delay.Milliseconds(0);
            
            Libraries.PSULoad_Functions.verifyPowerCalculationsForISUnits("FIM");
            Delay.Milliseconds(0);
            
            Libraries.PSULoad_Functions.verifyPowerCalculationsText("AC value has reached 100% for panel Node1-MX 4000,Intrinsicly safe units value has reached 100% for panel Node1-MX 4000");
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
