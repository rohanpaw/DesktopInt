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
    ///The Verify_DC_Unit_For_Built_In_Isolator_Devices_With_Non_Integral_LED_On_FIM_Loop_Considering_Trip_Curr recording.
    /// </summary>
    [TestModule("8d96ec4c-823d-433c-b189-1446cab232e4", ModuleType.Recording, 1)]
    public partial class Verify_DC_Unit_For_Built_In_Isolator_Devices_With_Non_Integral_LED_On_FIM_Loop_Considering_Trip_Curr : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Verify_DC_Unit_For_Built_In_Isolator_Devices_With_Non_Integral_LED_On_FIM_Loop_Considering_Trip_Curr instance = new Verify_DC_Unit_For_Built_In_Isolator_Devices_With_Non_Integral_LED_On_FIM_Loop_Considering_Trip_Curr();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_DC_Unit_For_Built_In_Isolator_Devices_With_Non_Integral_LED_On_FIM_Loop_Considering_Trip_Curr()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_DC_Unit_For_Built_In_Isolator_Devices_With_Non_Integral_LED_On_FIM_Loop_Considering_Trip_Curr Instance
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
            
            Libraries.Devices_Functions.AddDevicesfromGallery("LPBS 3000", "Sounders/Beacons");
            Delay.Milliseconds(0);
            
            //Libraries.Devices_Functions.verifyIsolatorCheckbox(ValueConverter.ArgumentFromString<bool>("ExpectedState", "True"));
            //Delay.Milliseconds(0);
            
            //Libraries.DC_Functions.verifyDCUnitsValue("287");
            //Delay.Milliseconds(0);
            
            //Libraries.Devices_Functions.verifyLoadingDetailsValue("287", "Current (DC Units)");
            //Delay.Milliseconds(0);
            
            Libraries.Common_Functions.clickOnPanelCalculationsTab();
            Delay.Milliseconds(0);
            
            // Current (DC Units)
            Libraries.Devices_Functions.verifyLoopLoadingDetailsValue("278", "Built-in Loop-A", "2");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.clickOnPropertiesTab();
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("P80SB", "Sounders/Beacons");
            Delay.Milliseconds(0);
            
            //Libraries.Devices_Functions.verifyIsolatorCheckbox(ValueConverter.ArgumentFromString<bool>("ExpectedState", "True"));
            //Delay.Milliseconds(0);
            
            //Libraries.DC_Functions.verifyDCUnitsValue("311");
            //Delay.Milliseconds(0);
            
            //Libraries.Devices_Functions.verifyLoadingDetailsValue("347", "Current (DC Units)");
            //Delay.Milliseconds(0);
            
            Libraries.Common_Functions.clickOnPanelCalculationsTab();
            Delay.Milliseconds(0);
            
            // Current (DC Units)
            Libraries.Devices_Functions.verifyLoopLoadingDetailsValue("305", "Built-in Loop-A", "2");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.clickOnPropertiesTab();
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
