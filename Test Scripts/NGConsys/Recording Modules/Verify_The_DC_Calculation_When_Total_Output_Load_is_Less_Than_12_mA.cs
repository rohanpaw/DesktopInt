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
    ///The Verify_The_DC_Calculation_When_Total_Output_Load_is_Less_Than_12_mA recording.
    /// </summary>
    [TestModule("a196fa3b-49b1-4b41-81b3-621b9d5c6a6f", ModuleType.Recording, 1)]
    public partial class Verify_The_DC_Calculation_When_Total_Output_Load_is_Less_Than_12_mA : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Verify_The_DC_Calculation_When_Total_Output_Load_is_Less_Than_12_mA instance = new Verify_The_DC_Calculation_When_Total_Output_Load_is_Less_Than_12_mA();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_The_DC_Calculation_When_Total_Output_Load_is_Less_Than_12_mA()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_The_DC_Calculation_When_Total_Output_Load_is_Less_Than_12_mA Instance
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

            Libraries.Panel_Functions.AddPanels(ValueConverter.ArgumentFromString<int>("NumberofPanels", "1"), "MX4000", "");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeExpander("Node");
            Delay.Milliseconds(0);
            
            //Libraries.Common_Functions.ClickOnNavigationTreeExpander("PFI");
            //Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("LPS 800", "Sounders/Beacons");
            Delay.Milliseconds(0);
            
            //Libraries.DC_Functions.verifyDCUnitsValue("280");
            //Delay.Milliseconds(0);
            
            //Libraries.Devices_Functions.verifyLoadingDetailsValue("280", "Current (DC Units)");
            //Delay.Milliseconds(0);
            
            Libraries.Common_Functions.clickOnPanelCalculationsTab();
            Delay.Milliseconds(0);
            
            // Current (DC Units)
            Libraries.Devices_Functions.verifyLoopLoadingDetailsValue("280", "Built-in Loop-A", "2");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.clickOnPropertiesTab();
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Site");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesFromGalleryUsingIndex("Generic Sounder", "Attachable Devices");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelName("Generic Sounder - 1");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.changeAndVerifyAlarmLoad(ValueConverter.ArgumentFromString<int>("AlarmLoad", "3"), "Valid", ValueConverter.ArgumentFromString<int>("expectedResult", "3"));
            Delay.Milliseconds(0);
            
            //Libraries.DC_Functions.verifyDCUnitsValue("280");
            //Delay.Milliseconds(0);
            
            //Libraries.Devices_Functions.verifyLoadingDetailsValue("280", "Current (DC Units)");
            //Delay.Milliseconds(0);
            
            Libraries.Common_Functions.clickOnPanelCalculationsTab();
            Delay.Milliseconds(0);
            
            // Current (DC Units)
            Libraries.Devices_Functions.verifyLoopLoadingDetailsValue("280", "Built-in Loop-A", "2");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.clickOnPropertiesTab();
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Site");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelName("Generic Sounder - 1");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.changeAndVerifyAlarmLoad(ValueConverter.ArgumentFromString<int>("AlarmLoad", "25"), "Valid", ValueConverter.ArgumentFromString<int>("expectedResult", "25"));
            Delay.Milliseconds(0);
            
            //Libraries.DC_Functions.verifyDCUnitsValue("348");
            //Delay.Milliseconds(0);
            
            //Libraries.Devices_Functions.verifyLoadingDetailsValue("348", "Current (DC Units)");
            //Delay.Milliseconds(0);
            
            Libraries.Common_Functions.clickOnPanelCalculationsTab();
            Delay.Milliseconds(0);
            
            // Current (DC Units)
            Libraries.Devices_Functions.verifyLoopLoadingDetailsValue("332", "Built-in Loop-A", "2");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.clickOnPropertiesTab();
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
