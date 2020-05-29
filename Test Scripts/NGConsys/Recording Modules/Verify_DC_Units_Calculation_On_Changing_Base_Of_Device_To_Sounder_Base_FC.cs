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
    ///The Verify_DC_Units_Calculation_On_Changing_Base_Of_Device_To_Sounder_Base_FC recording.
    /// </summary>
    [TestModule("72a71f0a-0c60-4adc-9939-7a35ae70d424", ModuleType.Recording, 1)]
    public partial class Verify_DC_Units_Calculation_On_Changing_Base_Of_Device_To_Sounder_Base_FC : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Verify_DC_Units_Calculation_On_Changing_Base_Of_Device_To_Sounder_Base_FC instance = new Verify_DC_Units_Calculation_On_Changing_Base_Of_Device_To_Sounder_Base_FC();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_DC_Units_Calculation_On_Changing_Base_Of_Device_To_Sounder_Base_FC()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_DC_Units_Calculation_On_Changing_Base_Of_Device_To_Sounder_Base_FC Instance
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

            Libraries.Panel_Functions.AddPanelsFC(ValueConverter.ArgumentFromString<int>("NumberofPanels", "1"), "FIRECLASS 64-2", "");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeExpander("Node");
            Delay.Milliseconds(0);
            
            //Libraries.Common_Functions.ClickOnNavigationTreeExpander("FIM");
            //Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
            Delay.Milliseconds(0);
            
            //Libraries.DC_Functions.verifyDCUnitsValue("220");
            //Delay.Milliseconds(0);
            
            //Libraries.DC_Functions.verifyMaxDCUnits("2200");
            //Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyLoadingDetailsValue("220", "Current (DC Units)");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyMaxLoadingDetailsValue("2200", "Current (DC Units)");
            Delay.Milliseconds(0);
            
            Libraries.DC_Functions.VerifyDCCalculationOnChangingBase("TC_63643_Verify_DC_Units_Calculation_On_Changing_Base_Of_Device_To_Sounder_Base", "Add Devices Loop A", "Add Sounder Base Device");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelName("410LPSY - R/W - 4");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.ChangeDeviceSensitivity("Low (90dB)");
            Delay.Milliseconds(0);
            
            //Libraries.DC_Functions.verifyDCUnitsValue("284.82");
            //Delay.Milliseconds(0);
            
            //Libraries.DC_Functions.verifyDCUnitsWorstCaseValue("284.82");
            //Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyLoadingDetailsValue("284.82", "Current (DC Units)");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyLoadingDetailsValue("284.82", "Current (worst case)");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.SaveProject("TC_63643");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.Application_Close(ValueConverter.ArgumentFromString<bool>("Save", "False"), ValueConverter.ArgumentFromString<bool>("SaveConfirmation", "False"), "");
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
