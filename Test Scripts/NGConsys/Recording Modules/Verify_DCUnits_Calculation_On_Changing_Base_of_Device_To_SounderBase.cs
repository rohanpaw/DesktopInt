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
    ///The Verify_DCUnits_Calculation_On_Changing_Base_of_Device_To_SounderBase recording.
    /// </summary>
    [TestModule("af5c6985-919e-4aad-8e11-50a102b911f7", ModuleType.Recording, 1)]
    public partial class Verify_DCUnits_Calculation_On_Changing_Base_of_Device_To_SounderBase : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Verify_DCUnits_Calculation_On_Changing_Base_of_Device_To_SounderBase instance = new Verify_DCUnits_Calculation_On_Changing_Base_of_Device_To_SounderBase();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_DCUnits_Calculation_On_Changing_Base_of_Device_To_SounderBase()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_DCUnits_Calculation_On_Changing_Base_of_Device_To_SounderBase Instance
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
            
            //Libraries.DC_Functions.verifyDCUnitsValue("220");
            //Delay.Milliseconds(0);
            
            //Libraries.DC_Functions.verifyMaxDCUnits("2200");
            //Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyLoadingDetailsValue("220", "Current (DC Units)");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyMaxLoadingDetailsValue("2200", "Current (DC Units)");
            Delay.Milliseconds(0);
            
            Libraries.DC_Functions.VerifyDCCalculationOnChangingBase("Verify DC Units Calculation on changing base of device to sounder base", "Add Devices Loop A", "Add Sounder Base Device");
            Delay.Milliseconds(0);
            
            //Libraries.DC_Functions.verifyDCUnitsValue("384.9");
            //Delay.Milliseconds(0);
            
            //Libraries.DC_Functions.verifyDCUnitsWorstCaseValue("384.9");
            //Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyLoadingDetailsValue("384.9", "Current (DC Units)");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyLoadingDetailsValue("384.9", "Current (worst case)");
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
