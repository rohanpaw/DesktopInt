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
    ///The Verify_Voltage_Drop_on_Other_Loops_for_PFI recording.
    /// </summary>
    [TestModule("24868221-cbc4-40c8-927e-77d4ec6d7e1e", ModuleType.Recording, 1)]
    public partial class Verify_Voltage_Drop_on_Other_Loops_for_PFI : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Verify_Voltage_Drop_on_Other_Loops_for_PFI instance = new Verify_Voltage_Drop_on_Other_Loops_for_PFI();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_Voltage_Drop_on_Other_Loops_for_PFI()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_Voltage_Drop_on_Other_Loops_for_PFI Instance
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
            
            Libraries.Common_Functions.ClickOnNavigationTreeExpander("PFI");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
            Delay.Milliseconds(0);
            
            Libraries.VoltageDrop_Functions.verifyVoltageDropCalculation("Verify Voltage Drop Calculation on adding devices in Multiple loops", "Add Devices Loop A");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-B");
            Delay.Milliseconds(0);
            
            verifyVoltDropValue("0.07");
            Delay.Milliseconds(0);
            
            verifyVoltDropWorstCaseValue("0.15");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-B");
            Delay.Milliseconds(0);
            
            Libraries.VoltageDrop_Functions.verifyVoltageDropCalculation("Verify Voltage Drop Calculation on adding devices in Multiple loops", "Add Devices Loop B");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-C");
            Delay.Milliseconds(0);
            
            verifyVoltDropValue("0.00");
            Delay.Milliseconds(0);
            
            verifyVoltDropWorstCaseValue("0.00");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-C");
            Delay.Milliseconds(0);
            
            Libraries.VoltageDrop_Functions.verifyVoltageDropCalculation("Verify Voltage Drop Calculation on adding devices in Multiple loops", "Add Devices Loop C");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-D");
            Delay.Milliseconds(0);
            
            Libraries.VoltageDrop_Functions.verifyVoltageDropCalculation("Verify Voltage Drop Calculation on adding devices in Multiple loops", "Add Devices Loop D");
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
