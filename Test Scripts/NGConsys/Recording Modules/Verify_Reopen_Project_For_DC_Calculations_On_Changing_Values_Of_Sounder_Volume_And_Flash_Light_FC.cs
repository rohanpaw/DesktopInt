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
    ///The Verify_Reopen_Project_For_DC_Calculations_On_Changing_Values_Of_Sounder_Volume_And_Flash_Light_FC recording.
    /// </summary>
    [TestModule("f9e2d895-763d-4ae2-b802-8e7e842ea8af", ModuleType.Recording, 1)]
    public partial class Verify_Reopen_Project_For_DC_Calculations_On_Changing_Values_Of_Sounder_Volume_And_Flash_Light_FC : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Verify_Reopen_Project_For_DC_Calculations_On_Changing_Values_Of_Sounder_Volume_And_Flash_Light_FC instance = new Verify_Reopen_Project_For_DC_Calculations_On_Changing_Values_Of_Sounder_Volume_And_Flash_Light_FC();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_Reopen_Project_For_DC_Calculations_On_Changing_Values_Of_Sounder_Volume_And_Flash_Light_FC()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_Reopen_Project_For_DC_Calculations_On_Changing_Values_Of_Sounder_Volume_And_Flash_Light_FC Instance
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

            Report.Log(ReportLevel.Info, "Delay", "Waiting for 5s.", new RecordItemIndex(0));
            Delay.Duration(5000, false);
            
            Libraries.Common_Functions.ReopenProject("TC_63532.pjl");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Delay", "Waiting for 5s.", new RecordItemIndex(2));
            Delay.Duration(5000, false);
            
            Libraries.Common_Functions.ClickOnNavigationTreeExpander("Node");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeExpander("FIM");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelName("410LPAV 65 - 1");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.VerifyDeviceSensitivity("Low (90dB)");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelName("410LPAV 65 - 2");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.VerifyDeviceMode("1 Hz");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelName("410LPBS- R/W - 3");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.VerifyDeviceSensitivity("High (103dB)");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelName("410LPBS- R/W - 4");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.VerifyDeviceMode("1 Hz");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyLoadingDetailsValueAfterReopen("311.4", "Current (DC Units)");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyLoadingDetailsValueAfterReopen("311.4", "Current (worst case)");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.Application_Close(ValueConverter.ArgumentFromString<bool>("Save", "False"), ValueConverter.ArgumentFromString<bool>("SaveConfirmation", "False"), "");
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
