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
    ///The Verify_The_DC_Calculations_on_Changing_Values_Of_Sounder_Volume_and_Flash_Light recording.
    /// </summary>
    [TestModule("5bd96ef8-bd6f-436f-b45f-bf57de82c162", ModuleType.Recording, 1)]
    public partial class Verify_The_DC_Calculations_on_Changing_Values_Of_Sounder_Volume_and_Flash_Light : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Verify_The_DC_Calculations_on_Changing_Values_Of_Sounder_Volume_and_Flash_Light instance = new Verify_The_DC_Calculations_on_Changing_Values_Of_Sounder_Volume_and_Flash_Light();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_The_DC_Calculations_on_Changing_Values_Of_Sounder_Volume_and_Flash_Light()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_The_DC_Calculations_on_Changing_Values_Of_Sounder_Volume_and_Flash_Light Instance
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
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.NodeExpander1' at Center.", repo.FormMe.NodeExpander1Info, new RecordItemIndex(1));
            repo.FormMe.NodeExpander1.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.LoopExpander1' at Center.", repo.FormMe.LoopExpander1Info, new RecordItemIndex(2));
            repo.FormMe.LoopExpander1.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse XButton2 Click item 'FormMe.Loop_A1' at Center.", repo.FormMe.Loop_A1Info, new RecordItemIndex(3));
            repo.FormMe.Loop_A1.Click(System.Windows.Forms.MouseButtons.XButton2);
            Delay.Milliseconds(200);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("LPBS 3000", "Sounders/Beacons");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelName("LPBS 3000 - 1");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.VerifyDeviceSensitivity("High (90dB)");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.ChangeDeviceSensitivity("Mid Range Low (70dB)");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.ChangeDeviceSensitivity("Mid Range High (80dB)");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.ChangeDeviceSensitivity("Low (60dB)");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelName("LPBS 3000 - 2");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.VerifyDeviceMode("0.5 Hz");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.ChangeDeviceMode("1 Hz");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.ChangeDeviceMode("0.5 Hz");
            Delay.Milliseconds(0);
            
            Libraries.DC_Functions.verifyDCUnitsValue("36");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.SaveProject("NGC-603");
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
