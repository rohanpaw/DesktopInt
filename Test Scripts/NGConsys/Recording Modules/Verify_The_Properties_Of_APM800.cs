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
    ///The Verify_The_Properties_Of_APM800 recording.
    /// </summary>
    [TestModule("c620bf21-83c7-4d0f-9160-e0373be4eb18", ModuleType.Recording, 1)]
    public partial class Verify_The_Properties_Of_APM800 : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Verify_The_Properties_Of_APM800 instance = new Verify_The_Properties_Of_APM800();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_The_Properties_Of_APM800()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_The_Properties_Of_APM800 Instance
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

            Libraries.Panel_Functions.AddPanels(ValueConverter.ArgumentFromString<int>("NumberofPanels", "1"), "MX4000", "");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.PanelNode1' at Center.", repo.FormMe.PanelNode1Info, new RecordItemIndex(1));
            repo.FormMe.PanelNode1.Click();
            Delay.Milliseconds(200);
            
            Libraries.Panel_Functions.ChangePSUType("PSB800");
            Delay.Milliseconds(0);
            
            Libraries.Panel_Functions.ChangeSecondPSUType("PSB800-K");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.NodeExpander1' at Center.", repo.FormMe.NodeExpander1Info, new RecordItemIndex(4));
            repo.FormMe.NodeExpander1.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.LoopExpander1' at Center.", repo.FormMe.LoopExpander1Info, new RecordItemIndex(5));
            repo.FormMe.LoopExpander1.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.Loop_A1' at Center.", repo.FormMe.Loop_A1Info, new RecordItemIndex(6));
            repo.FormMe.Loop_A1.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ProfileConsys1.tab_Points' at Center.", repo.ProfileConsys1.tab_PointsInfo, new RecordItemIndex(7));
            repo.ProfileConsys1.tab_Points.Click();
            Delay.Milliseconds(200);
            
            Libraries.Devices_Functions.SelectRowUsingLabelName("APM 800 - 1");
            Delay.Milliseconds(0);
            
            //Libraries.Devices_Functions.verifyDescription("APM800 part Of second PSU");
            //Delay.Milliseconds(0);
            
            Libraries.InventoryGrid_Functions.verifyInventoryDeviceProperty("Zone", "1 (Zone 01)");
            Delay.Milliseconds(0);
            
            Libraries.InventoryGrid_Functions.verifyInventoryDeviceProperty("Address", "1");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.editDeviceAddress("Address", "2");
            Delay.Milliseconds(0);
            
            Libraries.InventoryGrid_Functions.verifyInventoryDeviceProperty("Address", "2");
            Delay.Milliseconds(0);
            
            Libraries.InventoryGrid_Functions.verifyInventoryGridProperties("1", "Address", "2");
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
