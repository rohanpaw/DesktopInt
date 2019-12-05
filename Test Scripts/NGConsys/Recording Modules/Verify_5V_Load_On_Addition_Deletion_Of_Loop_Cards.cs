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
    ///The Verify_5V_Load_On_Addition_Deletion_Of_Loop_Cards recording.
    /// </summary>
    [TestModule("a9339177-ff7a-4b4f-bb8a-ee6d9c0f7f58", ModuleType.Recording, 1)]
    public partial class Verify_5V_Load_On_Addition_Deletion_Of_Loop_Cards : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Verify_5V_Load_On_Addition_Deletion_Of_Loop_Cards instance = new Verify_5V_Load_On_Addition_Deletion_Of_Loop_Cards();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_5V_Load_On_Addition_Deletion_Of_Loop_Cards()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_5V_Load_On_Addition_Deletion_Of_Loop_Cards Instance
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
            
            //Libraries.PSULoad_Functions.verify5VPSULoadValue("0.000", "PFI");
            //Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyLoadingDetailsValue("0.000", "Main Processor 5V (A)");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ProfileConsys1.tab_Points' at Center.", repo.ProfileConsys1.tab_PointsInfo, new RecordItemIndex(6));
            repo.ProfileConsys1.tab_Points.Click();
            Delay.Milliseconds(200);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Node");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromMainProcessorGallery("PLX800", "Loops", "PFI");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ProfileConsys1.tab_PhysicalLayout' at Center.", repo.ProfileConsys1.tab_PhysicalLayoutInfo, new RecordItemIndex(10));
            repo.ProfileConsys1.tab_PhysicalLayout.Click();
            Delay.Milliseconds(200);
            
            //Libraries.PSULoad_Functions.verify5VPSULoadValue("0.000", "PFI");
            //Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyLoadingDetailsValue("0.000", "Main Processor 5V (A)");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ProfileConsys1.tab_Points' at Center.", repo.ProfileConsys1.tab_PointsInfo, new RecordItemIndex(13));
            repo.ProfileConsys1.tab_Points.Click();
            Delay.Milliseconds(200);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Node");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectInventoryGridRow("6");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ProfileConsys1.btn_Delete' at Center.", repo.ProfileConsys1.btn_DeleteInfo, new RecordItemIndex(16));
            repo.ProfileConsys1.btn_Delete.Click();
            Delay.Milliseconds(200);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
            Delay.Milliseconds(0);
            
            //Libraries.PSULoad_Functions.verify5VPSULoadValue("0.020", "PFI");
            //Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyLoadingDetailsValue("0.020", "Main Processor 5V (A)");
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
