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
    ///The Verify_APM800_Should_Not_Get_Added_To_Shopping_List_When_Second_PSU_Selected_At_The_Time_Of_Panel_Se recording.
    /// </summary>
    [TestModule("7f577429-755c-43ce-83fb-59270bc8b76b", ModuleType.Recording, 1)]
    public partial class Verify_APM800_Should_Not_Get_Added_To_Shopping_List_When_Second_PSU_Selected_At_The_Time_Of_Panel_Se : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Verify_APM800_Should_Not_Get_Added_To_Shopping_List_When_Second_PSU_Selected_At_The_Time_Of_Panel_Se instance = new Verify_APM800_Should_Not_Get_Added_To_Shopping_List_When_Second_PSU_Selected_At_The_Time_Of_Panel_Se();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_APM800_Should_Not_Get_Added_To_Shopping_List_When_Second_PSU_Selected_At_The_Time_Of_Panel_Se()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_APM800_Should_Not_Get_Added_To_Shopping_List_When_Second_PSU_Selected_At_The_Time_Of_Panel_Se Instance
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

            //Libraries.Panel_Functions.AddPanelAndAddCPUAndPSU(ValueConverter.ArgumentFromString<int>("NumberofPanels", "1"), "MX1000", "");
            //Delay.Milliseconds(0);
            
            Libraries.Panel_Functions.AddPanelAndPSUDuringPanelSelection("MX1000", "PSB800", "PSB800-KM");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Node");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeExpander("Node");
            Delay.Milliseconds(0);
            
            //Libraries.Common_Functions.ClickOnNavigationTreeExpander("FIM");
            //Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.clickOnPointsTab();
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.VerifyDeviceUsingLabelName("APM 800 - 1");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Site");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.clickOnShoppingListTab();
            Delay.Milliseconds(0);
            
            Libraries.Export_Functions.SearchDeviceInExportUsingSKUOrDescription("557.202.027", ValueConverter.ArgumentFromString<bool>("sExist", "False"));
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
