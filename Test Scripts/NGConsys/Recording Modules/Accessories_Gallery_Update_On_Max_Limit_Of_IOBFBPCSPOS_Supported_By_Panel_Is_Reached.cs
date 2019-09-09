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
    ///The Accessories_Gallery_Update_On_Max_Limit_Of_IOBFBPCSPOS_Supported_By_Panel_Is_Reached recording.
    /// </summary>
    [TestModule("d7004e74-b6ad-48ca-9651-bdd5f6064efd", ModuleType.Recording, 1)]
    public partial class Accessories_Gallery_Update_On_Max_Limit_Of_IOBFBPCSPOS_Supported_By_Panel_Is_Reached : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Accessories_Gallery_Update_On_Max_Limit_Of_IOBFBPCSPOS_Supported_By_Panel_Is_Reached instance = new Accessories_Gallery_Update_On_Max_Limit_Of_IOBFBPCSPOS_Supported_By_Panel_Is_Reached();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Accessories_Gallery_Update_On_Max_Limit_Of_IOBFBPCSPOS_Supported_By_Panel_Is_Reached()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Accessories_Gallery_Update_On_Max_Limit_Of_IOBFBPCSPOS_Supported_By_Panel_Is_Reached Instance
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
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Node");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.tab_PanelAccessories' at Center.", repo.FormMe.tab_PanelAccessoriesInfo, new RecordItemIndex(2));
            repo.FormMe.tab_PanelAccessories.Click();
            Delay.Milliseconds(200);
            
            Libraries.Devices_Functions.AddAndVerifyMaxNumberOfPanelAccessories("TC_61_Verify_Accessories_Gallery_Update_On_Max_Limit_Of_IOB_FB_PCS_POS_Supported_By_Panel_Is_Reached", "Add Panels");
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
