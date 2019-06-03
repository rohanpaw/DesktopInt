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
    ///The Verify_FB_IOB_Slot_Allocation_In_Backplane recording.
    /// </summary>
    [TestModule("97a6fcd4-06d8-4d1a-822d-39f831aa97f0", ModuleType.Recording, 1)]
    public partial class Verify_FB_IOB_Slot_Allocation_In_Backplane : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Verify_FB_IOB_Slot_Allocation_In_Backplane instance = new Verify_FB_IOB_Slot_Allocation_In_Backplane();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_FB_IOB_Slot_Allocation_In_Backplane()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_FB_IOB_Slot_Allocation_In_Backplane Instance
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
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ProfileConsys1.NavigationTree.Expander' at Center.", repo.ProfileConsys1.NavigationTree.ExpanderInfo, new RecordItemIndex(1));
            repo.ProfileConsys1.NavigationTree.Expander.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.tab_PanelAccessories' at Center.", repo.FormMe.tab_PanelAccessoriesInfo, new RecordItemIndex(2));
            repo.FormMe.tab_PanelAccessories.Click();
            Delay.Milliseconds(200);
            
            Libraries.Devices_Functions.AddDevicefromPanelAccessoriesGallery("FB800", "Accessories");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.BackplaneOrXLMExternalLoopCard_Expander' at Center.", repo.FormMe.BackplaneOrXLMExternalLoopCard_ExpanderInfo, new RecordItemIndex(4));
            repo.FormMe.BackplaneOrXLMExternalLoopCard_Expander.Click();
            Delay.Milliseconds(200);
            
            Libraries.SlotCards_Functions.VerifyandClickOtherSlotCardsForBackplane1("Other Slot Cards  (1 of 6)");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelName("FB800-1");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ProfileConsys1.NavigationTree.Expander' at Center.", repo.ProfileConsys1.NavigationTree.ExpanderInfo, new RecordItemIndex(7));
            repo.ProfileConsys1.NavigationTree.Expander.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.tab_PanelAccessories' at Center.", repo.FormMe.tab_PanelAccessoriesInfo, new RecordItemIndex(8));
            repo.FormMe.tab_PanelAccessories.Click();
            Delay.Milliseconds(200);
            
            Libraries.Devices_Functions.VerifyPanelNodePanelAccessoriesGallery("FB800", "Accessories", "Disabled");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicefromPanelAccessoriesGallery("IOB800", "Accessories");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.BackplaneOrXLMExternalLoopCard_Expander' at Center.", repo.FormMe.BackplaneOrXLMExternalLoopCard_ExpanderInfo, new RecordItemIndex(11));
            repo.FormMe.BackplaneOrXLMExternalLoopCard_Expander.Click();
            Delay.Milliseconds(200);
            
            Libraries.SlotCards_Functions.VerifyandClickOtherSlotCardsForBackplane1("Other Slot Cards  (1 of 6)");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelName("IOB800-1");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ProfileConsys1.NavigationTree.Expander' at Center.", repo.ProfileConsys1.NavigationTree.ExpanderInfo, new RecordItemIndex(14));
            repo.ProfileConsys1.NavigationTree.Expander.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.tab_PanelAccessories' at Center.", repo.FormMe.tab_PanelAccessoriesInfo, new RecordItemIndex(15));
            repo.FormMe.tab_PanelAccessories.Click();
            Delay.Milliseconds(200);
            
            Libraries.Devices_Functions.AddDevicefromPanelAccessoriesGallery("IOB800", "Accessories");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.VerifyPanelNodePanelAccessoriesGallery("IOB800", "Accessories", "Disabled");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.BackplaneOrXLMExternalLoopCard_Expander' at Center.", repo.FormMe.BackplaneOrXLMExternalLoopCard_ExpanderInfo, new RecordItemIndex(18));
            repo.FormMe.BackplaneOrXLMExternalLoopCard_Expander.Click();
            Delay.Milliseconds(200);
            
            Libraries.SlotCards_Functions.VerifyandClickOtherSlotCardsForBackplane1("Other Slot Cards  (2 of 6)");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ProfileConsys1.NavigationTree.Expander' at Center.", repo.ProfileConsys1.NavigationTree.ExpanderInfo, new RecordItemIndex(20));
            repo.ProfileConsys1.NavigationTree.Expander.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.tab_PanelAccessories' at Center.", repo.FormMe.tab_PanelAccessoriesInfo, new RecordItemIndex(21));
            repo.FormMe.tab_PanelAccessories.Click();
            Delay.Milliseconds(200);
            
            Libraries.Devices_Functions.SelectPanelNodePanelAccessoriesRow("1");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Delay", "Waiting for 10s.", new RecordItemIndex(23));
            Delay.Duration(10000, false);
            
            Libraries.Devices_Functions.SelectPanelNodePanelAccessoriesRow("1");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ProfileConsys1.btn_Delete' at Center.", repo.ProfileConsys1.btn_DeleteInfo, new RecordItemIndex(25));
            repo.ProfileConsys1.btn_Delete.Click();
            Delay.Milliseconds(200);
            
            Libraries.Devices_Functions.VerifyPanelNodePanelAccessoriesGallery("FB800", "Accessories", "Enabled");
            Delay.Milliseconds(0);
            
            Libraries.Panel_Functions.DeletePanel(ValueConverter.ArgumentFromString<int>("NumberofPanels", "1"), "PanelNode", ValueConverter.ArgumentFromString<int>("rowNumber", "1"));
            Delay.Milliseconds(0);
            
            Libraries.Panel_Functions.AddPanels(ValueConverter.ArgumentFromString<int>("NumberofPanels", "1"), "Pro32xBB", "");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ProfileConsys1.NavigationTree.Expander' at Center.", repo.ProfileConsys1.NavigationTree.ExpanderInfo, new RecordItemIndex(29));
            repo.ProfileConsys1.NavigationTree.Expander.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.tab_PanelAccessories' at Center.", repo.FormMe.tab_PanelAccessoriesInfo, new RecordItemIndex(30));
            repo.FormMe.tab_PanelAccessories.Click();
            Delay.Milliseconds(200);
            
            Libraries.Devices_Functions.AddDevicefromPanelAccessoriesGallery("IOB800", "Accessories");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.BackplaneOrXLMExternalLoopCard_Expander' at Center.", repo.FormMe.BackplaneOrXLMExternalLoopCard_ExpanderInfo, new RecordItemIndex(32));
            repo.FormMe.BackplaneOrXLMExternalLoopCard_Expander.Click();
            Delay.Milliseconds(200);
            
            Libraries.SlotCards_Functions.VerifyandClickOtherSlotCardsForBackplane1("Other Slot Cards  (1 of 6)");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelName("IOB800-1");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ProfileConsys1.NavigationTree.Expander' at Center.", repo.ProfileConsys1.NavigationTree.ExpanderInfo, new RecordItemIndex(35));
            repo.ProfileConsys1.NavigationTree.Expander.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.tab_PanelAccessories' at Center.", repo.FormMe.tab_PanelAccessoriesInfo, new RecordItemIndex(36));
            repo.FormMe.tab_PanelAccessories.Click();
            Delay.Milliseconds(200);
            
            Libraries.Devices_Functions.AddDevicefromPanelAccessoriesGallery("FB800", "Accessories");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.BackplaneOrXLMExternalLoopCard_Expander' at Center.", repo.FormMe.BackplaneOrXLMExternalLoopCard_ExpanderInfo, new RecordItemIndex(38));
            repo.FormMe.BackplaneOrXLMExternalLoopCard_Expander.Click();
            Delay.Milliseconds(200);
            
            Libraries.SlotCards_Functions.VerifyandClickOtherSlotCardsForBackplane1("Other Slot Cards  (1 of 6)");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelName("FB800-1");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ProfileConsys1.NavigationTree.Expander' at Center.", repo.ProfileConsys1.NavigationTree.ExpanderInfo, new RecordItemIndex(41));
            repo.ProfileConsys1.NavigationTree.Expander.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.tab_PanelAccessories' at Center.", repo.FormMe.tab_PanelAccessoriesInfo, new RecordItemIndex(42));
            repo.FormMe.tab_PanelAccessories.Click();
            Delay.Milliseconds(200);
            
            Libraries.Devices_Functions.VerifyPanelNodePanelAccessoriesGallery("FB800", "Accessories", "Disabled");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicefromPanelAccessoriesGallery("IOB800", "Accessories");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.BackplaneOrXLMExternalLoopCard_Expander' at Center.", repo.FormMe.BackplaneOrXLMExternalLoopCard_ExpanderInfo, new RecordItemIndex(45));
            repo.FormMe.BackplaneOrXLMExternalLoopCard_Expander.Click();
            Delay.Milliseconds(200);
            
            Libraries.SlotCards_Functions.VerifyandClickOtherSlotCardsForBackplane1("Other Slot Cards  (2 of 6)");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ProfileConsys1.NavigationTree.Expander' at Center.", repo.ProfileConsys1.NavigationTree.ExpanderInfo, new RecordItemIndex(47));
            repo.ProfileConsys1.NavigationTree.Expander.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.tab_PanelAccessories' at Center.", repo.FormMe.tab_PanelAccessoriesInfo, new RecordItemIndex(48));
            repo.FormMe.tab_PanelAccessories.Click();
            Delay.Milliseconds(200);
            
            Libraries.Devices_Functions.VerifyPanelNodePanelAccessoriesGallery("FB800", "Accessories", "Disabled");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.VerifyPanelNodePanelAccessoriesGallery("IOB800", "Accessories", "Disabled");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectPanelNodePanelAccessoriesRow("3");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ProfileConsys1.btn_Delete' at Center.", repo.ProfileConsys1.btn_DeleteInfo, new RecordItemIndex(52));
            repo.ProfileConsys1.btn_Delete.Click();
            Delay.Milliseconds(200);
            
            Libraries.Devices_Functions.VerifyPanelNodePanelAccessoriesGallery("IOB800", "Accessories", "Enabled");
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}