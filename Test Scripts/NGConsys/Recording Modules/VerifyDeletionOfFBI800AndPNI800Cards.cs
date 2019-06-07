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
    ///The VerifyDeletionOfFBI800AndPNI800Cards recording.
    /// </summary>
    [TestModule("5db13383-7f65-4b13-98b6-ce5b4620edfb", ModuleType.Recording, 1)]
    public partial class VerifyDeletionOfFBI800AndPNI800Cards : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static VerifyDeletionOfFBI800AndPNI800Cards instance = new VerifyDeletionOfFBI800AndPNI800Cards();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public VerifyDeletionOfFBI800AndPNI800Cards()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static VerifyDeletionOfFBI800AndPNI800Cards Instance
        {
            get { return instance; }
        }

#region Variables

        /// <summary>
        /// Gets or sets the value of variable PanelNode.
        /// </summary>
        [TestVariable("361584a9-c082-463b-bb0a-d3f851f66bcb")]
        public string PanelNode
        {
            get { return repo.PanelNode; }
            set { repo.PanelNode = value; }
        }

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
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ProfileConsys1.NavigationTree.VerifyPanelNode' at Center.", repo.ProfileConsys1.NavigationTree.VerifyPanelNodeInfo, new RecordItemIndex(2));
            repo.ProfileConsys1.NavigationTree.VerifyPanelNode.Click();
            Delay.Milliseconds(200);
            
            Libraries.Devices_Functions.AddDevicesfromMainProcessorGallery("FBI800", "Slot Cards", "PFI");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.tab_PanelAccessories' at Center.", repo.FormMe.tab_PanelAccessoriesInfo, new RecordItemIndex(4));
            repo.FormMe.tab_PanelAccessories.Click();
            Delay.Milliseconds(200);
            
            Libraries.Devices_Functions.AddDevicefromPanelAccessoriesGallery("IOB800", "Accessories");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.tab_Inventory' at Center.", repo.FormMe.tab_InventoryInfo, new RecordItemIndex(6));
            repo.FormMe.tab_Inventory.Click();
            Delay.Milliseconds(200);
            
            Libraries.InventoryGrid_Functions.verifyInventoryGridProperties("6", "Label", "FBI800-1");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ProfileConsys1.NavigationTree.Expander' at Center.", repo.ProfileConsys1.NavigationTree.ExpanderInfo, new RecordItemIndex(8));
            repo.ProfileConsys1.NavigationTree.Expander.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.BackplaneOrXLMExternalLoopCard_Expander' at Center.", repo.FormMe.BackplaneOrXLMExternalLoopCard_ExpanderInfo, new RecordItemIndex(9));
            repo.FormMe.BackplaneOrXLMExternalLoopCard_Expander.Click();
            Delay.Milliseconds(200);
            
            Libraries.SlotCards_Functions.VerifyandClickOtherSlotCardsForBackplane1("Other Slot Cards  (2 of 6)");
            Delay.Milliseconds(0);
            
            Libraries.InventoryGrid_Functions.verifyInventoryGridProperties("1", "Label", "FBI800-1");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectInventoryGridRow("1");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ProfileConsys1.btn_Delete' at Center.", repo.ProfileConsys1.btn_DeleteInfo, new RecordItemIndex(13));
            repo.ProfileConsys1.btn_Delete.Click();
            Delay.Milliseconds(200);
            
            Libraries.SlotCards_Functions.VerifyandClickOtherSlotCardsForBackplane1("Other Slot Cards  (1 of 6)");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ProfileConsys1.NavigationTree.Expander' at Center.", repo.ProfileConsys1.NavigationTree.ExpanderInfo, new RecordItemIndex(15));
            repo.ProfileConsys1.NavigationTree.Expander.Click();
            Delay.Milliseconds(200);
            
            Libraries.Devices_Functions.AddDevicesfromMainProcessorGallery("PNI800", "Slot Cards", "PFI");
            Delay.Milliseconds(0);
            
            Libraries.InventoryGrid_Functions.verifyInventoryGridProperties("7", "Label", "PNI800-1");
            Delay.Milliseconds(0);
            
            Libraries.SlotCards_Functions.VerifyandClickOtherSlotCardsForBackplane1("Other Slot Cards  (2 of 6)");
            Delay.Milliseconds(0);
            
            Libraries.InventoryGrid_Functions.verifyInventoryGridProperties("1", "Label", "PNI800-1");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectInventoryGridRow("1");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ProfileConsys1.btn_Delete' at Center.", repo.ProfileConsys1.btn_DeleteInfo, new RecordItemIndex(21));
            repo.ProfileConsys1.btn_Delete.Click();
            Delay.Milliseconds(200);
            
            Libraries.SlotCards_Functions.VerifyandClickOtherSlotCardsForBackplane1("Other Slot Cards  (1 of 6)");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.SaveProjectFromFileOption("389_5");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ReopenProject("389_5");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ProfileConsys1.NavigationTree.Expander' at Center.", repo.ProfileConsys1.NavigationTree.ExpanderInfo, new RecordItemIndex(25));
            repo.ProfileConsys1.NavigationTree.Expander.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Validation", "Validating Exists on item 'FormMe.Backplane1AfterReopen'.", repo.FormMe.Backplane1AfterReopenInfo, new RecordItemIndex(26));
            Validate.Exists(repo.FormMe.Backplane1AfterReopenInfo);
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
