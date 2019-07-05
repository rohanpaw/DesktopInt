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
    ///The Verify_Loop_Numbering recording.
    /// </summary>
    [TestModule("a0b6028c-8bd4-43ba-abc5-93b4fc948ab6", ModuleType.Recording, 1)]
    public partial class Verify_Loop_Numbering : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Verify_Loop_Numbering instance = new Verify_Loop_Numbering();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_Loop_Numbering()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_Loop_Numbering Instance
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
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ProfileConsys1.NavigationTree.VerifyPanelNode' at Center.", repo.ProfileConsys1.NavigationTree.VerifyPanelNodeInfo, new RecordItemIndex(1));
            repo.ProfileConsys1.NavigationTree.VerifyPanelNode.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ProfileConsys1.NavigationTree.Expander' at Center.", repo.ProfileConsys1.NavigationTree.ExpanderInfo, new RecordItemIndex(2));
            repo.ProfileConsys1.NavigationTree.Expander.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ProfileConsys1.NavigationTree.Expand_LoopCard' at Center.", repo.ProfileConsys1.NavigationTree.Expand_LoopCardInfo, new RecordItemIndex(3));
            repo.ProfileConsys1.NavigationTree.Expand_LoopCard.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Validation", "Validating AttributeContains (Text>'Built-in Loop-A (0 of 125)') on item 'HwndWrapperProfileConsysExe0c643c73.BuiltInLoopA0Of125'.", repo.HwndWrapperProfileConsysExe0c643c73.BuiltInLoopA0Of125Info, new RecordItemIndex(4));
            Validate.AttributeContains(repo.HwndWrapperProfileConsysExe0c643c73.BuiltInLoopA0Of125Info, "Text", "Built-in Loop-A (0 of 125)");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Validation", "Validating AttributeContains (Text>'Built-in Loop-B (0 of 125)') on item 'HwndWrapperProfileConsysExe0c643c73.BuiltInLoopB0Of125'.", repo.HwndWrapperProfileConsysExe0c643c73.BuiltInLoopB0Of125Info, new RecordItemIndex(5));
            Validate.AttributeContains(repo.HwndWrapperProfileConsysExe0c643c73.BuiltInLoopB0Of125Info, "Text", "Built-in Loop-B (0 of 125)");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Validation", "Validating AttributeContains (Text>'Built-in Loop-C (0 of 125)') on item 'HwndWrapperProfileConsysExe0c643c73.BuiltInLoopC0Of125'.", repo.HwndWrapperProfileConsysExe0c643c73.BuiltInLoopC0Of125Info, new RecordItemIndex(6));
            Validate.AttributeContains(repo.HwndWrapperProfileConsysExe0c643c73.BuiltInLoopC0Of125Info, "Text", "Built-in Loop-C (0 of 125)");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Validation", "Validating AttributeContains (Text>'Built-in Loop-D (0 of 125)') on item 'HwndWrapperProfileConsysExe0c643c73.BuiltInLoopD0Of125'.", repo.HwndWrapperProfileConsysExe0c643c73.BuiltInLoopD0Of125Info, new RecordItemIndex(7));
            Validate.AttributeContains(repo.HwndWrapperProfileConsysExe0c643c73.BuiltInLoopD0Of125Info, "Text", "Built-in Loop-D (0 of 125)");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ProfileConsys1.NavigationTree.Expander' at Center.", repo.ProfileConsys1.NavigationTree.ExpanderInfo, new RecordItemIndex(8));
            repo.ProfileConsys1.NavigationTree.Expander.Click();
            Delay.Milliseconds(200);
            
            Libraries.Devices_Functions.AddDevicesfromMainProcessorGallery("PLX800", "Loops", "PFI");
            Delay.Milliseconds(0);
            
            Libraries.InventoryGrid_Functions.verifyInventoryGridProperties("6", "Address", "E");
            Delay.Milliseconds(0);
            
            Libraries.InventoryGrid_Functions.verifyInventoryGridProperties("7", "Address", "F");
            Delay.Milliseconds(0);
            
            Libraries.InventoryGrid_Functions.verifyInventoryGridProperties("8", "Address", "G");
            Delay.Milliseconds(0);
            
            Libraries.InventoryGrid_Functions.verifyInventoryGridProperties("9", "Address", "H");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.BackplaneOrXLMExternalLoopCard_Expander' at Center.", repo.FormMe.BackplaneOrXLMExternalLoopCard_ExpanderInfo, new RecordItemIndex(14));
            repo.FormMe.BackplaneOrXLMExternalLoopCard_Expander.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.PLXExternalLoopCard_Expander' at Center.", repo.FormMe.PLXExternalLoopCard_ExpanderInfo, new RecordItemIndex(15));
            repo.FormMe.PLXExternalLoopCard_Expander.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Validation", "Validating AttributeEqual (Text='PLX800-E (0 of 125)') on item 'FormMe.MainProcessorList.PLX800E0Of125'.", repo.FormMe.MainProcessorList.PLX800E0Of125Info, new RecordItemIndex(16));
            Validate.AttributeEqual(repo.FormMe.MainProcessorList.PLX800E0Of125Info, "Text", "PLX800-E (0 of 125)");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Validation", "Validating AttributeEqual (Text='PLX800-F (0 of 125)') on item 'FormMe.MainProcessorList.PLX800F0Of125'.", repo.FormMe.MainProcessorList.PLX800F0Of125Info, new RecordItemIndex(17));
            Validate.AttributeEqual(repo.FormMe.MainProcessorList.PLX800F0Of125Info, "Text", "PLX800-F (0 of 125)");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Validation", "Validating AttributeEqual (Text='PLX800-G (0 of 125)') on item 'FormMe.MainProcessorList.PLX800G0Of125'.", repo.FormMe.MainProcessorList.PLX800G0Of125Info, new RecordItemIndex(18));
            Validate.AttributeEqual(repo.FormMe.MainProcessorList.PLX800G0Of125Info, "Text", "PLX800-G (0 of 125)");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Validation", "Validating AttributeEqual (Text='PLX800-H (0 of 125)') on item 'FormMe.MainProcessorList.PLX800H0Of125'.", repo.FormMe.MainProcessorList.PLX800H0Of125Info, new RecordItemIndex(19));
            Validate.AttributeEqual(repo.FormMe.MainProcessorList.PLX800H0Of125Info, "Text", "PLX800-H (0 of 125)");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ProfileConsys1.NavigationTree.Expander' at Center.", repo.ProfileConsys1.NavigationTree.ExpanderInfo, new RecordItemIndex(20));
            repo.ProfileConsys1.NavigationTree.Expander.Click();
            Delay.Milliseconds(200);
            
            Libraries.Devices_Functions.AddDevicesfromMainProcessorGallery("PLX800", "Loops", "PFI");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromMainProcessorGallery("PLX800", "Loops", "PFI");
            Delay.Milliseconds(0);
            
            Libraries.InventoryGrid_Functions.verifyInventoryGridProperties("6", "Address", "5");
            Delay.Milliseconds(0);
            
            Libraries.InventoryGrid_Functions.verifyInventoryGridProperties("7", "Address", "6");
            Delay.Milliseconds(0);
            
            Libraries.InventoryGrid_Functions.verifyInventoryGridProperties("8", "Address", "7");
            Delay.Milliseconds(0);
            
            Libraries.InventoryGrid_Functions.verifyInventoryGridProperties("9", "Address", "8");
            Delay.Milliseconds(0);
            
            Libraries.InventoryGrid_Functions.editDeviceLabel("6", "Label", "TEXT-5");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.SaveProjectFromFileOption("418_81");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ReopenProject("418_81");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ProfileConsys1.NavigationTree.Expander' at Center.", repo.ProfileConsys1.NavigationTree.ExpanderInfo, new RecordItemIndex(30));
            repo.ProfileConsys1.NavigationTree.Expander.Click();
            Delay.Milliseconds(200);
            
            Libraries.InventoryGrid_Functions.verifyInventoryGridProperties("6", "Label", "TEXT-5");
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
