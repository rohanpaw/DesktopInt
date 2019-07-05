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
    ///The Repen_and_Verify_Logic_Of_Slot_Loop_Cards_Addition_In_Panels_Pro16xD_Pro16xBB_Pro815D_Pro885D_Pro815DCH_Pro21 recording.
    /// </summary>
    [TestModule("0a83c020-8e5c-4235-a7c3-66d8bf0ee58c", ModuleType.Recording, 1)]
    public partial class Repen_and_Verify_Logic_Of_Slot_Loop_Cards_Addition_In_Panels_Pro16xD_Pro16xBB_Pro815D_Pro885D_Pro815DCH_Pro21 : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Repen_and_Verify_Logic_Of_Slot_Loop_Cards_Addition_In_Panels_Pro16xD_Pro16xBB_Pro815D_Pro885D_Pro815DCH_Pro21 instance = new Repen_and_Verify_Logic_Of_Slot_Loop_Cards_Addition_In_Panels_Pro16xD_Pro16xBB_Pro815D_Pro885D_Pro815DCH_Pro21();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Repen_and_Verify_Logic_Of_Slot_Loop_Cards_Addition_In_Panels_Pro16xD_Pro16xBB_Pro815D_Pro885D_Pro815DCH_Pro21()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Repen_and_Verify_Logic_Of_Slot_Loop_Cards_Addition_In_Panels_Pro16xD_Pro16xBB_Pro815D_Pro885D_Pro815DCH_Pro21 Instance
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

            Libraries.Common_Functions.ReopenProject("TC_216_TC_02_03");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.Node1Expander_AfterReopen' at Center.", repo.FormMe.Node1Expander_AfterReopenInfo, new RecordItemIndex(1));
            repo.FormMe.Node1Expander_AfterReopen.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.Backplane1Expander_AfterReopen' at Center.", repo.FormMe.Backplane1Expander_AfterReopenInfo, new RecordItemIndex(2));
            repo.FormMe.Backplane1Expander_AfterReopen.Click();
            Delay.Milliseconds(200);
            
            Libraries.SlotCards_Functions.VerifyLoopCardDistributionInBackplaneOnReopen(ValueConverter.ArgumentFromString<int>("MaxLimitOfDevice", "3"));
            Delay.Milliseconds(0);
            
            Libraries.SlotCards_Functions.VerifyandClickOtherSlotCardsForBackplane1OnReopen("Other Slot Cards  (1 of 1)");
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
