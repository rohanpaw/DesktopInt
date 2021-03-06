///////////////////////////////////////////////////////////////////////////////
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
    ///The Verify_Multiple_PCH_Can_Be_Added_On_Same_Backplane recording.
    /// </summary>
    [TestModule("570e3bf8-e9f9-4539-8d50-5a61909473be", ModuleType.Recording, 1)]
    public partial class Verify_Multiple_PCH_Can_Be_Added_On_Same_Backplane : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Verify_Multiple_PCH_Can_Be_Added_On_Same_Backplane instance = new Verify_Multiple_PCH_Can_Be_Added_On_Same_Backplane();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_Multiple_PCH_Can_Be_Added_On_Same_Backplane()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_Multiple_PCH_Can_Be_Added_On_Same_Backplane Instance
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

            Libraries.SlotCards_Functions.VerifySlotCardsAndBackplanesDistributionWithOnePanel("TC_76704_Verify Multiple PCH can be added on same backplane", "Add Devices");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.SaveProject("TC_76704");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.Application_Close(ValueConverter.ArgumentFromString<bool>("Save", "False"), ValueConverter.ArgumentFromString<bool>("SaveConfirmation", "False"), "");
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
