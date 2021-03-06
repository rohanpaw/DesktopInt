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
    ///The Testing recording.
    /// </summary>
    [TestModule("d173c66d-c7b7-47fb-ab2e-60be9cff88e6", ModuleType.Recording, 1)]
    public partial class Testing : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Testing instance = new Testing();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Testing()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Testing Instance
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

            Libraries.Panel_Functions.AddPanels(ValueConverter.ArgumentFromString<int>("NumberofPanels", "1"), "Pro32xD", "");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeExpander("Node");
            Delay.Milliseconds(0);
            
            //Libraries.Devices_Functions.EnableISDevices();
            //Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("801 CH", "Detectors");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.clickOnPanelCalculationsTab();
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyLoopLoadingDetailsValue("1", "Built-in Loop-A", "1");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyLoopLoadingDetailsValue("0.08", "Built-in Loop-A", "4");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyMaxLoopLoadingDetailsValue("4000", "Built-in Loop-A", "2");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyMaxLoopLoadingDetailsValue("250", "Built-in Loop-A", "1");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyMaxLoopLoadingDetailsValue("14.40", "Built-in Loop-A", "3");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyMaxLoopLoadingDetailsValue("2200", "Built-in Loop-A", "2");
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
