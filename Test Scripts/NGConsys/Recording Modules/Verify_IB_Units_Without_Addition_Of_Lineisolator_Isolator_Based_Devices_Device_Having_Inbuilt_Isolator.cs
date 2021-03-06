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
    ///The Verify_IB_Units_Without_Addition_Of_Lineisolator_Isolator_Based_Devices_Device_Having_Inbuilt_Isolator recording.
    /// </summary>
    [TestModule("ed95de78-c489-4cd5-af64-ccb8d4aadac0", ModuleType.Recording, 1)]
    public partial class Verify_IB_Units_Without_Addition_Of_Lineisolator_Isolator_Based_Devices_Device_Having_Inbuilt_Isolator : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Verify_IB_Units_Without_Addition_Of_Lineisolator_Isolator_Based_Devices_Device_Having_Inbuilt_Isolator instance = new Verify_IB_Units_Without_Addition_Of_Lineisolator_Isolator_Based_Devices_Device_Having_Inbuilt_Isolator();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_IB_Units_Without_Addition_Of_Lineisolator_Isolator_Based_Devices_Device_Having_Inbuilt_Isolator()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_IB_Units_Without_Addition_Of_Lineisolator_Isolator_Based_Devices_Device_Having_Inbuilt_Isolator Instance
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

            Libraries.Panel_Functions.AddPanelsFC(ValueConverter.ArgumentFromString<int>("NumberofPanels", "1"), "FIRECLASS 64-2", "");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeExpander("Node");
            Delay.Milliseconds(0);
            
            //Libraries.Common_Functions.ClickOnNavigationTreeExpander("FIM");
            //Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("460PH", "Detectors");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("410BDM", "Detectors");
            Delay.Milliseconds(0);
            
            Libraries.IS_Functions.VerifyIsolatorUnits("3", "FIM");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Node");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
            Delay.Milliseconds(0);
            
            //Libraries.Devices_Functions.AddDevicesfromMultiplePointWizard("400PH", ValueConverter.ArgumentFromString<int>("DeviceQty", "5"));
            //Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.RightClickOnSelectedRow("2");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.clickContextMenuOptionOnRightClick("8");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromMultiplePointWizard("410RIM", ValueConverter.ArgumentFromString<int>("DeviceQty", "5"));
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelName("410BDM - 2");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.RightClickOnSelectedRow("2");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.clickContextMenuOptionOnRightClick("2");
            Delay.Milliseconds(0);
            
            //Libraries.Common_Functions.clickOnCutButton();
            //Delay.Milliseconds(0);
            
            Libraries.IS_Functions.VerifyIsolatorUnits("26", "FIM");
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
