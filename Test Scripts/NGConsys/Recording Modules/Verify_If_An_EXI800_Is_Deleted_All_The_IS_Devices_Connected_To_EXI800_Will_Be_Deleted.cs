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
    ///The Verify_If_An_EXI800_Is_Deleted_All_The_IS_Devices_Connected_To_EXI800_Will_Be_Deleted recording.
    /// </summary>
    [TestModule("69ebfa38-b960-483b-975f-3c0ee3f44810", ModuleType.Recording, 1)]
    public partial class Verify_If_An_EXI800_Is_Deleted_All_The_IS_Devices_Connected_To_EXI800_Will_Be_Deleted : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Verify_If_An_EXI800_Is_Deleted_All_The_IS_Devices_Connected_To_EXI800_Will_Be_Deleted instance = new Verify_If_An_EXI800_Is_Deleted_All_The_IS_Devices_Connected_To_EXI800_Will_Be_Deleted();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_If_An_EXI800_Is_Deleted_All_The_IS_Devices_Connected_To_EXI800_Will_Be_Deleted()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_If_An_EXI800_Is_Deleted_All_The_IS_Devices_Connected_To_EXI800_Will_Be_Deleted Instance
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

            Libraries.Devices_Functions.EnableISDevices();
            Delay.Milliseconds(0);
            
            Libraries.Panel_Functions.AddPanels(ValueConverter.ArgumentFromString<int>("NumberofPanels", "1"), "Pro32xD", "");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeExpander("Node");
            Delay.Milliseconds(0);
            
            //Libraries.Common_Functions.ClickOnNavigationTreeExpander("PFI");
            //Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
            Delay.Milliseconds(0);
            
            Libraries.IS_Functions.VerifyConnectedISDevicesOnEXI("TC_203_Verify_If_An_EXI800_Is_Deleted_All_The_IS_Devices_Connected_To_EXI800_Will_Be_Deleted", "Add EXI Devices", "Add IS Devices");
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
