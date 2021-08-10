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

namespace Fireclass.DenmarkFCRecordingsModule
{
#pragma warning disable 0436 //(CS0436) The type 'type' in 'assembly' conflicts with the imported type 'type2' in 'assembly'. Using the type defined in 'assembly'.
    /// <summary>
    ///The Verify_FC600_Series_Panel_Gallery_Items_For_Sounder_Circuits_Nodes recording.
    /// </summary>
    [TestModule("4633d0b4-244a-4fba-aa50-4626b820ecd7", ModuleType.Recording, 1)]
    public partial class Verify_FC600_Series_Panel_Gallery_Items_For_Sounder_Circuits_Nodes : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::Fireclass.FireclassRepository repository.
        /// </summary>
        public static global::Fireclass.FireclassRepository repo = global::Fireclass.FireclassRepository.Instance;

        static Verify_FC600_Series_Panel_Gallery_Items_For_Sounder_Circuits_Nodes instance = new Verify_FC600_Series_Panel_Gallery_Items_For_Sounder_Circuits_Nodes();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_FC600_Series_Panel_Gallery_Items_For_Sounder_Circuits_Nodes()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_FC600_Series_Panel_Gallery_Items_For_Sounder_Circuits_Nodes Instance
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

            TestProject.Libraries.Panel_Functions.AddPanelsMT(ValueConverter.ArgumentFromString<int>("NumberofPanels", "1"), "FC602S", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Common_Functions.ClickOnNavigationTreeExpander("Node");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Common_Functions.ClickOnNavigationTreeItem("Sounder Circuit1");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Gallery_Functions.verifyGalleryListItems("FC_Gallery_Sounder_FC600_FC700_Series_Panel", "Denmark", "Generic Sounder");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("Generic Sounder", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Common_Functions.VerifyNavigationTreeItemText("Sounder Circuit1 (1)");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Common_Functions.ClickOnNavigationTreeItem("Sounder Circuit2");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Gallery_Functions.verifyGalleryListItems("FC_Gallery_Sounder_FC600_FC700_Series_Panel", "Denmark", "Generic Sounder");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("Generic Sounder", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Common_Functions.VerifyNavigationTreeItemText("Sounder Circuit2 (1)");
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
