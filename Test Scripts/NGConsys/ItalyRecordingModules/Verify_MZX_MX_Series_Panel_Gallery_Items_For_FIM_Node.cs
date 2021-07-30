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

namespace TestProject.ItalyRecordingModules
{
#pragma warning disable 0436 //(CS0436) The type 'type' in 'assembly' conflicts with the imported type 'type2' in 'assembly'. Using the type defined in 'assembly'.
    /// <summary>
    ///The Verify_MZX_MX_Series_Panel_Gallery_Items_For_FIM_Node recording.
    /// </summary>
    [TestModule("67a7b131-d680-4509-8be2-97cdcfac5d20", ModuleType.Recording, 1)]
    public partial class Verify_MZX_MX_Series_Panel_Gallery_Items_For_FIM_Node : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Verify_MZX_MX_Series_Panel_Gallery_Items_For_FIM_Node instance = new Verify_MZX_MX_Series_Panel_Gallery_Items_For_FIM_Node();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_MZX_MX_Series_Panel_Gallery_Items_For_FIM_Node()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_MZX_MX_Series_Panel_Gallery_Items_For_FIM_Node Instance
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

            Libraries.Panel_Functions.AddPanels(ValueConverter.ArgumentFromString<int>("NumberofPanels", "1"), "MZX125", "");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeExpander("Node");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("FIM");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingModelNameFromInventoryTab("FIM800");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelNameFromInventoryTab("I/O Locale");
            Delay.Milliseconds(0);
            
            // Attached Functionality
            Libraries.Gallery_Functions.verifyGalleryListItems("Gallery_AttachedFunctionality_FIM_Node", "Italy", "IOB800(x1)");
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
