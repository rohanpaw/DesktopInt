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

namespace TestProject.DenmarkRecordingModules
{
#pragma warning disable 0436 //(CS0436) The type 'type' in 'assembly' conflicts with the imported type 'type2' in 'assembly'. Using the type defined in 'assembly'.
    /// <summary>
    ///The Verify_Pro_Panel_Site_Node_Gallery_And_Addition_Of_Panels_And_Devices recording.
    /// </summary>
    [TestModule("d74ff49a-f567-40fd-abd7-5e4577a644bd", ModuleType.Recording, 1)]
    public partial class Verify_Pro_Panel_Site_Node_Gallery_And_Addition_Of_Panels_And_Devices : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Verify_Pro_Panel_Site_Node_Gallery_And_Addition_Of_Panels_And_Devices instance = new Verify_Pro_Panel_Site_Node_Gallery_And_Addition_Of_Panels_And_Devices();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_Pro_Panel_Site_Node_Gallery_And_Addition_Of_Panels_And_Devices()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_Pro_Panel_Site_Node_Gallery_And_Addition_Of_Panels_And_Devices Instance
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

            Libraries.Common_Functions.ClickOnNavigationTreeItem("Site");
            Delay.Milliseconds(0);
            
            Libraries.Gallery_Functions.verifyGalleryListItems("Gallery_Panels", "Denmark", "Pro32xD");
            Delay.Milliseconds(0);
            
            Libraries.Gallery_Functions.verifyGalleryListItems("Gallery_OtherNodes", "Denmark", "TXG Node");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Site");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.clickOnSiteAccessoriesTab();
            Delay.Milliseconds(0);
            
            Libraries.Gallery_Functions.verifyGalleryListItems("Gallery_Accessories", "Denmark", "ANC125");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("ANC125", "Accessories");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("ANC250", "Accessories");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.clickOnShoppingListTab();
            Delay.Milliseconds(0);
            
            //Libraries.Export_Functions.SearchDeviceInExportUsingSKUOrDescription("557.202.622", ValueConverter.ArgumentFromString<bool>("sExist", "True"));
            //Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
