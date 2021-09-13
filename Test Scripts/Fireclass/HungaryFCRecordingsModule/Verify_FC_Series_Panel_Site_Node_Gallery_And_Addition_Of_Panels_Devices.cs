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

namespace Fireclass.HungaryFCRecordingsModule
{
#pragma warning disable 0436 //(CS0436) The type 'type' in 'assembly' conflicts with the imported type 'type2' in 'assembly'. Using the type defined in 'assembly'.
    /// <summary>
    ///The Verify_FC_Series_Panel_Site_Node_Gallery_And_Addition_Of_Panels_Devices recording.
    /// </summary>
    [TestModule("2b32989a-8333-4a4d-9f8b-924c8b07d8b8", ModuleType.Recording, 1)]
    public partial class Verify_FC_Series_Panel_Site_Node_Gallery_And_Addition_Of_Panels_Devices : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::Fireclass.FireclassRepository repository.
        /// </summary>
        public static global::Fireclass.FireclassRepository repo = global::Fireclass.FireclassRepository.Instance;

        static Verify_FC_Series_Panel_Site_Node_Gallery_And_Addition_Of_Panels_Devices instance = new Verify_FC_Series_Panel_Site_Node_Gallery_And_Addition_Of_Panels_Devices();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_FC_Series_Panel_Site_Node_Gallery_And_Addition_Of_Panels_Devices()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_FC_Series_Panel_Site_Node_Gallery_And_Addition_Of_Panels_Devices Instance
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

            TestProject.Libraries.Common_Functions.ClickOnNavigationTreeItem("Site");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Gallery_Functions.verifyGalleryListItems("FC_Gallery_Panels", "Hungary", "FC702S");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Gallery_Functions.verifyGalleryListItems("FC_Gallery_OtherNodes", "Hungary", "3rd Party Interface");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Common_Functions.ClickOnNavigationTreeItem("Site");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Common_Functions.clickOnSiteAccessoriesTab();
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Gallery_Functions.verifyGalleryListItems("FC_Gallery_Accessories", "Hungary", "ANC1");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("ANC1-D", "Accessories");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("FC-ANC-E", "Accessories");
            Delay.Milliseconds(0);
            
            //TestProject.Libraries.Common_Functions.clickOnShoppingListTab();
            //Delay.Milliseconds(0);
            
            //TestProject.Libraries.Export_Functions.SearchDeviceInExportUsingSKUOrDescription("557.202.622", ValueConverter.ArgumentFromString<bool>("sExist", "True"));
            //Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
