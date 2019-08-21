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
    ///The Verify_devices_order_for_Sounders_in_Non_dropped_gallery recording.
    /// </summary>
    [TestModule("d5921199-49c5-4e6c-a74b-fa8920dacd22", ModuleType.Recording, 1)]
    public partial class Verify_devices_order_for_Sounders_in_Non_dropped_gallery : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Verify_devices_order_for_Sounders_in_Non_dropped_gallery instance = new Verify_devices_order_for_Sounders_in_Non_dropped_gallery();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_devices_order_for_Sounders_in_Non_dropped_gallery()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_devices_order_for_Sounders_in_Non_dropped_gallery Instance
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

            Libraries.Panel_Functions.AddPanels(ValueConverter.ArgumentFromString<int>("NumberofPanels", "1"), "Pro32xD", "");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeExpander("Node");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeExpander("PFI");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("LPS 800", "Sounders/Beacons");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelName("LPS 800 - 1");
            Delay.Milliseconds(0);
            
            Libraries.Gallery_Functions.verifyNonDroppedGallery("Conventional Sounders", "Squashni Sounder", "Generic Sounder");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelName("Generic Sounder");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.changeAndVerifyAlarmLoad(ValueConverter.ArgumentFromString<int>("AlarmLoad", "10"), "Valid", ValueConverter.ArgumentFromString<int>("expectedResult", "10"));
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelName("LPS 800 - 1");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGalleryNotHavingImages("Multi-Tone Sounder", "Conventional Sounders");
            Delay.Milliseconds(0);
            
            Libraries.Gallery_Functions.clickOnGalleryDropDown("Conventional Sounders");
            Delay.Milliseconds(0);
            
            Libraries.Gallery_Functions.GetandVerifyTextofSpecifiedGalleryIndexItem("Conventional Sounders", "1", "Generic Sounder");
            Delay.Milliseconds(0);
            
            Libraries.Gallery_Functions.clickOnGalleryDropDown("Conventional Sounders");
            Delay.Milliseconds(0);
            
            Libraries.Gallery_Functions.GetandVerifyTextofSpecifiedGalleryIndexItem("Conventional Sounders", "2", "Squashni Sounder");
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
