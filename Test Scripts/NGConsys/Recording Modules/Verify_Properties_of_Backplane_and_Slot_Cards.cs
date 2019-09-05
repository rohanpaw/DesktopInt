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
    ///The Verify_Properties_of_Backplane_and_Slot_Cards recording.
    /// </summary>
    [TestModule("f8b16155-ecb9-47df-a1d6-27bca78454e5", ModuleType.Recording, 1)]
    public partial class Verify_Properties_of_Backplane_and_Slot_Cards : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Verify_Properties_of_Backplane_and_Slot_Cards instance = new Verify_Properties_of_Backplane_and_Slot_Cards();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_Properties_of_Backplane_and_Slot_Cards()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_Properties_of_Backplane_and_Slot_Cards Instance
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
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Node");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromPanelNodeGallery("PLX800", "Loops", "");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Backplane");
            Delay.Milliseconds(0);
            
            Libraries.InventoryGrid_Functions.verifyInventoryGridProperties("1", "Label", "PLX800-E");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.VerifyProductCodeInSearchProperties("557.202.842");
            Delay.Milliseconds(0);
            
            //Report.Log(ReportLevel.Info, "Validation", "Validating ContainsImage (Screenshot: 'Screenshot1' with region {X=0,Y=0,Width=380,Height=380}) on item 'FormMe.PLX_Image'.", repo.FormMe.PLX_ImageInfo, new RecordItemIndex(7));
            //Validate.ContainsImage(repo.FormMe.PLX_ImageInfo, PLX_Image_Screenshot1, PLX_Image_Screenshot1_Options);
            //Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.verifyDescription("(Slot card) Provides MX Loop interfaces for connection fire detectors and ancillaries. Card is connected to panel using Internal N-Bus interface");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.VerifyRegionNameInSearchProperties("PLX800-E");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.VerifyProductInSearchProperties("PLX800");
            Delay.Milliseconds(0);
            
            Libraries.InventoryGrid_Functions.editDeviceLabel("1", "Label", "PLX800");
            Delay.Milliseconds(0);
            
            VerifyRegionNameInSearchProperties("PLX800");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.editRegionName("PLX");
            Delay.Milliseconds(0);
            
            Libraries.InventoryGrid_Functions.verifyInventoryGridProperties("1", "Label", "PLX");
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
        CompressedImage PLX_Image_Screenshot1
        { get { return repo.FormMe.PLX_ImageInfo.GetScreenshot1(new Rectangle(0, 0, 380, 380)); } }

        Imaging.FindOptions PLX_Image_Screenshot1_Options
        { get { return Imaging.FindOptions.Default; } }

#endregion
    }
#pragma warning restore 0436
}
