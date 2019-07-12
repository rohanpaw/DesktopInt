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
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ProfileConsys1.NavigationTree.Expander' at Center.", repo.ProfileConsys1.NavigationTree.ExpanderInfo, new RecordItemIndex(1));
            repo.ProfileConsys1.NavigationTree.Expander.Click();
            Delay.Milliseconds(200);
            
            Libraries.Devices_Functions.AddDevicesfromPanelNodeGallery("PLX800", "Loops", "");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.BackplaneOrXLMExternalLoopCard_Expander' at Center.", repo.FormMe.BackplaneOrXLMExternalLoopCard_ExpanderInfo, new RecordItemIndex(3));
            repo.FormMe.BackplaneOrXLMExternalLoopCard_Expander.Click();
            Delay.Milliseconds(200);
            
            //Libraries.Devices_Functions.verifyDescription("Backplane Assembly with 6 slots for use with PxD");
            //Delay.Milliseconds(0);
            
            //Libraries.Devices_Functions.verifyPicture();
            //Delay.Milliseconds(0);
            
            //Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.cell_Properties_backplane' at Center.", repo.FormMe.cell_Properties_backplaneInfo, new RecordItemIndex(6));
            //repo.FormMe.cell_Properties_backplane.Click();
            //Delay.Milliseconds(200);
            
            //Report.Log(ReportLevel.Info, "Validation", "Validating ContainsImage (Screenshot: 'Screenshot1' with region {X=0,Y=0,Width=313,Height=735}) on item 'FormMe.BackplaneImage'.", repo.FormMe.BackplaneImageInfo, new RecordItemIndex(7));
            //Validate.ContainsImage(repo.FormMe.BackplaneImageInfo, BackplaneImage_Screenshot1, BackplaneImage_Screenshot1_Options);
            //Delay.Milliseconds(0);
            
            //Libraries.Devices_Functions.ErasePictureFromProperties();
            //Delay.Milliseconds(0);
            
            // Have to verify the value for backplane SKU
            //VerifyProductCodeInSearchProperties("123");
            //Delay.Milliseconds(0);
            
            //Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'FormMe.Inventory_Tab_Backplane' at Center.", repo.FormMe.Inventory_Tab_BackplaneInfo, new RecordItemIndex(10));
            //repo.FormMe.Inventory_Tab_Backplane.Click();
            //Delay.Milliseconds(200);
            
            //Libraries.InventoryGrid_Functions.verifyInventoryGridProperties("1", "SKU", "123.456.789");
            //Delay.Milliseconds(0);
            
            //Libraries.InventoryGrid_Functions.verifyInventoryGridProperties("1", "Model", "PLX800");
            //Delay.Milliseconds(0);
            
            //Libraries.InventoryGrid_Functions.verifyInventoryGridProperties("1", "Label", "PLX800-E");
            //Delay.Milliseconds(0);
            
            // Slot address wont be displayed in grid as told
            //Libraries.InventoryGrid_Functions.verifyInventoryGridProperties("1", "Address", "E");
            //Delay.Milliseconds(0);
            
            //Libraries.InventoryGrid_Functions.verifyInventoryGridProperties("1", "Connection", "PLX/External Loop Card 2");
            //Delay.Milliseconds(0);
            
            //Libraries.InventoryGrid_Functions.EditDevicePropertyWhichAreReadOnly("1", "SKU", "123");
            //Delay.Milliseconds(0);
            
            //Libraries.InventoryGrid_Functions.EditDevicePropertyWhichAreReadOnly("1", "Address", "12");
            //Delay.Milliseconds(0);
            
            //Libraries.InventoryGrid_Functions.EditDevicePropertyWhichAreReadOnly("1", "Connection", "abc");
            //Delay.Milliseconds(0);
            
            //Libraries.InventoryGrid_Functions.EditDevicePropertyWhichAreReadOnly("1", "Model", "pqr");
            //Delay.Milliseconds(0);
            
            Libraries.InventoryGrid_Functions.verifyInventoryGridProperties("1", "Label", "PLX800-E");
            Delay.Milliseconds(0);
            
            VerifyProductCodeInSearchProperties("123.456.789");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Validation", "Validating ContainsImage (Screenshot: 'Screenshot1' with region {X=0,Y=0,Width=380,Height=380}) on item 'FormMe.PLX_Image'.", repo.FormMe.PLX_ImageInfo, new RecordItemIndex(22));
            Validate.ContainsImage(repo.FormMe.PLX_ImageInfo, PLX_Image_Screenshot1, PLX_Image_Screenshot1_Options);
            Delay.Milliseconds(0);
            
            try {
                //Libraries.Devices_Functions.verifyDescription("(Slot card) Provides MX Loop interfaces for connection fire detectors and ancillaries. Card is connected to panel using Internal N-Bus interface");
                //Delay.Milliseconds(0);
            } catch(Exception ex) { Report.Log(ReportLevel.Warn, "Module", "(Optional Action) " + ex.Message, new RecordItemIndex(23)); }
            
            VerifyRegionNameInSearchProperties("PLX800-E");
            Delay.Milliseconds(0);
            
            //VerifyProductInSearchProperties("PLX800");
            //Delay.Milliseconds(0);
            
            Libraries.InventoryGrid_Functions.editDeviceLabel("1", "Label", "PLX800");
            Delay.Milliseconds(0);
            
            VerifyRegionNameInSearchProperties("PLX800");
            Delay.Milliseconds(0);
            
            editRegionName("PLX");
            Delay.Milliseconds(0);
            
            Libraries.InventoryGrid_Functions.verifyInventoryGridProperties("1", "Label", "PLX");
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
        CompressedImage BackplaneImage_Screenshot1
        { get { return repo.FormMe.BackplaneImageInfo.GetScreenshot1(new Rectangle(0, 0, 313, 735)); } }

        Imaging.FindOptions BackplaneImage_Screenshot1_Options
        { get { return Imaging.FindOptions.Default; } }

        CompressedImage PLX_Image_Screenshot1
        { get { return repo.FormMe.PLX_ImageInfo.GetScreenshot1(new Rectangle(0, 0, 380, 380)); } }

        Imaging.FindOptions PLX_Image_Screenshot1_Options
        { get { return Imaging.FindOptions.Default; } }

#endregion
    }
#pragma warning restore 0436
}