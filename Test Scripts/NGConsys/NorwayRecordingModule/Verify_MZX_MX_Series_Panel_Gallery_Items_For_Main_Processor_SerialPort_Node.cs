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

namespace TestProject.NorwayRecordingModule
{
#pragma warning disable 0436 //(CS0436) The type 'type' in 'assembly' conflicts with the imported type 'type2' in 'assembly'. Using the type defined in 'assembly'.
    /// <summary>
    ///The Verify_MZX_MX_Series_Panel_Gallery_Items_For_Main_Processor_SerialPort_Node recording.
    /// </summary>
    [TestModule("30b677bb-2937-42ff-abc1-6c25221aa4bf", ModuleType.Recording, 1)]
    public partial class Verify_MZX_MX_Series_Panel_Gallery_Items_For_Main_Processor_SerialPort_Node : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Verify_MZX_MX_Series_Panel_Gallery_Items_For_Main_Processor_SerialPort_Node instance = new Verify_MZX_MX_Series_Panel_Gallery_Items_For_Main_Processor_SerialPort_Node();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_MZX_MX_Series_Panel_Gallery_Items_For_Main_Processor_SerialPort_Node()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_MZX_MX_Series_Panel_Gallery_Items_For_Main_Processor_SerialPort_Node Instance
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

            Libraries.Panel_Functions.AddPanels(ValueConverter.ArgumentFromString<int>("NumberofPanels", "1"), "MZX254S", "");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeExpander("Node");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Main");
            Delay.Milliseconds(0);
            
            // Repeaters
            Libraries.Gallery_Functions.verifyGalleryListItems("Gallery_MainProcessor_Repeaters_MZX_Panels", "Norway", "MXR");
            Delay.Milliseconds(0);
            
            // Miscellaneous
            Libraries.Gallery_Functions.verifyGalleryListItems("Gallery_Miscellaneous_MZX_Panels", "Norway", "PR1D2");
            Delay.Milliseconds(0);
            
            // Printers
            Libraries.Gallery_Functions.verifyGalleryListItems("Gallery_Printers_MZX_Panels", "Norway", "LCD800");
            Delay.Milliseconds(0);
            
            Libraries.Gallery_Functions.verifyGalleryExistsWithDropdown(ValueConverter.ArgumentFromString<bool>("GalleryVisibility", "False"), "IOB800(x1)");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Serial");
            Delay.Milliseconds(0);
            
            Libraries.Gallery_Functions.verifyGalleryExistsWithDropdown(ValueConverter.ArgumentFromString<bool>("GalleryVisibility", "False"), "XLM800");
            Delay.Milliseconds(0);
            
            // Printers
            Libraries.Gallery_Functions.verifyGalleryListItems("Gallery_Printers_MZX_Panels", "Norway", "LCD800");
            Delay.Milliseconds(0);
            
            Libraries.Gallery_Functions.verifyGalleryExistsWithDropdown(ValueConverter.ArgumentFromString<bool>("GalleryVisibility", "False"), "FBI800");
            Delay.Milliseconds(0);
            
            Libraries.Gallery_Functions.verifyGalleryExistsWithDropdown(ValueConverter.ArgumentFromString<bool>("GalleryVisibility", "False"), "PR1D2");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingLabelNameFromInventoryTab("Printer 1");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("LCD800", "Printers");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.clickOnInventoryTab();
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.SelectRowUsingModelNameFromInventoryTab("LCD800");
            Delay.Milliseconds(0);
            
            //Libraries.Gallery_Functions.verifyGalleryExistsWithDropdown(ValueConverter.ArgumentFromString<bool>("GalleryVisibility", "False"), "LCD800");
            //Delay.Milliseconds(0);
            
            //Libraries.Devices_Functions.SelectRowUsingLabelNameFromInventoryTab("Printer 2");
            //Delay.Milliseconds(0);
            
            //Libraries.Devices_Functions.SelectRowUsingSKUFromInventoryTab("557.202.090");
            //Delay.Milliseconds(0);
            
            //Libraries.Devices_Functions.SelectRowUsingConnectionFromInventoryTab("Serial Port");
            //Delay.Milliseconds(0);
            
            //Libraries.Common_Functions.clickOnInventoryTab();
            //Delay.Milliseconds(0);
            
            //Libraries.Panel_Functions.DeletePanel(ValueConverter.ArgumentFromString<int>("NumberofPanels", "1"), "Node1", ValueConverter.ArgumentFromString<int>("rowNumber", "1"));
            //Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
