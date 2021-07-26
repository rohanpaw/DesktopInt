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

namespace Fireclass.BelgiumFCRecordingsModule
{
#pragma warning disable 0436 //(CS0436) The type 'type' in 'assembly' conflicts with the imported type 'type2' in 'assembly'. Using the type defined in 'assembly'.
    /// <summary>
    ///The Verify_FC700_Series_Panel_Gallery_Items_For_RBUS_Node_And_Maximum_Number_Of_Devices_Added recording.
    /// </summary>
    [TestModule("7348d116-e574-49bc-b72e-af49652f26c2", ModuleType.Recording, 1)]
    public partial class Verify_FC700_Series_Panel_Gallery_Items_For_RBUS_Node_And_Maximum_Number_Of_Devices_Added : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::Fireclass.FireclassRepository repository.
        /// </summary>
        public static global::Fireclass.FireclassRepository repo = global::Fireclass.FireclassRepository.Instance;

        static Verify_FC700_Series_Panel_Gallery_Items_For_RBUS_Node_And_Maximum_Number_Of_Devices_Added instance = new Verify_FC700_Series_Panel_Gallery_Items_For_RBUS_Node_And_Maximum_Number_Of_Devices_Added();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_FC700_Series_Panel_Gallery_Items_For_RBUS_Node_And_Maximum_Number_Of_Devices_Added()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_FC700_Series_Panel_Gallery_Items_For_RBUS_Node_And_Maximum_Number_Of_Devices_Added Instance
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

            TestProject.Libraries.Panel_Functions.AddPanelsFC(ValueConverter.ArgumentFromString<int>("NumberofPanels", "1"), "FC702S", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Common_Functions.ClickOnNavigationTreeExpander("Node");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Common_Functions.ClickOnNavigationTreeItem("RBus");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Gallery_Functions.verifyGalleryListItems("FC_Gallery_RBUS_Repeaters_FC700_Series_Panels", "Belgium", "FC32AR");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Gallery_Functions.verifyGalleryListItems("FC_Gallery_Miscellaneous_FC_Series_Panels", "Belgium", "FC1D2");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Gallery_Functions.verifyGalleryExistsWithDropdown(ValueConverter.ArgumentFromString<bool>("GalleryVisibility", "False"), "Generic Printer");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("FC32AR", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("FC32AR", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("FC32DR", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("FC32AR", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("FC32DR", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("FC32AR", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("FC32AR", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("MPM800", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("MPM800", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("MPM800", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("MPM800", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("MPM800", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("MPM800", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("MPM800", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("MPM800", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("MPM800", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("Zonal Alarm Display max 80", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Common_Functions.VerifyNavigationTreeItemText("RBus (16 of 16)");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Panel_Functions.DeletePanel(ValueConverter.ArgumentFromString<int>("NumberofPanels", "1"), "Node1", ValueConverter.ArgumentFromString<int>("rowNumber", "1"));
            Delay.Milliseconds(0);
            
            // ANOTHER PANEL TEST CASE
            Report.Log(ReportLevel.Info, "Section", "ANOTHER PANEL TEST CASE", new RecordItemIndex(25));
            
            TestProject.Libraries.Panel_Functions.AddPanelsFC(ValueConverter.ArgumentFromString<int>("NumberofPanels", "1"), "FC718D", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Common_Functions.ClickOnNavigationTreeExpander("Node");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Common_Functions.ClickOnNavigationTreeItem("RBus");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Gallery_Functions.verifyGalleryListItems("FC_Gallery_RBUS_Repeaters_FC700_Series_Panels", "Belgium", "FC32AR");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Gallery_Functions.verifyGalleryListItems("FC_Gallery_Miscellaneous_FC_Series_Panels", "Belgium", "FC1D2");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Gallery_Functions.verifyGalleryExistsWithDropdown(ValueConverter.ArgumentFromString<bool>("GalleryVisibility", "False"), "Generic Printer");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("FC32AR", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("FC32AR", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("FC32DR", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("FC32AR", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("FC32DR", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("FC32AR", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("FC32AR", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("MPM800", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("MPM800", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("MPM800", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("MPM800", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("MPM800", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("MPM800", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("MPM800", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("MPM800", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("MPM800", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Devices_Functions.AddDevicesfromGallery("Zonal Alarm/Fault Display 40", "");
            Delay.Milliseconds(0);
            
            TestProject.Libraries.Common_Functions.VerifyNavigationTreeItemText("RBus (16 of 16)");
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
