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
    ///The Verify_impact_of_character_based_repeater_on_loops recording.
    /// </summary>
    [TestModule("b49d100e-b1ad-4cf1-b81e-76e671dad72c", ModuleType.Recording, 1)]
    public partial class Verify_impact_of_character_based_repeater_on_loops : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Verify_impact_of_character_based_repeater_on_loops instance = new Verify_impact_of_character_based_repeater_on_loops();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_impact_of_character_based_repeater_on_loops()
        {
            sRowIndex = "";
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_impact_of_character_based_repeater_on_loops Instance
        {
            get { return instance; }
        }

#region Variables

        /// <summary>
        /// Gets or sets the value of variable sRowIndex.
        /// </summary>
        [TestVariable("bed0e8e5-b2ac-4236-9869-81955090d441")]
        public string sRowIndex
        {
            get { return repo.sRowIndex; }
            set { repo.sRowIndex = value; }
        }

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

            Libraries.Panel_Functions.AddPanels(ValueConverter.ArgumentFromString<int>("NumberofPanels", "1"), "Pro32xD", "");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Node");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("PLX800", "Loops");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("PLX800", "Loops");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Delay", "Waiting for 10s.", new RecordItemIndex(4));
            Delay.Duration(10000, false);
            
            Libraries.InventoryGrid_Functions.VerifyInventoryGrid(ValueConverter.ArgumentFromString<int>("iStartRowIndex", "7"), ValueConverter.ArgumentFromString<int>("iEndRowIndex", "14"), "557.202.842");
            Delay.Milliseconds(100);
            
            Libraries.Devices_Functions.VerifyGalleryItem("Repeaters", "MX2 Repeater", "Disabled");
            Delay.Milliseconds(100);
            
            Libraries.InventoryGrid_Functions.DeleteItemfromInventory(ValueConverter.ArgumentFromString<int>("iRowNumber", "8"), "PLX800", "557.202.842");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Delay", "Waiting for 10s.", new RecordItemIndex(8));
            Delay.Duration(10000, false);
            
            Libraries.InventoryGrid_Functions.VerifyRowNotExist(ValueConverter.ArgumentFromString<int>("iRowNumber", "11"), "PLX800", "557.202.842");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.VerifyGalleryItem("Repeaters", "MX2 Repeater", "Enabled");
            Delay.Milliseconds(100);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("MX2 Repeater", "Repeaters");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Site");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Node");
            Delay.Milliseconds(0);
            
            Libraries.InventoryGrid_Functions.VerifyRowExist(ValueConverter.ArgumentFromString<int>("iRowNumber", "11"), "MX2 Repeater", "557.200.206");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.VerifyGalleryItem("Repeaters", "MX2 Repeater", "Disabled");
            Delay.Milliseconds(100);
            
            Libraries.Common_Functions.Application_Close(ValueConverter.ArgumentFromString<bool>("Save", "True"), ValueConverter.ArgumentFromString<bool>("SaveConfirmation", "True"), "NGC-473");
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
