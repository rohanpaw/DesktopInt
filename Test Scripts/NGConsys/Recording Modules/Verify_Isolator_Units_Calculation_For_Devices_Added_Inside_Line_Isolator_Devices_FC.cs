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

namespace TestProject.Recording_Modules
{
#pragma warning disable 0436 //(CS0436) The type 'type' in 'assembly' conflicts with the imported type 'type2' in 'assembly'. Using the type defined in 'assembly'.
    /// <summary>
    ///The Verify_Isolator_Units_Calculation_For_Devices_Added_Inside_Line_Isolator_Devices_FC recording.
    /// </summary>
    [TestModule("5dfa8b39-1149-4806-aa49-30f39f6765a2", ModuleType.Recording, 1)]
    public partial class Verify_Isolator_Units_Calculation_For_Devices_Added_Inside_Line_Isolator_Devices_FC : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Verify_Isolator_Units_Calculation_For_Devices_Added_Inside_Line_Isolator_Devices_FC instance = new Verify_Isolator_Units_Calculation_For_Devices_Added_Inside_Line_Isolator_Devices_FC();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_Isolator_Units_Calculation_For_Devices_Added_Inside_Line_Isolator_Devices_FC()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_Isolator_Units_Calculation_For_Devices_Added_Inside_Line_Isolator_Devices_FC Instance
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

            Libraries.Panel_Functions.AddPanelsFC(ValueConverter.ArgumentFromString<int>("NumberofPanels", "1"), "FIRECLASS 64-2", "");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeExpander("Node");
            Delay.Milliseconds(0);
            
            //Libraries.Common_Functions.ClickOnNavigationTreeExpander("FIM");
            //Delay.Milliseconds(0);
            
            Libraries.Common_Functions.ClickOnNavigationTreeItem("Built-in Loop-A");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("460PH", "Detectors");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("460P", "Detectors");
            Delay.Milliseconds(0);
            
            //Libraries.Devices_Functions.AddDevicesfromGallery("460PC", "Detectors");
            //Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("420CP", "Call Points");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("410RIM", "Ancillary");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("410DIM", "Ancillary");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("441AVB", "Detectors");
            Delay.Milliseconds(0);
            
            //Libraries.Devices_Functions.AddDevicesfromGallery("410MIM", "Ancillary");
            //Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("410SIO", "Ancillary");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.AddDevicesfromGallery("421CP", "Call Points");
            Delay.Milliseconds(0);
            
            //Libraries.Devices_Functions.AddDevicesfromGallery("410CIM", "Ancillary");
            //Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.DragAndDropDevicesInPhysicalLayout("A:8+", "A:6+");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.clickOnPointsTab();
            Delay.Milliseconds(0);
            
            //Libraries.IS_Functions.VerifyIsolatorUnits("3.5", "FIM");
            //Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.VerifyValueOfParameterInPhysicalLayout("14", "9.5");
            Delay.Milliseconds(0);
            
            Libraries.Devices_Functions.VerifyValueOfParameterInPhysicalLayout("15", "6.5");
            Delay.Milliseconds(0);
            
            //Libraries.Devices_Functions.VerifyValueOfParameterInPhysicalLayout("9", "5");
            //Delay.Milliseconds(0);
            
            //Libraries.Devices_Functions.VerifyValueOfParameterInPhysicalLayout("10", "1.5");
            //Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
