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
    ///The Verify_Create_New_Project recording.
    /// </summary>
    [TestModule("c65c37e7-54e2-4fe9-a1ab-af6120b0cb7b", ModuleType.Recording, 1)]
    public partial class Verify_Create_New_Project : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Verify_Create_New_Project instance = new Verify_Create_New_Project();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Verify_Create_New_Project()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Verify_Create_New_Project Instance
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

            Report.Log(ReportLevel.Info, "Validation", "Validating Exists on item 'ProfileConsys1.SiteNode'.", repo.ProfileConsys1.SiteNodeInfo, new RecordItemIndex(0));
            Validate.Exists(repo.ProfileConsys1.SiteNodeInfo);
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Validation", "Validating AttributeContains (Text>'Panels') on item 'HwndWrapperProfileConsysExe0c643c73.PanelsGallery'.", repo.HwndWrapperProfileConsysExe0c643c73.PanelsGalleryInfo, new RecordItemIndex(1));
            Validate.AttributeContains(repo.HwndWrapperProfileConsysExe0c643c73.PanelsGalleryInfo, "Text", "Panels");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Validation", "Validating Exists on item 'HwndWrapperProfileConsysExe0c643c73.OtherNodesGallery'.", repo.HwndWrapperProfileConsysExe0c643c73.OtherNodesGalleryInfo, new RecordItemIndex(2));
            Validate.Exists(repo.HwndWrapperProfileConsysExe0c643c73.OtherNodesGalleryInfo);
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Validation", "Validating Exists on item 'HwndWrapperProfileConsysExe0c643c73.tab_SiteAccessories'.", repo.HwndWrapperProfileConsysExe0c643c73.tab_SiteAccessoriesInfo, new RecordItemIndex(3));
            Validate.Exists(repo.HwndWrapperProfileConsysExe0c643c73.tab_SiteAccessoriesInfo);
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
