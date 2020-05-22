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
    ///The CreateProject_German recording.
    /// </summary>
    [TestModule("19660af1-4e58-4290-badb-20687f7b3401", ModuleType.Recording, 1)]
    public partial class CreateProject_German : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static CreateProject_German instance = new CreateProject_German();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public CreateProject_German()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static CreateProject_German Instance
        {
            get { return instance; }
        }

#region Variables

        /// <summary>
        /// Gets or sets the value of variable sListIndex.
        /// </summary>
        [TestVariable("ca40602d-7ca3-47bc-a519-7a0fe6a76634")]
        public string sListIndex
        {
            get { return repo.sListIndex; }
            set { repo.sListIndex = value; }
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

            Report.Log(ReportLevel.Info, "Delay", "Waiting for 1s.", new RecordItemIndex(0));
            Delay.Duration(1000, false);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ProfileConsys1.File' at Center.", repo.ProfileConsys1.FileInfo, new RecordItemIndex(1));
            repo.ProfileConsys1.File.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ProfileConsys1.TextNew' at Center.", repo.ProfileConsys1.TextNewInfo, new RecordItemIndex(2));
            repo.ProfileConsys1.TextNew.Click();
            Delay.Milliseconds(200);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ProfileConsys1.PARTRight.txt_MarketNew' at Center.", repo.ProfileConsys1.PARTRight.txt_MarketNewInfo, new RecordItemIndex(3));
            repo.ProfileConsys1.PARTRight.txt_MarketNew.Click();
            Delay.Milliseconds(200);
            
            Select_Market("Germa");
            Delay.Milliseconds(0);
            
            ListItem(ValueConverter.ArgumentFromString<int>("iListIndex", "3"));
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'ProfileConsys1.PARTRight.btn_CreateNewProject' at Center.", repo.ProfileConsys1.PARTRight.btn_CreateNewProjectInfo, new RecordItemIndex(6));
            repo.ProfileConsys1.PARTRight.btn_CreateNewProject.Click();
            Delay.Milliseconds(200);
            
            //Select_ProjectName("Verify");
            //Delay.Milliseconds(0);
            
            //Select_ClientName("JCI");
            //Delay.Milliseconds(0);
            
            //Select_ClientAddress("JCI");
            //Delay.Milliseconds(0);
            
            //Select_InstallerName("JCI");
            //Delay.Milliseconds(0);
            
            //Select_InstallerAddress("JCI");
            //Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Delay", "Waiting for 5s.", new RecordItemIndex(12));
            Delay.Duration(5000, false);
            
            Libraries.Common_Functions.CreateProjectFCParameters("9", "JCI");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.CreateProjectFCParameters("10", "JCI");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.CreateProjectFCParameters("11", "JCI");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.CreateProjectFCParameters("12", "JCI");
            Delay.Milliseconds(0);
            
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Left Click item 'CreateNewProject.CreateNewProjectContainer.btn_OK' at Center.", repo.CreateNewProject.CreateNewProjectContainer.btn_OKInfo, new RecordItemIndex(17));
            repo.CreateNewProject.CreateNewProjectContainer.btn_OK.Click();
            Delay.Milliseconds(200);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
