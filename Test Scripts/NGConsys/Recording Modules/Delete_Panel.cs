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
    ///The Delete_Panel recording.
    /// </summary>
    [TestModule("8be0600b-c1ac-42b1-a20d-f61faad8c119", ModuleType.Recording, 1)]
    public partial class Delete_Panel : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::TestProject.NGConsysRepository repository.
        /// </summary>
        public static global::TestProject.NGConsysRepository repo = global::TestProject.NGConsysRepository.Instance;

        static Delete_Panel instance = new Delete_Panel();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public Delete_Panel()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static Delete_Panel Instance
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

            Libraries.Panel_Functions.AddPanels(ValueConverter.ArgumentFromString<int>("NumberofPanels", "3"), "Pro32xD,P485D,MX1000", "");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.VerifyElementVisibilityInNavigationTree(ValueConverter.ArgumentFromString<bool>("sExists", "True"), "Node1");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.VerifyElementVisibilityInNavigationTree(ValueConverter.ArgumentFromString<bool>("sExists", "True"), "Node2");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.VerifyElementVisibilityInNavigationTree(ValueConverter.ArgumentFromString<bool>("sExists", "True"), "Node3");
            Delay.Milliseconds(0);
            
            Libraries.Panel_Functions.DeletePanel(ValueConverter.ArgumentFromString<int>("NumberofPanels", "3"), "Node1", ValueConverter.ArgumentFromString<int>("rowNumber", "1"));
            Delay.Milliseconds(0);
            
            //Libraries.Panel_Functions.DeletePanel("Node1", ValueConverter.ArgumentFromString<int>("rowNumber", "1"));
            //Delay.Milliseconds(0);
            
            //Libraries.Panel_Functions.DeletePanel("Node2", ValueConverter.ArgumentFromString<int>("rowNumber", "2"));
            //Delay.Milliseconds(0);
            
            //Libraries.Panel_Functions.DeleteSinglePanel("Node3", ValueConverter.ArgumentFromString<int>("rowNumber", "3"));
            //Delay.Milliseconds(0);
            
            Libraries.Common_Functions.VerifyElementVisibilityInNavigationTree(ValueConverter.ArgumentFromString<bool>("sExists", "False"), "Node2");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.VerifyElementVisibilityInNavigationTree(ValueConverter.ArgumentFromString<bool>("sExists", "False"), "Node3");
            Delay.Milliseconds(0);
            
            Libraries.Common_Functions.Application_Close(ValueConverter.ArgumentFromString<bool>("Save", "False"), ValueConverter.ArgumentFromString<bool>("SaveConfirmation", "True"), "");
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
