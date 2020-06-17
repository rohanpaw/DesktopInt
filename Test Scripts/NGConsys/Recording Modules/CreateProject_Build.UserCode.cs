﻿///////////////////////////////////////////////////////////////////////////////
//
// This file was automatically generated by RANOREX.
// Your custom recording code should go in this file.
// The designer will only add methods to this file, so your custom code won't be overwritten.
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
using Ranorex.Core.Repository;
using Ranorex.Core.Testing;

namespace TestProject.Recording_Modules
{
    public partial class CreateProject_Build
    {
        /// <summary>
        /// This method gets called right after the recording has been started.
        /// It can be used to execute recording specific initialization code.
        /// </summary>
        private void Init()
        {
            // Your recording specific initialization code goes here.
        }

        public void Select_InstallerAddress(string sInstallerAddress)
        {
        	repo.CreateNewProject.CreateNewProjectContainer.txt_InstallerAddress.Click();
            Keyboard.Press(sInstallerAddress);
        }

        public void Select_InstallerName(string sInstallerName)
        {
         	repo.CreateNewProject.CreateNewProjectContainer.txt_InstallerName.Click();
            Keyboard.Press(sInstallerName);
        }

        public void Select_ClientAddress(string sClientAddress)
        {
           	repo.CreateNewProject.CreateNewProjectContainer.txt_ClientAddress.Click();           	
            Keyboard.Press(sClientAddress);
        }

        public void Select_ClientName(string sClientName)
        {
          	repo.CreateNewProject.CreateNewProjectContainer.txt_ClientName.Click();          	
            Keyboard.Press(sClientName);
        }

        public void Select_ProjectName(string sProjectName)
        {
        	repo.CreateNewProject.CreateNewProjectContainer.txt_ProjectName.Click();
        	//ProjectName=sProjectName;
        	Delay.Duration(1000, false);
            Keyboard.Press(sProjectName);
        }

        public void ListItem(int iListIndex)
        {
        	sListIndex=iListIndex.ToString();
        	repo.ProfileConsys1.lst_Market.Click();
        }

        public void Select_Market(string sMarket)
        {
        	
            Keyboard.Press(sMarket);
        }

    }
}
