/*
 * Created by Ranorex
 * User: jbhosash
 * Date: 11/2/2017
 * Time: 11:57 AM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Threading;
using WinForms = System.Windows.Forms;
using System.Diagnostics;

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace TestProject
{
	/// <summary>
	/// Description of App_Open.
	/// </summary>
	[TestModule("F25381C6-DB7E-4BC9-A0C1-478AE17EEF2D", ModuleType.UserCode, 1)]
	public class Application_Open : ITestModule
	{
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public Application_Open()
		{
			// Do not delete - a parameterless constructor is required!
		}

		int _App_Mode;
		[TestVariable("172fdd8c-61ee-4b80-8c82-2670b96cc40b")]
		public int App_Mode
		{
			get { return _App_Mode; }
			set { _App_Mode = value; }
		}

		string _App_Path;
		[TestVariable("1124e7b3-695e-4731-8f79-d1e83c8943fc")]
		public string App_Path
		{
			get { return _App_Path; }
			set { _App_Path = value; }
		}



		/// <summary>
		/// Performs the playback of actions in this module.
		/// </summary>
		/// <remarks>You should not call this method directly, instead pass the module
		/// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
		/// that will in turn invoke this method.</remarks>
		void ITestModule.Run()
		{
			Mouse.DefaultMoveTime = 300;
			Keyboard.DefaultKeyPressTime = 100;
			Delay.SpeedFactor = 1.0;
			
			//check if app process is running and kill 
			Process[] processes = null; 
			processes = Process.GetProcessesByName("ProfileConsys");
			foreach (Process proces in processes) 
			{ 
				proces.Kill(); 
			}
			
			
			//check if app process is hanged and kill 
			Process[] processes1 = null; 
			processes1 = Process.GetProcessesByName("WerFault.exe");
			foreach (Process proces in processes1) 
			{ 
				proces.Kill(); 
			}
			
			
			
			Host.Local.RunApplication("C:\\Windows\\System32\\cmd.exe", "", "", false);
			Delay.Milliseconds(500);
			
			Delay.Milliseconds(500);
			Keyboard.Press("cd ");
			Keyboard.Press(App_Path);
			Keyboard.Press("{Return}");
			if(App_Mode==0)
			{
				Keyboard.Press("ProfileConsys.exe");
			}
			else if(App_Mode==1)
			{
				Keyboard.Press("ProfileConsys.exe -mode:Consys");
			}
			else
			{
				Keyboard.Press("ProfileConsys.exe -mode:fast");
			}
			
			Keyboard.Press("{Return}");
			Keyboard.Press("exit");
			Keyboard.Press("{Return}");
			
			
		}
	}
}
