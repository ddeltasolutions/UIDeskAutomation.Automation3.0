using System;
using System.Diagnostics;
using System.Collections.Generic;
using UIAutomationClient;
using System.Threading;

namespace UIDeskAutomationLib
{
	/// <summary>
	/// Class used for automating Google Chrome browser
	/// </summary>
	public class ChromeBrowser: UIDA_Window
	{
		private const int MAX_RETRIES_COUNT = 100;
		
		private IUIAutomationElement document = null;
		
		/// <summary>
        /// Creates a new Google Chrome browser automation object
        /// </summary>
		internal ChromeBrowser(): base(null)
		{
			
		}
	
		/// <summary>
        /// Starts a browser instance/window
        /// </summary>
		internal void StartBrowser()
		{
			try
			{
				Process proc = Process.Start("chrome.exe");
			}
			catch {}
			
			for (int i = 0; i < 10; i++)
			{
				if (FindDocument() == true)
				{
					break;
				}
			}
		}
		
		private bool FindDocument()
		{
			try
			{
				IntPtr hWndCurrent = IntPtr.Zero;
				int retries = 0;
				Thread.Sleep(1000);
				
				while (retries < MAX_RETRIES_COUNT)
				{
					hWndCurrent = UnsafeNativeFunctions.FindWindow("Chrome_WidgetWin_1", null);
					if (hWndCurrent != IntPtr.Zero)
					{
						break;
					}
					
					Thread.Sleep(100);
					retries++;
				}
				
				if (hWndCurrent != null)
				{
					base.hWnd = hWndCurrent;
					base.uiElement = Engine.uiAutomation.ElementFromHandle(hWndCurrent);
					
					retries = 0;
					while (retries <= 10)
					{
						IUIAutomationCondition condition = Engine.uiAutomation.CreatePropertyCondition(
							UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_DocumentControlTypeId);
						document = uiElement.FindFirst(TreeScope.TreeScope_Children, condition);
						if (document != null)
						{
							break;
						}
						Thread.Sleep(100);
						retries++;
					}
				}
				else
				{
					Engine.TraceInLogFile("Browser window not found");
					return false;
				}
			}
			catch (Exception ex)
			{
				Engine.TraceInLogFile(ex.Message);
				return false;
			}
			
			return true;
		}
		
		private IUIAutomationElement addressBar = null;
		
		/// <summary>
        /// Navigates to a specified url
        /// </summary>
		/// <param name="url">url to navigate to</param>
		public void Navigate(string url)
		{
			Array oldRuntimeId = null;
			if (document != null)
			{
				try
				{
					oldRuntimeId = document.GetRuntimeId();
				}
				catch {}
			}
		
			if (addressBar == null)
			{
				for (int i = 0; i < MAX_RETRIES_COUNT; i++)
				{
					IUIAutomationCondition condition1 = Engine.uiAutomation.CreatePropertyCondition(UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_EditControlTypeId);
					IUIAutomationCondition condition2 = Engine.uiAutomation.CreatePropertyCondition(UIA_PropertyIds.UIA_AccessKeyPropertyId, "Ctrl+L");					
					IUIAutomationCondition condition = Engine.uiAutomation.CreateAndCondition(condition1, condition2);
				
					addressBar = uiElement.FindFirst(TreeScope.TreeScope_Descendants, condition);
					if (addressBar != null)
					{
						break;
					}
					Thread.Sleep(200);
				}
				
				if (addressBar == null)
				{
					Engine.TraceInLogFile("address bar not found");
					return;
				}
			}
			
			/*try
			{
				addressBar.SetFocus();
			}
			catch (Exception ex)
			{
				Engine.TraceInLogFile("Exception: " + ex.Message);
			}*/
			
			UIDA_Edit edit = new UIDA_Edit(addressBar);
			edit.SelectAll();
			edit.SendKeys(url);
			edit.KeyPress(VirtualKeys.Enter);
			
			Thread.Sleep(500);
			int retries = 0;
			while (retries < 20)
			{
				IUIAutomationCondition cond = Engine.uiAutomation.CreatePropertyCondition(UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_DocumentControlTypeId);
				IUIAutomationElement newDocument = uiElement.FindFirst(TreeScope.TreeScope_Descendants, cond);
				
				if (newDocument != null)
				{
					if (oldRuntimeId == null || Engine.uiAutomation.CompareRuntimeIds(newDocument.GetRuntimeId(), oldRuntimeId) == 0)
					{
						document = newDocument;
						break;
					}
				}
				
				Thread.Sleep(500);
				retries++;
			}
		}
		
		/// <summary>
        /// Press the browser Back button
        /// </summary>
		public void Back()
		{
			UIDA_ToolBar toolBar = this.ToolBar(null, true);
			if (toolBar == null)
			{
				return;
			}
			UIDA_Button[] buttons = toolBar.Buttons(null, true);
			if (buttons.Length == 0)
			{
				return;
			}
			
			UIDA_Button backButton = buttons[0];
			backButton.Invoke();
		}
		
		/// <summary>
        /// Press the browser Forward button
        /// </summary>
		public void Forward()
		{
			UIDA_ToolBar toolBar = this.ToolBar(null, true);
			if (toolBar == null)
			{
				return;
			}
			UIDA_Button[] buttons = toolBar.Buttons(null, true);
			if (buttons.Length < 2)
			{
				return;
			}
		
			UIDA_Button forwardButton = buttons[1];
			forwardButton.Invoke();
		}
		
		/*private static string GetDocumentValue(IUIAutomationElement document)
		{
			IUIAutomationValuePattern valuePattern = document.GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId) as IUIAutomationValuePattern;
			string currentValue = null;
			
			if (valuePattern != null)
			{
				try
				{
					currentValue = valuePattern.CurrentValue;
				}
				catch {}
			}
			return currentValue;
		}*/
	}
}