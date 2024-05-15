using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a CheckBox UI element
    /// </summary>
    public class UIDA_CheckBox : ElementBase
    {
        /// <summary>
        /// Creates a UIDA_CheckBox using an IUIAutomationElement
        /// </summary>
        /// <param name="el">UI Automation Element</param>
        public UIDA_CheckBox(IUIAutomationElement el)
        {
            base.uiElement = el;
        }
        
        /// <summary>
        /// Toggles between the states of a checkbox
        /// </summary>
        public void Toggle()
        {
            if (this.IsAlive == false)
            {
                throw new Exception("This UI element is not available to the user anymore.");
            }
			
			if (this.uiElement.CurrentFrameworkId == "WPF")
			{
				// for WPF use mouse click because TogglePattern.Toggle() does not call the checkbox's Click event handler
				this.Click();
				Engine.GetInstance().Sleep(100);
				return;
			}

            object togglePatternObject = uiElement.GetCurrentPattern(UIA_PatternIds.UIA_TogglePatternId);
            IUIAutomationTogglePattern togglePattern = togglePatternObject as IUIAutomationTogglePattern;
            if (togglePattern != null)
            {
                togglePattern.Toggle();
            }
            else
            {
                throw new Exception("Cannot toggle checkbox");
            }
        }

        /// <summary>
        /// Checks a checkbox
        /// </summary>
        public void Check()
        {
            if (this.IsAlive == false)
            {
                throw new Exception("This UI element is not available to the user anymore.");
            }

            object togglePatternObject = uiElement.GetCurrentPattern(UIA_PatternIds.UIA_TogglePatternId);
            IUIAutomationTogglePattern togglePattern = togglePatternObject as IUIAutomationTogglePattern;
            if (togglePattern != null)
            {
                /*int count = 0;
                while (togglePattern.CurrentToggleState != ToggleState.ToggleState_On)
                {
                    if (count >= 3)
                    {
                        throw new Exception("Cannot check the checkbox");
                    }
                    togglePattern.Toggle();
                    count++;
                }

                return;*/
				
				if (togglePattern.CurrentToggleState != ToggleState.ToggleState_On)
				{
					if (this.uiElement.CurrentFrameworkId == "WPF")
					{
						this.Click();
						Engine.GetInstance().Sleep(100);
					}
					else
					{
						togglePattern.Toggle();
					}
				}
				
				if (togglePattern.CurrentToggleState != ToggleState.ToggleState_On)
				{
					if (this.uiElement.CurrentFrameworkId == "WPF")
					{
						this.Click();
						Engine.GetInstance().Sleep(100);
					}
					else
					{
						togglePattern.Toggle();
					}
				}
				
				if (togglePattern.CurrentToggleState == ToggleState.ToggleState_On)
				{
					return;
				}
            }

            // current element does not support Toggle Pattern
            // todo: simulate click
            IntPtr hwnd = IntPtr.Zero;

            try
            {
                hwnd = this.uiElement.CurrentNativeWindowHandle;
            }
            catch
            { }

            bool isWin32Button = false;

            if (hwnd != IntPtr.Zero)
            {
                StringBuilder className = new StringBuilder(256);
                UnsafeNativeFunctions.GetClassName(hwnd, className, 256);

                if (className.ToString() == "Button")
                {
                    isWin32Button = true;

                    // common Win32 checkbox window
                    IntPtr result = UnsafeNativeFunctions.SendMessage(hwnd,
                        ButtonMessages.BM_GETCHECK, IntPtr.Zero, IntPtr.Zero);

                    if (result.ToInt32() != (int)ButtonMessages.BST_CHECKED)
                    {
                        UnsafeNativeFunctions.SendMessage(hwnd, ButtonMessages.BM_SETCHECK,
                            new IntPtr(ButtonMessages.BST_CHECKED), IntPtr.Zero);
                    }
                }
            }

            if (!isWin32Button)
            {

            }
        }

        /// <summary>
        /// Unchecks a checkbox
        /// </summary>
        public void Uncheck()
        {
            if (this.IsAlive == false)
            {
                throw new Exception("This UI element is not available to the user anymore.");
            }

            object togglePatternObject = uiElement.GetCurrentPattern(UIA_PatternIds.UIA_TogglePatternId);
            IUIAutomationTogglePattern togglePattern = togglePatternObject as IUIAutomationTogglePattern;
            if (togglePattern != null)
            {
                //object supports Toggle pattern

                /*int count = 0;
                while (togglePattern.CurrentToggleState != ToggleState.ToggleState_Off)
                {
                    if (count >= 3)
                    {
                        throw new Exception("Cannot uncheck the checkbox");
                    }
                    togglePattern.Toggle();
                    count++;
                }*/
				
				if (togglePattern.CurrentToggleState != ToggleState.ToggleState_Off)
				{
					if (this.uiElement.CurrentFrameworkId == "WPF")
					{
						this.Click();
						Engine.GetInstance().Sleep(100);
					}
					else
					{
						togglePattern.Toggle();
					}
				}
				
				if (togglePattern.CurrentToggleState != ToggleState.ToggleState_Off)
				{
					if (this.uiElement.CurrentFrameworkId == "WPF")
					{
						this.Click();
						Engine.GetInstance().Sleep(100);
					}
					else
					{
						togglePattern.Toggle();
					}
				}

				if (togglePattern.CurrentToggleState == ToggleState.ToggleState_Off)
				{
					return;
				}
            }

            // current element does not support Toggle Pattern
            // todo: use Win32 messages for common Win32 checkboxes
            IntPtr hwnd = IntPtr.Zero;

            try
            {
                hwnd = this.uiElement.CurrentNativeWindowHandle;
            }
            catch
            { }

            bool isWin32Button = false;

            if (hwnd != IntPtr.Zero)
            {
                StringBuilder className = new StringBuilder(256);
                UnsafeNativeFunctions.GetClassName(hwnd, className, 256);

                if (className.ToString() == "Button")
                {
                    isWin32Button = true;

                    // common Win32 checkbox window
                    IntPtr result = UnsafeNativeFunctions.SendMessage(hwnd,
                        ButtonMessages.BM_GETCHECK, IntPtr.Zero, IntPtr.Zero);

                    if (result.ToInt32() != (int)ButtonMessages.BST_UNCHECKED)
                    {
                        UnsafeNativeFunctions.SendMessage(hwnd, ButtonMessages.BM_SETCHECK,
                            new IntPtr(ButtonMessages.BST_UNCHECKED), IntPtr.Zero);
                    }
                }
            }

            if (!isWin32Button)
            {
                //handle when checkbox is not a common Win32 checkbox
            }
        }

		/// <summary>
        /// Puts the checkbox in an indeterminate state if it is possible
        /// </summary>
        public void SetIndeterminate()
        {
            if (this.IsAlive == false)
            {
                throw new Exception("This UI element is not available to the user anymore.");
            }

            object togglePatternObject = uiElement.GetCurrentPattern(UIA_PatternIds.UIA_TogglePatternId);
            IUIAutomationTogglePattern togglePattern = togglePatternObject as IUIAutomationTogglePattern;
            if (togglePattern != null)
            {
                /*int count = 0;
                while (togglePattern.CurrentToggleState != ToggleState.ToggleState_Indeterminate)
                {
                    if (count >= 3)
                    {
                        throw new Exception("Cannot set indeterminate state");
                    }
                    togglePattern.Toggle();
                    count++;
                }*/
				
				if (togglePattern.CurrentToggleState != ToggleState.ToggleState_Indeterminate)
				{
					if (this.uiElement.CurrentFrameworkId == "WPF")
					{
						this.Click();
						Engine.GetInstance().Sleep(100);
					}
					else
					{
						togglePattern.Toggle();
					}
				}
				
				if (togglePattern.CurrentToggleState != ToggleState.ToggleState_Indeterminate)
				{
					if (this.uiElement.CurrentFrameworkId == "WPF")
					{
						this.Click();
						Engine.GetInstance().Sleep(100);
					}
					else
					{
						togglePattern.Toggle();
					}
				}

				if (togglePattern.CurrentToggleState == ToggleState.ToggleState_Indeterminate)
				{
					return;
				}
            }

            // current element does not support Toggle Pattern
            // todo: use Win32 messages for common Win32 checkboxes
            IntPtr hwnd = IntPtr.Zero;

            try
            {
                hwnd = this.uiElement.CurrentNativeWindowHandle;
            }
            catch
            { }

            bool isWin32Button = false;

            if (hwnd != IntPtr.Zero)
            {
                StringBuilder className = new StringBuilder(256);
                UnsafeNativeFunctions.GetClassName(hwnd, className, 256);

                if (className.ToString() == "Button")
                {
                    isWin32Button = true;

                    // common Win32 checkbox window
                    IntPtr result = UnsafeNativeFunctions.SendMessage(hwnd,
                        ButtonMessages.BM_GETCHECK, IntPtr.Zero, IntPtr.Zero);

                    if (result.ToInt32() != (int)ButtonMessages.BST_INDETERMINATE)
                    {
                        UnsafeNativeFunctions.SendMessage(hwnd, ButtonMessages.BM_SETCHECK,
                            new IntPtr(ButtonMessages.BST_INDETERMINATE), IntPtr.Zero);
                    }
                }
            }

            if (!isWin32Button)
            {
                //handle when checkbox is not a common Win32 checkbox
            }
        }

        /// <summary>
        /// Gets or sets the checked state of the checkbox. 
        /// true = checked, false = unchecked, null = indeterminate
        /// </summary>
        public bool? IsChecked
        {
            get
            {
                if (this.IsAlive == false)
                {
                    throw new Exception("This UI element is not available to the user anymore.");
                }

                object togglePatternObject = uiElement.GetCurrentPattern(UIA_PatternIds.UIA_TogglePatternId);
                IUIAutomationTogglePattern togglePattern = togglePatternObject as IUIAutomationTogglePattern;
                if (togglePattern != null)
                {
                    if (togglePattern.CurrentToggleState == ToggleState.ToggleState_Indeterminate)
                    {
                        return null;
                    }
                    else if (togglePattern.CurrentToggleState == ToggleState.ToggleState_Off)
                    {
                        return false;
                    }
                    else
                    {
                        // ToggleState.ToggleState_On
                        return true;
                    }
                }

                IntPtr hwnd = IntPtr.Zero;

                try
                {
                    hwnd = this.uiElement.CurrentNativeWindowHandle;
                }
                catch
                { }

                if (hwnd != IntPtr.Zero)
                {
                    StringBuilder className = new StringBuilder(256);
                    UnsafeNativeFunctions.GetClassName(hwnd, className, 256);

                    if (className.ToString() == "Button")
                    {
                        // common Win32 checkbox window
                        IntPtr result = UnsafeNativeFunctions.SendMessage(hwnd,
                            ButtonMessages.BM_GETCHECK, IntPtr.Zero, IntPtr.Zero);

                        if (result.ToInt32() == (int)ButtonMessages.BST_UNCHECKED)
                        {
                            //unchecked
                            return false;
                        }
                        else if (result.ToInt32() == (int)ButtonMessages.BST_CHECKED)
                        {
                            //checked
                            return true;
                        }
                        else if (result.ToInt32() == (int)ButtonMessages.BST_INDETERMINATE)
                        {
                            //indeterminate
                            return null;
                        }
                    }
                }

                throw new Exception("Cannot get checked state");
            }
            set
            {
                if (value == true)
                {
                    this.Check();
                }
                else if (value == false)
                {
                    this.Uncheck();
                }
                else // value is null
                {
                    this.SetIndeterminate();
                }
            }
        }
		
		/// <summary>
        /// Gets the text of the checkbox
        /// </summary>
		public string Text
		{
			get
			{
				return this.GetText();
			}
		}
		
		private UIA_AutomationPropertyChangedEventHandler UIA_StateChangedEventHandler = null;
		/// <summary>
        /// Delegate for State Changed event
        /// </summary>
		/// <param name="sender">The checkbox that sent the event</param>
		/// <param name="newState">true if checkbox is checked, false if not checked and null if it's in an indeterminate state</param>
		public delegate void StateChanged(UIDA_CheckBox sender, bool? newState);
		internal StateChanged StateChangedHandler = null;
		
		/// <summary>
        /// Attaches/detaches a handler to checked state changed event
        /// </summary>
		public event StateChanged StateChangedEvent
		{
			add
			{
				try
				{
					if (this.StateChangedHandler == null)
					{
						this.UIA_StateChangedEventHandler = new UIA_AutomationPropertyChangedEventHandler(this);
			
						if (base.uiElement.CurrentFrameworkId == "WinForm")
						{
							Engine.uiAutomation.AddPropertyChangedEventHandler(base.uiElement, TreeScope.TreeScope_Element, 
								null, this.UIA_StateChangedEventHandler, new int[] { UIA_PropertyIds.UIA_NamePropertyId });
						}
						else
						{
							Engine.uiAutomation.AddPropertyChangedEventHandler(base.uiElement, TreeScope.TreeScope_Element, 
								null, this.UIA_StateChangedEventHandler, new int[] { UIA_PropertyIds.UIA_ToggleToggleStatePropertyId });
						}
					}
					
					this.StateChangedHandler += value;
				}
				catch {}
			}
			remove
			{
				try
				{
					this.StateChangedHandler -= value;
				
					if (this.StateChangedHandler == null)
					{
						if (this.UIA_StateChangedEventHandler == null)
						{
							return;
						}
						
						System.Threading.Tasks.Task.Run(() => 
						{
							try
							{
								Engine.uiAutomation.RemovePropertyChangedEventHandler(base.uiElement, 
									this.UIA_StateChangedEventHandler);
								UIA_StateChangedEventHandler = null;
							}
							catch { }
						}).Wait(5000);
					}
				}
				catch {}
			}
		}
    }
}
