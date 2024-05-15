using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a button UI element
    /// </summary>
    public class UIDA_Button: ElementBase
    {
		/// <summary>
        /// Creates a UIDA_Button using an IUIAutomationElement
        /// </summary>
        /// <param name="el">UI Automation Element</param>
        public UIDA_Button(IUIAutomationElement el)
        {
            base.uiElement = el;
        }
        
        /// <summary>
        /// Presses the button
        /// </summary>
        public void Press()
        {
			IUIAutomationTogglePattern togglePattern = uiElement.GetCurrentPattern(UIA_PatternIds.UIA_TogglePatternId) as IUIAutomationTogglePattern;
			
			if (togglePattern != null && this.uiElement.CurrentFrameworkId == "WPF")
			{
				//For WPF ToggleButtons use mouse click because TogglePattern.Toggle() does not call the button's Click event handler
				this.Click();
			}
			else
			{
				try
				{
					base.Invoke();
				}
				catch
				{
					base.Click();
				}
			}
        }
		
		/// <summary>
        /// Gets the text of the button
        /// </summary>
		public string Text
		{
			get
			{
				return this.GetText();
			}
		}
		
		/// <summary>
        /// Gets a boolean to determine if the button is pressed or not. This is supported only for toggle buttons.
        /// </summary>
        public bool IsPressed
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
                    if (togglePattern.CurrentToggleState == ToggleState.ToggleState_On)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }

                throw new Exception("Could not get the pressed state of the button. This is supported only for toggle buttons.");
            }
            /*set
            {
                if (this.IsAlive == false)
                {
                    throw new Exception("This UI element is not available to the user anymore.");
                }

                object togglePatternObject = uiElement.GetCurrentPattern(UIA_PatternIds.UIA_TogglePatternId);
                IUIAutomationTogglePattern togglePattern = togglePatternObject as IUIAutomationTogglePattern;
                if (togglePattern != null)
                {
                    if ((value == false && togglePattern.CurrentToggleState == ToggleState.ToggleState_On) ||
						(value == true && togglePattern.CurrentToggleState != ToggleState.ToggleState_On))
                    {
                        togglePattern.Toggle();
                    }
					return;
                }

                throw new Exception("Could not get the pressed state of the button. This is supported only for toggle buttons.");
            }*/
        }
		
		private UIA_AutomationEventHandler UIA_ClickedEventHandler = null;
		private UIA_AutomationPropertyChangedEventHandler UIA_PropertyChangedEventHandler = null;
		/// <summary>
        /// Delegate for clicked event
        /// </summary>
		/// <param name="sender">The button that sent the event</param>
		public delegate void Clicked(UIDA_Button sender);
		internal Clicked ClickedHandler = null;
		
		/// <summary>
        /// Attaches/detaches a handler to click event
        /// </summary>
		public event Clicked ClickedEvent
		{
			add
			{
				try
				{
					if (this.ClickedHandler == null)
					{
						if (base.uiElement.CurrentFrameworkId == "WinForm")
						{
							this.UIA_PropertyChangedEventHandler = new UIA_AutomationPropertyChangedEventHandler(this);
							Engine.uiAutomation.AddPropertyChangedEventHandler(base.uiElement, TreeScope.TreeScope_Element, 
								null, this.UIA_PropertyChangedEventHandler, new int[] { UIA_PropertyIds.UIA_NamePropertyId });
						}
						else
						{
							this.UIA_ClickedEventHandler = new UIA_AutomationEventHandler(this);
							Engine.uiAutomation.AddAutomationEventHandler(UIA_EventIds.UIA_Invoke_InvokedEventId, 
								base.uiElement, TreeScope.TreeScope_Element, null, this.UIA_ClickedEventHandler);
						}
					}
					
					this.ClickedHandler += value;
				}
				catch {}
			}
			remove
			{
				try
				{
					this.ClickedHandler -= value;
				
					if (this.ClickedHandler == null)
					{
						if (base.uiElement.CurrentFrameworkId == "WinForm")
						{
							RemoveEventHandlerWinForm();
						}
						else
						{
							RemoveEventHandler();
						}
					}
				}
				catch {}
			}
		}
		
		private void RemoveEventHandlerWinForm()
		{
			if (this.UIA_PropertyChangedEventHandler == null)
			{
				return;
			}
			
			System.Threading.Tasks.Task.Run(() => 
			{
				try
				{
					Engine.uiAutomation.RemovePropertyChangedEventHandler(base.uiElement, 
						this.UIA_PropertyChangedEventHandler);
					UIA_PropertyChangedEventHandler = null;
				}
				catch { }
			}).Wait(5000);
		}
		
		private void RemoveEventHandler()
		{
			if (this.UIA_ClickedEventHandler == null)
			{
				return;
			}
			
			System.Threading.Tasks.Task.Run(() => 
			{
				try
				{
					Engine.uiAutomation.RemoveAutomationEventHandler(UIA_EventIds.UIA_Invoke_InvokedEventId, 
						base.uiElement, this.UIA_ClickedEventHandler);
					UIA_ClickedEventHandler = null;
				}
				catch { }
			}).Wait(5000);
		}
    }
}
