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
				base.Invoke();
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
    }
}
