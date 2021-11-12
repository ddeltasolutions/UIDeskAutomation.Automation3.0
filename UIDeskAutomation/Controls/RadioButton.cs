using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using System.Windows.Automation;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a RadioButton UI element.
    /// </summary>
    public class UIDA_RadioButton: ElementBase
    {
        public UIDA_RadioButton(IUIAutomationElement el)
        {
            this.uiElement = el;
        }

        /// <summary>
        /// Gets the selected state of the radio button.
        /// </summary>
        /// <returns>true - is selected, false otherwise</returns>
        public bool IsSelected
        {
			get
			{
				object selectionItemPatternObj = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_SelectionItemPatternId);
				IUIAutomationSelectionItemPattern selectionItemPattern = selectionItemPatternObj as IUIAutomationSelectionItemPattern;

				if (selectionItemPattern == null)
				{
					Engine.TraceInLogFile("RadioButton.GetIsSelected() - SelectionItemPattern not supported");
					throw new Exception("RadioButton.GetIsSelected() - SelectionItemPattern not supported");
				}

				return (selectionItemPattern.CurrentIsSelected != 0);
			}
        }

        /// <summary>
        /// Selects a radio button.
        /// </summary>
        public void Select()
        {
			this.Click();
			//Engine.GetInstance().Sleep(100);
			
			/*if (this.uiElement.CurrentFrameworkId == "WPF")
			{
				// for WPF use mouse click because IUIAutomationSelectionItemPattern.Select() is not calling the radio button's Click event handler
				this.Click();
				return;
			}
			
            object selectionItemPatternObj = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_SelectionItemPatternId);
            IUIAutomationSelectionItemPattern selectionItemPattern = selectionItemPatternObj as IUIAutomationSelectionItemPattern;

            if (selectionItemPattern == null)
            {
                Engine.TraceInLogFile("RadioButton.Select() - SelectionItemPattern not supported");
                throw new Exception("RadioButton.Select() - SelectionItemPattern not supported");
            }

            selectionItemPattern.Select(); */
        }
		
		/// <summary>
        /// Gets the text of the radio button
        /// </summary>
		public string Text
		{
			get
			{
				return this.GetText();
			}
		}
    }
}
