using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// This class represents a Static Text control.
    /// </summary>
    public class UIDA_Label: ElementBase
    {
        public UIDA_Label(IUIAutomationElement el)
        {
            this.uiElement = el;
        }
		
		/// <summary>
        /// Gets the text of the label.
        /// </summary>
        /// <returns>text of the label</returns>
		public string Text
		{
			get
			{
				return this.GetText();
			}
		}
    }
}
