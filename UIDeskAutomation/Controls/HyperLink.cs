using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a Hyper Link control.
    /// </summary>
    public class UIDA_HyperLink: ElementBase
    {
        public UIDA_HyperLink(IUIAutomationElement el)
        {
            this.uiElement = el;
        }
        
        /// <summary>
        /// Accesses the link
        /// </summary>
        public void AccessLink()
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
}
