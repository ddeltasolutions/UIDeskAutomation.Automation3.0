using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// This class represents a custom control.
    /// </summary>
    public class UIDA_Custom: ElementBase
    {
        /// <summary>
        /// Creates a UIDA_Custom using an IUIAutomationElement
        /// </summary>
        /// <param name="el">UI Automation Element</param>
        public UIDA_Custom(IUIAutomationElement el)
        {
            this.uiElement = el;
        }
    }
}
