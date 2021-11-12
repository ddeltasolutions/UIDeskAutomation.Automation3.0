using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a Group UI element
    /// </summary>
    public class UIDA_Group: ElementBase
    {
        /// <summary>
        /// Creates a UIDA_Group using an IUIAutomationElement
        /// </summary>
        /// <param name="el">UI Automation Element</param>
        public UIDA_Group(IUIAutomationElement el)
        {
            base.uiElement = el;
        }
    }
}
