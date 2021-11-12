using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a Document UI element
    /// </summary>
    public class UIDA_Document: UIDA_Edit
    {
        /// <summary>
        /// Creates a UIDA_Document using an IUIAutomationElement
        /// </summary>
        /// <param name="el">UI Automation Element</param>
        public UIDA_Document(IUIAutomationElement el)
        {
            base.uiElement = el;
        }
    }
}
