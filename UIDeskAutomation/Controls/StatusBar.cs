using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a StatusBar control.
    /// </summary>
    public class UIDA_StatusBar: ElementBase
    {
        public UIDA_StatusBar(IUIAutomationElement el)
        {
            this.uiElement = el;
        }
    }
}
