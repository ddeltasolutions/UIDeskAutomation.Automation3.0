using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a Tooltip control.
    /// </summary>
    public class UIDA_ToolTip: ElementBase
    {
        public UIDA_ToolTip(IUIAutomationElement el)
        {
            this.uiElement = el;
        }
    }
}
