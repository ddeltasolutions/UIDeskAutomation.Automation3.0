using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a Separator control.
    /// </summary>
    public class UIDA_Separator: ElementBase
    {
        public UIDA_Separator(IUIAutomationElement el)
        {
            base.uiElement = el;
        }
    }
}
