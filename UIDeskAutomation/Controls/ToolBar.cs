using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a Toolbar control.
    /// </summary>
    public class UIDA_ToolBar: ElementBase
    {
        public UIDA_ToolBar(IUIAutomationElement el)
        {
            this.uiElement = el;
        }
    }
}
