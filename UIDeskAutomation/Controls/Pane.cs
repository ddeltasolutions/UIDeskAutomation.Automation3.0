using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Class that represents a Pane ui element
    /// </summary>
    public class UIDA_Pane : ElementBase
    {
        public UIDA_Pane(IUIAutomationElement el)
        {
            base.uiElement = el;
        }
    }
}
