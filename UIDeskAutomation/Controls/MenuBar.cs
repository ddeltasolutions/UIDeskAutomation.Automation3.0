using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Class that represents a MenuBar ui element
    /// </summary>
    public class UIDA_MenuBar: ElementBase
    {
        public UIDA_MenuBar(IUIAutomationElement el)
        {
            base.uiElement = el;
        }
    }
}
