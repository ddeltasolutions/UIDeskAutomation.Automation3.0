using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a Thumb control.
    /// </summary>
    public class UIDA_Thumb: ElementBase
    {
        public UIDA_Thumb(IUIAutomationElement el)
        {
            this.uiElement = el;
        }
    }
}
