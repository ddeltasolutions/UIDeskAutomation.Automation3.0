using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Class that represents a Top Level Menu element, like a Contextual menu.
    /// </summary>
    public class UIDA_TopLevelMenu : ElementBase
    {
        public UIDA_TopLevelMenu(IUIAutomationElement el)
        {
            base.uiElement = el;
        }
    }
}
