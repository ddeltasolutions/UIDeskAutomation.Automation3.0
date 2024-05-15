using System;
using System.Collections.Generic;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// This class represents a header item.
    /// </summary>
    public class UIDA_HeaderItem : ElementBase
    {
        /// <summary>
        /// Creates a UIDA_HeaderItem using an IUIAutomationElement
        /// </summary>
        /// <param name="el">UI Automation Element</param>
        public UIDA_HeaderItem(IUIAutomationElement el)
        {
            this.uiElement = el;
        }

        /// <summary>
        /// Gets the Text of header item.
        /// </summary>
        public string Text
        {
            get
            {
                IUIAutomationElement text = this.FindFirst(UIA_ControlTypeIds.UIA_TextControlTypeId, null,
                    false, false, true);

                if (text == null)
                {
                    Engine.TraceInLogFile("Header Item: cannot get text");
                    throw new Exception("Header Item: cannot get text");
                }

                string textString = null;

                try
                {
                    textString = text.CurrentName;
                }
                catch (Exception ex)
                {
                    Engine.TraceInLogFile("HeaderItem text: " + ex.Message);
                    throw new Exception("HeaderItem text: " + ex.Message);
                }

                return textString;
            }
        }
    }
}
