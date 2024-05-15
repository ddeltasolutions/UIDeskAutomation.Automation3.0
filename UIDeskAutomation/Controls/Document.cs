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
		
		/// <summary>
        /// Attaches/detaches a handler to text changed event
        /// </summary>
		public event TextChanged TextChangedEvent
		{
			add
			{
				base.TextChangedEvent += value;
			}
			remove
			{
				base.TextChangedEvent -= value;
			}
		}
		
		/// <summary>
        /// Attaches/detaches a handler to text selection changed event
        /// </summary>
		public event TextSelectionChanged TextSelectionChangedEvent
		{
			add
			{
				base.TextSelectionChangedEvent += value;
			}
			remove
			{
				base.TextSelectionChangedEvent -= value;
			}
		}
    }
}
