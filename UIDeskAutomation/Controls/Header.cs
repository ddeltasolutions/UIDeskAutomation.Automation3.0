using System;
using System.Collections.Generic;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
	/// <summary>
	/// This class represents a header.
	/// </summary>
	public class UIDA_Header: ElementBase
	{
		/// <summary>
		/// Creates a UIDA_Header using an IUIAutomationElement
		/// </summary>
		/// <param name="el">UI Automation Element</param>
		public UIDA_Header(IUIAutomationElement el)
		{
			this.uiElement = el;
		}

		/// <summary>
		/// Gets the header items in a header.
		/// </summary>
		public UIDA_HeaderItem[] Items
		{
			get
			{
				List<IUIAutomationElement> items = this.FindAll(UIA_ControlTypeIds.UIA_HeaderItemControlTypeId,
					null, false, false, true);

				List<UIDA_HeaderItem> headerItems = new List<UIDA_HeaderItem>();

				foreach (IUIAutomationElement item in items)
				{ 
					UIDA_HeaderItem headerItem = new UIDA_HeaderItem(item);
					headerItems.Add(headerItem);
				}

				return headerItems.ToArray();
			}
		}
	}
}