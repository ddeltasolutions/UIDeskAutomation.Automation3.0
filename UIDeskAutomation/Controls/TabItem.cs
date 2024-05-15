using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a Tab Item control.
    /// </summary>
    public class UIDA_TabItem: ElementBase
    {
        public UIDA_TabItem(IUIAutomationElement el, UIDA_TabCtrl parent)
        {
            this.uiElement = el;
            this.parent = parent;
        }
		
		public UIDA_TabItem(IUIAutomationElement el)
		{
			this.uiElement = el;
			
			IUIAutomationTreeWalker tw = Engine.uiAutomation.ControlViewWalker;
			IUIAutomationElement uiparent = tw.GetParentElement(this.uiElement);
			
			while (uiparent != null)
			{
				if (uiparent.CurrentControlType == UIA_ControlTypeIds.UIA_TabControlTypeId)
				{
					break;
				}
				uiparent = tw.GetParentElement(uiparent);
			}
			
			if (uiparent != null)
			{
				this.parent = new UIDA_TabCtrl(uiparent);
			}
		}

        private UIDA_TabCtrl parent = null;

        /// <summary>
        /// Returns true if the current tab item is selected, false otherwise.
        /// </summary>
        public bool IsSelected
        {
            get
            {
                object selectionItemPatternObj = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_SelectionItemPatternId);
                IUIAutomationSelectionItemPattern selectionItemPattern = selectionItemPatternObj as IUIAutomationSelectionItemPattern;

                if (selectionItemPattern != null)
                {
                    try
                    {
                        return (selectionItemPattern.CurrentIsSelected != 0);
                    }
                    catch (Exception ex)
                    {
                        Engine.TraceInLogFile("TreeItem.IsSelected failed: " +
                            ex.Message);

                        throw new Exception("TreeItem.IsSelected failed: " +
                            ex.Message);
                    }
                }

                Engine.TraceInLogFile("TreeItem.IsSelected failed");
                throw new Exception("TreeItem.IsSelected failed");
            }
        }

        /// <summary>
        /// Selects the current tab item.
        /// </summary>
        public void Select()
        {
            object selectionItemPatternObj = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_SelectionItemPatternId);
            IUIAutomationSelectionItemPattern selectionItemPattern = selectionItemPatternObj as IUIAutomationSelectionItemPattern;

            if (selectionItemPattern != null)
            {
                try
                {
                    selectionItemPattern.Select();

                    return;
                }
                catch (Exception ex)
                {
                    Engine.TraceInLogFile("TabItem.Select() failed: " + ex.Message);
                    throw new Exception("TabItem.Select() failed: " + ex.Message);
                }
            }
        }

        /// <summary>
        /// Gets the zero based index of the current tab item.
        /// </summary>
        public int Index
        {
            get
            {
				if (this.parent == null)
				{
					return -1; // this tab item is not part of a tab control
				}
                UIDA_TabItem[] tabItems = this.parent.Items;

                for (int i = 0; i < tabItems.Length; i++)
                {
                    UIDA_TabItem tabItem = tabItems[i];

                    if (Helper.CompareAutomationElements(
                        tabItem.uiElement, this.uiElement) == true)
                    {
                        return i;
                    }
                }

                return -1;
            }
        }
		
		/// <summary>
        /// Gets the text of the tab item. The same as calling GetText().
        /// </summary>
		public string Text
		{
			get
			{
				return this.GetText();
			}
		}
    }
}
