using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Class that represents a Tree control.
    /// </summary>
    public class UIDA_Tree: ElementBase
    {
        public UIDA_Tree(IUIAutomationElement el)
        {
            this.uiElement = el;
        }

        /// <summary>
        /// Gets the tree root item.
        /// </summary>
        /// <returns>UIDA_TreeItem element</returns>
        public UIDA_TreeItem GetRoot()
        {
            IUIAutomationElement rootElement = this.FindFirst(UIA_ControlTypeIds.UIA_TreeItemControlTypeId, null,
                false, false, true);

            if (rootElement == null)
            {
                Engine.TraceInLogFile("GetRoot() method - root element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("GetRoot() method - root element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_TreeItem root = new UIDA_TreeItem(rootElement);
            return root;
        }
		
		private UIA_AutomationEventHandler UIA_ElementSelectedEventHandler = null;
		
		/// <summary>
        /// Delegate for Element Selected event
        /// </summary>
		/// <param name="sender">The tree control that sent the event.</param>
		/// <param name="selectedItem">the new selected tree item.</param>
		public delegate void ElementSelected(UIDA_Tree sender, UIDA_TreeItem selectedItem);
		internal ElementSelected ElementSelectedHandle = null;
		
		/// <summary>
        /// Attaches/detaches a handler to element selected event
        /// </summary>
		public event ElementSelected ElementSelectedEvent
		{
			add
			{
				try
				{
					if (this.ElementSelectedHandle == null)
					{
						this.UIA_ElementSelectedEventHandler = new UIA_AutomationEventHandler(this);
		
						Engine.uiAutomation.AddAutomationEventHandler(UIA_EventIds.UIA_SelectionItem_ElementSelectedEventId, 
							base.uiElement, TreeScope.TreeScope_Subtree, null, this.UIA_ElementSelectedEventHandler);
					}
					
					this.ElementSelectedHandle += value;
				}
				catch {}
			}
			remove
			{
				try
				{
					this.ElementSelectedHandle -= value;
				
					if (this.ElementSelectedHandle == null)
					{
						if (this.UIA_ElementSelectedEventHandler == null)
						{
							return;
						}
						
						System.Threading.Tasks.Task.Run(() => 
						{
							try
							{
								Engine.uiAutomation.RemoveAutomationEventHandler(UIA_EventIds.UIA_SelectionItem_ElementSelectedEventId, 
									base.uiElement, this.UIA_ElementSelectedEventHandler);
								UIA_ElementSelectedEventHandler = null;
							}
							catch { }
						}).Wait(5000);
					}
				}
				catch {}
			}
		}
    }
}
