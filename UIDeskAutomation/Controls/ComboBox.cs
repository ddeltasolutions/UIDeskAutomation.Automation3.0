using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a ComboBox UI element.
    /// </summary>
    public class UIDA_ComboBox: ElementBase
    {
        /// <summary>
        /// Creates a UIDA_ComboBox using an IUIAutomationElement
        /// </summary>
        /// <param name="el">UI Automation Element</param>
        public UIDA_ComboBox(IUIAutomationElement el)
        {
            this.uiElement = el;
        }

        /// <summary>
        /// Gets a collection with all ComboBox items.
        /// </summary>
        public UIDA_ListItem[] Items
        {
            get
            {
                if (uiElement.CurrentFrameworkId == "WPF")
                {
                    List<UIDA_ListItem> returnRows = new List<UIDA_ListItem>();
                    
                    object objectPattern = uiElement.GetCurrentPattern(UIA_PatternIds.UIA_ItemContainerPatternId);
                    IUIAutomationItemContainerPattern itemContainerPattern = objectPattern as IUIAutomationItemContainerPattern;
                    if (itemContainerPattern == null)
                    {
                        List<IUIAutomationElement> allListItems = 
                            this.FindAll(UIA_ControlTypeIds.UIA_ListItemControlTypeId, null, true, false, true);
                        foreach (IUIAutomationElement listItemEl in allListItems)
                        {
                            UIDA_ListItem listItem = new UIDA_ListItem(listItemEl);
                            returnRows.Add(listItem);
                        }
                        return returnRows.ToArray();
                    }
                    
                    IUIAutomationElement crt = null;
                    do
                    {
                        crt = itemContainerPattern.FindItemByProperty(crt, /*null*/ 0, null);
                        if (crt != null)
                        {
                            UIDA_ListItem listItem = new UIDA_ListItem(crt);
                            returnRows.Add(listItem);
                        }
                    }
                    while (crt != null);
                    
                    return returnRows.ToArray();
                }
                else
                {
                    UIDA_Pane desktop = Engine.GetInstance().GetDesktopPane();
                    UIDA_List[] lists = desktop.Lists(uiElement.CurrentName, false, true);
                    if (lists.Length > 0)
                    {
                        return lists[0].ListItems();
                    }
                    
                    List<IUIAutomationElement> allItems =
                        this.FindAll(UIA_ControlTypeIds.UIA_ListItemControlTypeId, null, true, false, true);

                    List<UIDA_ListItem> returnCollection = new List<UIDA_ListItem>();

                    foreach (IUIAutomationElement el in allItems)
                    {
                        UIDA_ListItem listItem = new UIDA_ListItem(el);
                        returnCollection.Add(listItem);
                    }

                    return returnCollection.ToArray();
                }
            }
        }

        /// <summary>
        /// Searches for a combobox item.
        /// </summary>
        /// <param name="name">name of combobox item</param>
        /// <param name="caseSensitive">true is name search is done case sensitive, default true</param>
        /// <returns>ComboBox item</returns>
        public UIDA_ListItem ListItem(string name = null, bool caseSensitive = true)
        {
            UIDA_Pane desktop = Engine.GetInstance().GetDesktopPane();
            UIDA_List[] lists = desktop.Lists(uiElement.CurrentName, false, true);
            if (lists.Length > 0)
            {
                return lists[0].ListItem(name, false, caseSensitive);
            }

            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_ListItemControlTypeId, name,
                true, false, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("ComboBox::ListItem method - list item element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ComboBox::ListItem method - list item element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_ListItem listItem = new UIDA_ListItem(returnElement);
            return listItem;
        }

        /// <summary>
        /// Searches for a ComboBox item with specified name at specified index.
        /// </summary>
        /// <param name="name">text of ComboBox item</param>
        /// <param name="index">index</param>
        /// <param name="caseSensitive">true is name search is done case sensitive, false otherwise, default true</param>
        /// <returns>UIDA_ListItem element of ComboBox</returns>
        public UIDA_ListItem ListItemAt(string name, int index, bool caseSensitive = true)
        {
            UIDA_Pane desktop = Engine.GetInstance().GetDesktopPane();
            UIDA_List[] lists = desktop.Lists(uiElement.CurrentName, false, true);
            if (lists.Length > 0)
            {
                return lists[0].ListItemAt(name, index, false, caseSensitive);
            }
            
            if (index < 0)
            {
                Engine.TraceInLogFile("ComboBox::ListItemAt method - index cannot be negative");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ComboBox::ListItemAt method - index cannot be negative");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_ListItemControlTypeId, name, index, true,
                false, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("ComboBox::ListItemAt method - ComboBox item element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ComboBox::ListItemAt method - ComboBox item element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("ComboBox::ListItemAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ComboBox::ListItemAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_ListItem listItem = new UIDA_ListItem(returnElement);
            return listItem;
        }

        /// <summary>
        /// Returns a collection of ListItems that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of ListItem elements, use null to return all ListItems</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_ListItem elements</returns>
        new public UIDA_ListItem[] ListItems(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            UIDA_Pane desktop = Engine.GetInstance().GetDesktopPane();
            UIDA_List[] lists = desktop.Lists(uiElement.CurrentName, false, true);
            if (lists.Length > 0)
            {
                return lists[0].ListItems(name, false, caseSensitive);
            }
            
            List<IUIAutomationElement> allListItems = FindAll(UIA_ControlTypeIds.UIA_ListItemControlTypeId,
                name, searchDescendants, false, caseSensitive);

            List<UIDA_ListItem> listitems = new List<UIDA_ListItem>();
            if (allListItems != null)
            {
                foreach (IUIAutomationElement crtEl in allListItems)
                {
                    listitems.Add(new UIDA_ListItem(crtEl));
                }
            }
            return listitems.ToArray();
        }

        /// <summary>
        /// Sets a text for a drop-down ComboBox.
        /// </summary>
        /// <param name="text">text to set</param>
        public void SetText(string text)
        {
			IUIAutomationValuePattern valuePattern = uiElement.GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId) as IUIAutomationValuePattern;
			if (valuePattern != null)
			{
				try
				{
					valuePattern.SetValue(text);
					return;
				}
				catch {}
			}
		
            UIDA_Edit edit = this.Edit();

            if (edit == null)
            {
                Engine.TraceInLogFile("ComboBox::SetText failed: cannot set text, ComboBox should have DropDown style");
                throw new Exception("ComboBox::SetText failed: cannot set text. ComboBox should have DropDown style");
            }

            edit.SetText(text);
        }

        /// <summary>
        /// Gets the selected ComboBox item. If this returns null you can try the Text property to get the text of the selected item.
        /// </summary>
        public UIDA_ListItem SelectedItem
        {
            get
            {
                object objectPattern = uiElement.GetCurrentPattern(UIA_PatternIds.UIA_SelectionPatternId);
                IUIAutomationSelectionPattern selectionPattern = objectPattern as IUIAutomationSelectionPattern;
                
                if (selectionPattern == null)
                {
                    UIDA_ListItem[] allItems = this.Items;

                    foreach (UIDA_ListItem item in allItems)
                    {
                        if (item.IsSelected)
                        {
                            return item;
                        }
                    }
                }
                else
                {
                    // SelectionPattern is supported
                    IUIAutomationElementArray selection = selectionPattern.GetCurrentSelection();
                    if (selection.Length >= 1)
                    {
                        UIDA_ListItem listItem = new UIDA_ListItem(selection.GetElement(0));
                        return listItem;
                    }
                }
                
                // no item is selected
                return null;
            }
        }
		
		/// <summary>
        /// Gets the text of the selected item.
        /// </summary>
		public string Text
		{
			get
			{
				return this.GetText();
			}
		}
        
        /// <summary>
        /// Expands the combobox.
        /// </summary>
        public void Expand()
        {
            object objectPattern = uiElement.GetCurrentPattern(UIA_PatternIds.UIA_ExpandCollapsePatternId);
            IUIAutomationExpandCollapsePattern expandCollapsePattern = objectPattern as IUIAutomationExpandCollapsePattern;
            
            if (expandCollapsePattern != null)
            {
                expandCollapsePattern.Expand();
            }
            else
            {
                UIDA_Button[] buttons = this.Buttons();
                if (buttons.Length > 0)
                {
                    buttons[0].Invoke();
                }
            }
        }
        
        /// <summary>
        /// Collapses the combobox.
        /// </summary>
        public void Collapse()
        {
            object objectPattern = uiElement.GetCurrentPattern(UIA_PatternIds.UIA_ExpandCollapsePatternId);
            IUIAutomationExpandCollapsePattern expandCollapsePattern = objectPattern as IUIAutomationExpandCollapsePattern;
            
            if (expandCollapsePattern != null)
            {
                expandCollapsePattern.Collapse();
            }
            else
            {
                UIDA_Button[] buttons = this.Buttons();
                if (buttons.Length > 0)
                {
                    buttons[0].Invoke();
                }
            }
        }
        
        /// <summary>
        /// Selects an item in a ComboBox by the item index.
        /// </summary>
        /// <param name="index">item index, starts with 1</param>
        public void Select(int index)
        {
            string fid = this.uiElement.CurrentFrameworkId;
            //if (fid == "WPF" || fid == "WinForm")
            {
                Expand();
                Thread.Sleep(100);
            }
            
            UIDA_ListItem listItem = ListItemAt(null, index);
            if (listItem == null)
            {
                Engine.TraceInLogFile("Item not found");
                throw new Exception("Item not found");
            }
			
            if (fid == "Win32" || fid == "WinForm")
            {
                listItem.BringIntoView();
				//listItem.Click();
				
				try
				{
					tagRECT boundingRect = listItem.uiElement.CurrentBoundingRectangle;
					Engine.GetInstance().ClickScreenCoordinatesAt( (boundingRect.left + boundingRect.right) / 2, 
						(boundingRect.top + boundingRect.bottom) / 2 );
				}
				catch
				{
					listItem.KeyPress(VirtualKeys.Enter);
				}
            }
            else
            {
                listItem.Select();
            }
            
			if (fid == "WPF")
            {
                Collapse();
            }
        }
        
        /// <summary>
        /// Selects an item in a ComboBox by the item text. Wildcards can be used.
        /// </summary>
        /// <param name="itemText">item text</param>
        /// <param name="caseSensitive">true if the item text search is done case sensitive</param>
        public void Select(string itemText = null, bool caseSensitive = true)
        {
            string fid = this.uiElement.CurrentFrameworkId;
            //if (fid == "WPF" || fid == "WinForm")
            {
                Expand();
                Thread.Sleep(100);
            }
            
            UIDA_ListItem listItem = ListItem(itemText, caseSensitive);
            if (listItem == null)
            {
                Engine.TraceInLogFile("Item not found");
                throw new Exception("Item not found");
            }
            
            if (fid == "Win32" || fid == "WinForm")
            {
                listItem.BringIntoView();
				//listItem.Click();
				
				try
				{
					tagRECT boundingRect = listItem.uiElement.CurrentBoundingRectangle;
					Engine.GetInstance().ClickScreenCoordinatesAt( (boundingRect.left + boundingRect.right) / 2, 
						(boundingRect.top + boundingRect.bottom) / 2 );
				}
				catch
				{
					listItem.KeyPress(VirtualKeys.Enter);
				}
            }
            else
            {
                listItem.Select();
            }
            
            if (fid == "WPF")
            {
                Collapse();
            }
        }
		
		private UIA_AutomationPropertyChangedEventHandler UIA_SelectionChangedEventHandler = null;
		private UIA_AutomationEventHandler UIA_ElementSelectedEventHandler = null;
		/// <summary>
        /// Delegate for Selection Changed event
        /// </summary>
		/// <param name="sender">The ComboBox that sent the event</param>
		/// <param name="selectedItem">the text of the selected item</param>
		public delegate void SelectionChanged(UIDA_ComboBox sender, string selectedItem);
		internal SelectionChanged SelectionChangedHandle = null;
		
		/// <summary>
        /// Attaches/detaches a handler to selection changed event
        /// </summary>
		public event SelectionChanged SelectionChangedEvent
		{
			add
			{
				try
				{
					if (this.SelectionChangedHandle == null)
					{
						string cfid = base.uiElement.CurrentFrameworkId;
						if (cfid == "WinForm" || cfid == "XAML")
						{
							this.UIA_ElementSelectedEventHandler = new UIA_AutomationEventHandler(this);
						
							Engine.uiAutomation.AddAutomationEventHandler(UIA_EventIds.UIA_SelectionItem_ElementSelectedEventId, 
								base.uiElement, TreeScope.TreeScope_Subtree, null, this.UIA_ElementSelectedEventHandler);
						}
						else
						{
							this.UIA_SelectionChangedEventHandler = new UIA_AutomationPropertyChangedEventHandler(this);
						
							Engine.uiAutomation.AddPropertyChangedEventHandler(base.uiElement, TreeScope.TreeScope_Element, 
								null, this.UIA_SelectionChangedEventHandler, new int[] { UIA_PropertyIds.UIA_ValueValuePropertyId });
						}
					}
					
					this.SelectionChangedHandle += value;
				}
				catch {}
			}
			remove
			{
				try
				{
					this.SelectionChangedHandle -= value;
				
					if (this.SelectionChangedHandle == null)
					{
						string cfid = base.uiElement.CurrentFrameworkId;
						if (cfid == "WinForm" || cfid == "XAML")
						{
							RemoveEventHandlerWinForm();
						}
						else
						{
							RemoveEventHandler();
						}
					}
				}
				catch {}
			}
		}
		
		private void RemoveEventHandlerWinForm()
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
		
		private void RemoveEventHandler()
		{
			if (this.UIA_SelectionChangedEventHandler == null)
			{
				return;
			}
			
			System.Threading.Tasks.Task.Run(() => 
			{
				try
				{
					Engine.uiAutomation.RemovePropertyChangedEventHandler(base.uiElement, 
						this.UIA_SelectionChangedEventHandler);
					UIA_SelectionChangedEventHandler = null;
				}
				catch { }
			}).Wait(5000);
		}
    }
}
