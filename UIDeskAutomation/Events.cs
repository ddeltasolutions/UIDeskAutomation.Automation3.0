using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using UIAutomationClient;
using System.IO;

namespace UIDeskAutomationLib
{
	internal class UIA_AutomationEventHandler: IUIAutomationEventHandler
	{
		private UIDA_Window window = null;
		private UIDA_Edit edit = null;
		private UIDA_List list = null;
		private UIDA_Button button = null;
		private UIDA_RadioButton radioButton = null;
		private UIDA_ComboBox comboBox = null;
		private UIDA_Tree tree = null;
		private UIDA_MenuItem menuitem = null;
		
		private string previousSelectedText = "";
	
		public UIA_AutomationEventHandler(UIDA_Window window)
		{
			this.window = window;
		}
		
		public UIA_AutomationEventHandler(UIDA_Edit edit)
		{
			this.edit = edit;
		}
		
		public UIA_AutomationEventHandler(UIDA_List list)
		{
			this.list = list;
		}
		
		public UIA_AutomationEventHandler(UIDA_Button button)
		{
			this.button = button;
		}
		
		public UIA_AutomationEventHandler(UIDA_RadioButton radioButton)
		{
			this.radioButton = radioButton;
		}
		
		public UIA_AutomationEventHandler(UIDA_ComboBox comboBox)
		{
			this.comboBox = comboBox;
		}
		
		public UIA_AutomationEventHandler(UIDA_Tree tree)
		{
			this.tree = tree;
		}
		
		public UIA_AutomationEventHandler(UIDA_MenuItem menuitem)
		{
			this.menuitem = menuitem;
		}
	
		public void HandleAutomationEvent(IUIAutomationElement sender, int eventId)
		{
			if (this.window != null)
			{
				if (eventId == UIA_EventIds.UIA_Window_WindowClosedEventId && this.window.OnWindowClosed != null)
				{
					DateTime now = DateTime.Now;
					if (lastTime != null && (now - lastTime.Value) < TimeSpan.FromMilliseconds(100))
					{
						lastTime = now;
						return;
					}
					lastTime = now;
					this.window.OnWindowClosed();
				}
			}
			else if (this.edit != null)
			{
				if (eventId == UIA_EventIds.UIA_Text_TextChangedEventId && this.edit.TextChangedHandler != null)
				{
					string text = null;
					try
					{
						text = this.edit.GetText();
					}
					catch { }
					
					this.edit.TextChangedHandler(this.edit, text);
				}
				else if (eventId == UIA_EventIds.UIA_Text_TextSelectionChangedEventId && this.edit.TextSelectionChangedHandler != null)
				{
					string selectedText = null;
					try
					{
						selectedText = this.edit.GetSelectedText();
					}
					catch {}
					
					if (selectedText != previousSelectedText)
					{
						previousSelectedText = selectedText;
						this.edit.TextSelectionChangedHandler(this.edit, selectedText);
					}
				}
			}
			else if (this.list != null)
			{
				if (eventId == UIA_EventIds.UIA_SelectionItem_ElementSelectedEventId && 
					this.list.ElementSelectedHandler != null)
				{
					this.list.ElementSelectedHandler(this.list, new UIDA_ListItem(sender));
				}
				else if (eventId == UIA_EventIds.UIA_SelectionItem_ElementAddedToSelectionEventId &&
					this.list.ElementAddedToSelectionHandler != null)
				{
					this.list.ElementAddedToSelectionHandler(this.list, new UIDA_ListItem(sender));
				}
				else if (eventId == UIA_EventIds.UIA_SelectionItem_ElementRemovedFromSelectionEventId &&
					this.list.ElementRemovedFromSelectionHandler != null)
				{
					this.list.ElementRemovedFromSelectionHandler(this.list, new UIDA_ListItem(sender));
				}
			}
			else if (this.tree != null)
			{
				if (eventId == UIA_EventIds.UIA_SelectionItem_ElementSelectedEventId && 
					this.tree.ElementSelectedHandle != null)
				{
					this.tree.ElementSelectedHandle(this.tree, new UIDA_TreeItem(sender));
				}
			}
			else if (this.button != null)
			{
				if (eventId == UIA_EventIds.UIA_Invoke_InvokedEventId && this.button.ClickedHandler != null)
				{
					this.button.ClickedHandler(this.button);
				}
			}
			else if (this.radioButton != null)
			{
				if (eventId == UIA_EventIds.UIA_Invoke_InvokedEventId && this.radioButton.StateChangedHandler != null)
				{
					bool isSelected = true;
					try
					{
						isSelected = this.radioButton.IsSelected;
					}
					catch { }
					
					this.radioButton.StateChangedHandler(this.radioButton, isSelected);
				}
			}
			else if (this.comboBox != null)
			{
				if (sender.CurrentControlType == UIA_ControlTypeIds.UIA_ListItemControlTypeId && 
					eventId == UIA_EventIds.UIA_SelectionItem_ElementSelectedEventId && 
					this.comboBox.SelectionChangedHandle != null)
				{
					string text = null;
					try
					{
						text = (new UIDA_ListItem(sender)).GetText();
					}
					catch { }
					
					this.comboBox.SelectionChangedHandle(this.comboBox, text);
				}
			}
			else if (this.menuitem != null)
			{
				if (eventId == UIA_EventIds.UIA_Invoke_InvokedEventId && this.menuitem.ClickedHandler != null)
				{
					this.menuitem.ClickedHandler(this.menuitem);
				}
				if (eventId == UIA_EventIds.UIA_MenuOpenedEventId && this.menuitem.ExpandedHandler != null)
				{
					DateTime now = DateTime.Now;
					if (lastTime != null && (now - lastTime.Value) < TimeSpan.FromMilliseconds(100))
					{
						lastTime = now;
						return;
					}
					lastTime = now;
					this.menuitem.ExpandedHandler(this.menuitem);
				}
			}
		}
		
		private DateTime? lastTime = null;
	}
	
	internal class UIA_AutomationPropertyChangedEventHandler: IUIAutomationPropertyChangedEventHandler
	{
		private UIDA_CheckBox checkBox = null;
		private UIDA_RadioButton radioButton = null;
		private UIDA_ComboBox comboBox = null;
		private GenericSpinner genericSpinner = null;
		private UIDA_Edit edit = null;
		private UIDA_Button button = null;
		private UIDA_TreeItem treeItem = null;
		private UIDA_MenuItem menuItem = null;
		private UIDA_Window window = null;
		
		public UIA_AutomationPropertyChangedEventHandler(UIDA_CheckBox checkBox)
		{
			this.checkBox = checkBox;
		}
		
		public UIA_AutomationPropertyChangedEventHandler(UIDA_RadioButton radioButton)
		{
			this.radioButton = radioButton;
		}
		
		public UIA_AutomationPropertyChangedEventHandler(UIDA_ComboBox comboBox)
		{
			this.comboBox = comboBox;
		}
		
		public UIA_AutomationPropertyChangedEventHandler(GenericSpinner genericSpinner)
		{
			this.genericSpinner = genericSpinner;
		}
		
		public UIA_AutomationPropertyChangedEventHandler(UIDA_Edit edit)
		{
			this.edit = edit;
		}
		
		public UIA_AutomationPropertyChangedEventHandler(UIDA_Button button)
		{
			this.button = button;
		}
		
		public UIA_AutomationPropertyChangedEventHandler(UIDA_TreeItem treeItem)
		{
			this.treeItem = treeItem;
		}
		
		public UIA_AutomationPropertyChangedEventHandler(UIDA_MenuItem menuItem)
		{
			this.menuItem = menuItem;
		}
		
		public UIA_AutomationPropertyChangedEventHandler(UIDA_Window window)
		{
			this.window = window;
		}
	
		public void HandlePropertyChangedEvent(IUIAutomationElement sender, int propertyId, object newValue)
		{
			if (this.checkBox != null)
			{
				if (this.checkBox.StateChangedHandler == null)
				{
					return;
				}
			
				if (propertyId == UIA_PropertyIds.UIA_ToggleToggleStatePropertyId)
				{
					/*if (newValue == null)
					{
						this.checkBox.StateChangedHandler(this.checkBox, null);
						return;
					}
					
					ToggleState toggleState = (ToggleState)newValue;
					if (toggleState == ToggleState.ToggleState_On)
					{
						this.checkBox.StateChangedHandler(this.checkBox, true);
					}
					else if (toggleState == ToggleState.ToggleState_Off)
					{
						this.checkBox.StateChangedHandler(this.checkBox, false);
					}
					else // indeterminate
					{
						this.checkBox.StateChangedHandler(this.checkBox, null);
					}*/
					
					this.checkBox.StateChangedHandler(this.checkBox, this.checkBox.IsChecked);
				}
				else if (propertyId == UIA_PropertyIds.UIA_NamePropertyId && 
					sender.CurrentFrameworkId == "WinForm")
				{
					bool? isChecked = null;
					try
					{
						isChecked = this.checkBox.IsChecked;
					}
					catch { }
					
					this.checkBox.StateChangedHandler(this.checkBox, isChecked);
				}
			}
			else if (this.radioButton != null)
			{
				if (this.radioButton.StateChangedHandler == null)
				{
					return;
				}
			
				if (propertyId == UIA_PropertyIds.UIA_SelectionItemIsSelectedPropertyId)
				{
					if (!(newValue is bool))
					{
						return;
					}
					
					bool state = (bool)newValue;
					this.radioButton.StateChangedHandler(this.radioButton, state);
				}
				else if (propertyId == UIA_PropertyIds.UIA_NamePropertyId && sender.CurrentFrameworkId == "WinForm")
				{
					bool isSelected = false;
					try
					{
						isSelected = (new UIDA_RadioButton(sender)).IsSelected;
					}
					catch { }
					
					this.radioButton.StateChangedHandler(this.radioButton, isSelected);
				}
					
			}
			else if (this.comboBox != null)
			{
				if (propertyId == UIA_PropertyIds.UIA_ValueValuePropertyId && 
					this.comboBox.SelectionChangedHandle != null)
				{
					if (newValue != null && !(newValue is string))
					{
						return;
					}
					
					this.comboBox.SelectionChangedHandle(this.comboBox, (string)newValue);
				}
			}
			else if (this.genericSpinner != null)
			{
				if (propertyId == UIA_PropertyIds.UIA_RangeValueValuePropertyId && 
					this.genericSpinner.ValueChangedHandler != null)
				{
					double newValueDouble = 0.0;
					try
					{
						newValueDouble = (double)newValue;
					}
					catch { return; }
					/*if (!(newValue is double))
					{
						return;
					}*/
					
					this.genericSpinner.ValueChangedHandler(this.genericSpinner, newValueDouble);
				}
				else if (propertyId == UIA_PropertyIds.UIA_ValueValuePropertyId && 
					this.genericSpinner.ValueChangedHandler != null)
				{
					this.genericSpinner.ValueChangedHandler(this.genericSpinner, this.genericSpinner.Value);
				}
			}
			else if (this.edit != null)
			{
				if (propertyId == UIA_PropertyIds.UIA_ValueValuePropertyId && 
					this.edit.TextChangedHandler != null && sender.CurrentFrameworkId /*== "Win32"*/ != "WPF")
				{
					string text = null;
					try
					{
						text = this.edit.GetText();
					}
					catch { }
				
					this.edit.TextChangedHandler(this.edit, text);
				}
			}
			else if (this.button != null)
			{
				if (propertyId == UIA_PropertyIds.UIA_NamePropertyId && this.button.ClickedHandler != null && 
					sender.CurrentFrameworkId == "WinForm")
				{
					this.button.ClickedHandler(this.button);
				}
			}
			else if (this.treeItem != null)
			{
				if (propertyId == UIA_PropertyIds.UIA_ExpandCollapseExpandCollapseStatePropertyId && this.treeItem.ExpandedHandler != null)
				{
					try
					{
						ExpandCollapseState state = (ExpandCollapseState)newValue;
						if (state == ExpandCollapseState.ExpandCollapseState_Expanded)
						{
							this.treeItem.ExpandedHandler(this.treeItem, true);
						}
						else
						{
							this.treeItem.ExpandedHandler(this.treeItem, false);
						}
					}
					catch { }
				}
			}
			else if (this.menuItem != null)
			{
				if (propertyId == UIA_PropertyIds.UIA_ExpandCollapseExpandCollapseStatePropertyId && this.menuItem.ExpandedHandler != null)
				{
					try
					{
						ExpandCollapseState state = (ExpandCollapseState)newValue;
						if (state == ExpandCollapseState.ExpandCollapseState_Expanded)
						{
							this.menuItem.ExpandedHandler(this.menuItem);
						}
					}
					catch { }
				}
			}
			/*else if (this.window != null)
			{
				if (propertyId == UIA_PropertyIds.UIA_WindowWindowVisualStatePropertyId && newValue != null)
				{
					try
					{
						WindowVisualState state = (WindowVisualState)newValue;
						if (state == WindowVisualState.WindowVisualState_Minimized)
						{
							if (this.window.MinimizedHandler != null)
							{
								this.window.MinimizedHandler(this.window);
							}
						}
						else if (state == WindowVisualState.WindowVisualState_Maximized)
						{
							if (this.window.MaximizedHandler != null)
							{
								this.window.MaximizedHandler(this.window);
							}
						}
						else if (state == WindowVisualState.WindowVisualState_Normal)
						{
							if (this.window.RestoredHandler != null)
							{
								this.window.RestoredHandler(this.window);
							}
						}
					}
					catch {}
				}
				else if (propertyId == UIA_PropertyIds.UIA_BoundingRectanglePropertyId)
				{
					try
					{
						UIDA_Rect newRect = null;
						try
						{
							tagRECT rect = this.window.uiElement.CurrentBoundingRectangle;
							newRect = new UIDA_Rect(rect.left, rect.top, rect.right, rect.bottom);
						}
						catch {}
						
						if (this.window.WindowMovedHandler != null)
						{
							this.window.WindowMovedHandler(this.window, newRect);
						}
					}
					catch {}
				}
			}*/
		}
	}
}