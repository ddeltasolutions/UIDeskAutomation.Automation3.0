using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a listitem UI element.
    /// </summary>
    public class UIDA_ListItem: ElementBase
    {
        public UIDA_ListItem(IUIAutomationElement el)
        {
            base.uiElement = el;
        }

        /// <summary>
        /// Deselects any selected items and selects the current list item
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
                }
                catch (Exception ex)
                {
                    Engine.TraceInLogFile("ListItem.Select failed: " + ex.Message);
                }
            }
        }

        /// <summary>
        /// Removes the current list item from the collection of selected list items
        /// </summary>
        public void RemoveFromSelection()
        {
            object selectionItemPatternObj = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_SelectionItemPatternId);
            IUIAutomationSelectionItemPattern selectionItemPattern = selectionItemPatternObj as IUIAutomationSelectionItemPattern;

            if (selectionItemPattern != null)
            {
                try
                {
                    selectionItemPattern.RemoveFromSelection();
                }
                catch (Exception ex)
                {
                    Engine.TraceInLogFile("ListItem.Deselect failed: " + ex.Message);
                }
            }
        }

        /// <summary>
        /// Adds the current list item to the collection of selected list items
        /// </summary>
        public void AddToSelection()
        {
            object selectionItemPatternObj = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_SelectionItemPatternId);
            IUIAutomationSelectionItemPattern selectionItemPattern = selectionItemPatternObj as IUIAutomationSelectionItemPattern;

            if (selectionItemPattern != null)
            {
                try
                {
                    selectionItemPattern.AddToSelection();
                }
                catch (Exception ex)
                {
                    Engine.TraceInLogFile("AddToSelection failed: " + ex.Message);
                }
            }
        }

        /// <summary>
        /// Gets the zero based index of this list item in the list items collection.
        /// </summary>
        public int Index
        {
            get
            {
                IUIAutomationTreeWalker treeWalker = Engine.uiAutomation.ControlViewWalker;
                IUIAutomationElement parent = treeWalker.GetParentElement(this.uiElement);

				UIDA_ListItem[] listItems = null;
                while (parent != null)
                {
                    try
                    {
                        if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_ListControlTypeId)
                        {
							UIDA_List parentList = new UIDA_List(parent);
							listItems = parentList.Items;
                            break;
                        }
						if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_ComboBoxControlTypeId)
						{
							UIDA_ComboBox parentComboBox = new UIDA_ComboBox(parent);
							listItems = parentComboBox.Items;
                            break;
						}
                    }
                    catch
                    {
                        break;
                    }

                    parent = treeWalker.GetParentElement(parent);
                }

                if (parent == null)
                {
                    Engine.TraceInLogFile("Error getting ListItem index");
                    throw new Exception("Error getting ListItem index");
                }

                //UIDA_List parentList = new UIDA_List(parent);
                //UIDA_ListItem[] listItems = parentList.Items;

                int index = -1;

                foreach (UIDA_ListItem currentListItem in listItems)
                {
                    index++;

                    Array runtimeId1 = null;
                    Array runtimeId2 = null;

                    try
                    {
                        runtimeId1 = this.uiElement.GetRuntimeId();
                        runtimeId2 = currentListItem.uiElement.GetRuntimeId();
                    }
                    catch { }

                    if ((runtimeId1 == null) || (runtimeId2 == null))
                    {
                        continue;
                    }

                    if (runtimeId1.Length != runtimeId2.Length)
                    {
                        continue;
                    }

                    bool different = false;

                    for (int i = 0; i < runtimeId1.Length; i++)
                    {
                        int id1 = (int)runtimeId1.GetValue(i);
                        int id2 = (int)runtimeId2.GetValue(i);

                        if (id1 != id2)
                        {
                            different = true;
                            break;
                        }
                    }

                    if (different == true)
                    {
                        continue;
                    }

                    return index;
                }

                return -1; // not found.
            }
        }

        /// <summary>
        /// Returns true if the current list item is selected, false otherwise.
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
                        Engine.TraceInLogFile("ListItem.IsSelected failed: " +
                            ex.Message);
                    }
                }

                return false;
            }
        }

        /// <summary>
        /// Brings the current List Item into viewable area of the parent List control.
        /// </summary>
        public void BringIntoView()
        {
            object scrollItemPatternObj = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_ScrollItemPatternId);
            IUIAutomationScrollItemPattern scrollItemPattern = scrollItemPatternObj as IUIAutomationScrollItemPattern;

            if (scrollItemPattern == null)
            {
                Engine.TraceInLogFile("ListItem.BringIntoView method failed");
                throw new Exception("ListItem.BringIntoView method failed");
            }

            try
            {
                scrollItemPattern.ScrollIntoView();
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("ListItem.BringIntoView method failed: " +
                    ex.Message);

                throw new Exception("ListItem.BringIntoView method failed: " +
                    ex.Message);
            }
        }

        /// <summary>
        /// Gets or sets the checked state of the current list item.
        /// true = checked, false = unchecked, null = indeterminate state
        /// </summary>
        public bool? IsChecked
        {
            get
            {
                IntPtr hwndList = this.GetWindow();

                if (hwndList != IntPtr.Zero)
                {
                    StringBuilder className = new StringBuilder(256);
                    UnsafeNativeFunctions.GetClassName(hwndList, className, 256);

                    if (className.ToString() == "SysListView32")
                    { 
                        // Win32 standard listview control
                        int index = this.Index;

                        IntPtr itemState = UnsafeNativeFunctions.SendMessage(
				            hwndList, WindowMessages.LVM_GETITEMSTATE, 
                            new IntPtr(index), 
                            new IntPtr(Win32Constants.LVIS_STATEIMAGEMASK));

                        return ((itemState.ToInt32() & 0x2000) != 0);
                    }
                }

                object togglePatternObj = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_TogglePatternId);
                IUIAutomationTogglePattern togglePattern = togglePatternObj as IUIAutomationTogglePattern;

                if (togglePattern == null)
                {
                    Engine.TraceInLogFile("ListItem.IsChecked failed");
                    throw new Exception("ListItem.IsChecked failed");
                }

                try
                {
                    if (togglePattern.CurrentToggleState == ToggleState.ToggleState_On)
                    {
                        return true;
                    }
                    else if (togglePattern.CurrentToggleState == ToggleState.ToggleState_Off)
                    {
                        return false;
                    }
                    else
                    {
                        return null;
                    }
                }
                catch (Exception ex)
                { 
                    Engine.TraceInLogFile("ListItem.IsChecked failed: " + ex.Message);
                    throw new Exception("ListItem.Checked failed: " + ex.Message);
                }
            }

            set
            {
                IntPtr hwndList = this.GetWindow();

                if (hwndList != IntPtr.Zero)
                {
                    StringBuilder className = new StringBuilder(256);
                    UnsafeNativeFunctions.GetClassName(hwndList, className, 256);

                    if (className.ToString() == "SysListView32")
                    {
                        LV_ITEM lvItem = new LV_ITEM();

                        int index = this.Index;

                        lvItem.stateMask = Win32Constants.LVIS_STATEIMAGEMASK;
                        if (value == true)
                        {
                            lvItem.state = 0x2000; // check list item
                        }
                        else
                        {
                            lvItem.state = 0x1000; // uncheck list item
                        }

                        uint processId = 0;

                        UnsafeNativeFunctions.GetWindowThreadProcessId(hwndList,
                            out processId);

                        IntPtr hProcess = UnsafeNativeFunctions.OpenProcess(
                            ProcessAccessFlags.VirtualMemoryWrite | ProcessAccessFlags.VirtualMemoryOperation, 
                            false, (int)processId);

                        if (hProcess != IntPtr.Zero)
                        {
                            IntPtr lvItemPtrExtern = UnsafeNativeFunctions.VirtualAllocEx(
                                hProcess, IntPtr.Zero, (uint)Marshal.SizeOf(lvItem), 
                                AllocationType.Reserve | AllocationType.Commit, 
                                MemoryProtection.ReadWrite);
                            Debug.Assert(lvItemPtrExtern != IntPtr.Zero);
                            
                            IntPtr lvItemPtr = Marshal.AllocHGlobal(Marshal.SizeOf(lvItem));
                            Debug.Assert(lvItemPtr != IntPtr.Zero);

                            System.Runtime.InteropServices.Marshal.StructureToPtr(lvItem,
                                lvItemPtr, false);

                            IntPtr numberOfBytesWritten = IntPtr.Zero;

                            bool bResult = UnsafeNativeFunctions.WriteProcessMemory(
                                hProcess, lvItemPtrExtern, lvItemPtr,
                                Marshal.SizeOf(lvItem), out numberOfBytesWritten);
                            Debug.Assert(bResult == true);

                            Marshal.FreeHGlobal(lvItemPtr);

                            IntPtr retPtr = UnsafeNativeFunctions.SendMessage(hwndList,
                                WindowMessages.LVM_SETITEMSTATE, new IntPtr(index),
                                lvItemPtrExtern);
                            Debug.Assert(retPtr != IntPtr.Zero);

                            bResult = UnsafeNativeFunctions.VirtualFreeEx(hProcess, 
                                lvItemPtrExtern, Marshal.SizeOf(lvItem), 
                                FreeType.Decommit | FreeType.Release);
                            //Debug.Assert(bResult == true);

                            bResult = UnsafeNativeFunctions.CloseHandle(hProcess);
                            Debug.Assert(bResult == true);
                        }

                        return;
                    }
                }

                object togglePatternObj = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_TogglePatternId);
                IUIAutomationTogglePattern togglePattern = togglePatternObj as IUIAutomationTogglePattern;

                if (togglePattern == null)
                {
                    if (value == true)
                    {
                        Engine.TraceInLogFile("Cannot check list item");
                        throw new Exception("Cannot check list item");
                    }
                    else if (value == false)
                    {
                        Engine.TraceInLogFile("Cannot uncheck list item");
                        throw new Exception("Cannot uncheck list item");
                    }
					else
					{
						Engine.TraceInLogFile("Cannot set list item checked state");
                        throw new Exception("Cannot set list item checked state");
					}
                }
                
                ToggleState desiredState = ToggleState.ToggleState_Indeterminate;
                if (value == true)
                {
                    desiredState = ToggleState.ToggleState_On;
                }
                else if (value == false)
                {
                    desiredState = ToggleState.ToggleState_Off;
                }
                
                int tries = 3;
                while (tries > 0)
                {
                    if (togglePattern.CurrentToggleState == desiredState)
                    {
                        break;
                    }
                    togglePattern.Toggle();
                    tries--;
                }
                
                if (tries == 0)
                {
                    tries = 3;
                    while (tries > 0)
                    {
                        if (togglePattern.CurrentToggleState == desiredState)
                        {
                            break;
                        }
                        //this.SimulateDoubleClick();
						this.BringIntoView();
						this.DoubleClick();
                        tries--;
                    }
                }
                
                if (tries == 0)
                {
                    Engine.TraceInLogFile("Cannot set list item checked state");
                    throw new Exception("Cannot set list item checked state");
                }
            }
        }
		
		/// <summary>
        /// Gets the text of the list item.
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
