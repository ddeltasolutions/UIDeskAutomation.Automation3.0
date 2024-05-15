using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// This class represents a Data Grid control.
    /// </summary>
    public class UIDA_DataGrid: ElementBase
    {
        /// <summary>
        /// Creates a UIDA_DataGrid using an IUIAutomationElement
        /// </summary>
        /// <param name="el">UI Automation Element</param>
        public UIDA_DataGrid(IUIAutomationElement el)
        {
            this.uiElement = el;
        }

        /// <summary>
        /// Gets the header of a datagrid control.
        /// </summary>
        public UIDA_Header Header
        {
            get
            {
                IUIAutomationElement header = this.FindFirst(UIA_ControlTypeIds.UIA_HeaderControlTypeId, null,
                    false, false, true);

                if (header == null)
                {
                    Engine.TraceInLogFile("DataGrid Header not found");

                    if (Engine.ThrowExceptionsWhenSearch == true)
                    {
                        throw new Exception("DataGrid Header not found");
                    }
                    else
                    {
                        return null;
                    }
                }

                UIDA_Header dataGridHeader = new UIDA_Header(header);
                return dataGridHeader;
            }
        }

        /// <summary>
        /// Gets the column count.
        /// </summary>
        public int ColumnCount
        {
            get
            {
                IUIAutomationGridPattern gridPattern = this.GetGridPattern();

                if (gridPattern != null)
                {
                    try
                    {
                        return gridPattern.CurrentColumnCount;
                    }
                    catch (Exception ex)
                    {
                        Engine.TraceInLogFile(
                            "DataGrid.ColumnCount - cannot get column count");

                        throw new Exception(
                            "DataGrid.ColumnCount - cannot get column count");
                    }
                }

                Engine.TraceInLogFile("DataGrid.ColumnCount - GridPattern not supported");
                throw new Exception("DataGrid.ColumnCount - GridPattern not supported");
            }
        }

        /// <summary>
        /// Gets the number of rows.
        /// </summary>
        public int RowCount
        {
            get
            {
                IUIAutomationGridPattern gridPattern = this.GetGridPattern();

                if (gridPattern != null)
                {
                    try
                    {
                        return gridPattern.CurrentRowCount;
                    }
                    catch (Exception ex)
                    {
                        Engine.TraceInLogFile("DataGrid.RowCount - cannot get row count");
                        throw new Exception("DataGrid.RowCount - cannot get row count");
                    }
                }

                Engine.TraceInLogFile("DataGrid.RowCount - GridPattern not supported");
                throw new Exception("DataGrid.RowCount - GridPattern not supported");
            }
        }

        /// <summary>
        /// Gets a boolean that tells if multiple items can be selected in the data grid.
        /// </summary>
        public bool CanSelectMultiple
        {
            get
            {
                IUIAutomationSelectionPattern selectionPattern = this.GetSelectionPattern();

                if (selectionPattern != null)
                {
                    try
                    {
                        return (selectionPattern.CurrentCanSelectMultiple != 0);
                    }
                    catch (Exception ex)
                    {
                        Engine.TraceInLogFile("DataGrid.CanSelectMultiple failed: " +
                            ex.Message);

                        throw new Exception("DataGrid.CanSelectMultiple failed: " +
                            ex.Message);
                    }
                }
                else
                {
                    // selectionPattern == null
                    Engine.TraceInLogFile(
                        "DataGrid.CanSelectMultiple - SelectionPattern not supported");

                    throw new Exception(
                        "DataGrid.CanSelectMultiple - SelectionPattern not supported");
                }
            }
        }

        /// <summary>
        /// Returns all rows in the current DataGrid control.
        /// </summary>
        public UIDA_DataItem[] Rows
        {
            get
            {
                if (uiElement.CurrentFrameworkId != "WPF")
                {
                    List<IUIAutomationElement> rows = this.FindAll(UIA_ControlTypeIds.UIA_DataItemControlTypeId,
                        null, true, false, true);

                    List<UIDA_DataItem> returnRows = new List<UIDA_DataItem>();

                    for (int i = 0; i < rows.Count; i++)
                    {
                        IUIAutomationElement row = rows[i];
                        UIDA_DataItem dataItem = new UIDA_DataItem(row);
                        dataItem.m_index = i;
                        dataItem.grid = uiElement;

                        returnRows.Add(dataItem);
                    }

                    return returnRows.ToArray();
                }
                else
                {
                    List<UIDA_DataItem> returnRows = new List<UIDA_DataItem>();
                    
                    object objectPattern = uiElement.GetCurrentPattern(UIA_PatternIds.UIA_ItemContainerPatternId);
                    IUIAutomationItemContainerPattern itemContainerPattern = objectPattern as IUIAutomationItemContainerPattern;
                    if (itemContainerPattern == null)
                    {
                        List<IUIAutomationElement> rows = this.FindAll(UIA_ControlTypeIds.UIA_DataItemControlTypeId,
                            null, true, false, true);
                        
                        for (int i = 0; i < rows.Count; i++)
                        {
                            IUIAutomationElement row = rows[i];
                            UIDA_DataItem dataItem = new UIDA_DataItem(row);
                            dataItem.m_index = i;
                            dataItem.grid = uiElement;
                            returnRows.Add(dataItem);
                        }
                        return returnRows.ToArray();
                    }
                    
                    IUIAutomationElement crt = null;
                    int idx = 0;
                    do
                    {
                        crt = itemContainerPattern.FindItemByProperty(crt, 0, null);
                        if (crt != null)
                        {
                            UIDA_DataItem dataItem = new UIDA_DataItem(crt);
                            dataItem.m_index = idx;
                            dataItem.grid = uiElement;
                            returnRows.Add(dataItem);
                        }
                        idx++;
                    }
                    while (crt != null);
                    
                    return returnRows.ToArray();
                }
            }
        }
        
        /// <summary>
        /// Gets the row at the specified index.
        /// </summary>
        /// <param name="index">zero based index</param>
        public UIDA_DataItem this[int index]
        {
            get
            {
                if (uiElement.CurrentFrameworkId != "WPF")
                {
                    return this.Rows[index];
                }
                else
                {
                    object objectPattern = uiElement.GetCurrentPattern(UIA_PatternIds.UIA_ItemContainerPatternId);
                    IUIAutomationItemContainerPattern itemContainerPattern = objectPattern as IUIAutomationItemContainerPattern;
                    
                    if (itemContainerPattern == null)
                    {
                        return Rows[index];
                    }
                    else
                    {
                        if (index < 0)
                        {
                            throw new Exception("Index cannot be negative");
                        }
                        IUIAutomationElement crt = null;
                        int idx = index;
                        do
                        {
                            crt = itemContainerPattern.FindItemByProperty(crt, 0, null);
                            if (crt == null)
                            {
                                throw new Exception("Index too big");
                            }
                            index--;
                        }
                        while (index >= 0);
                        
                        UIDA_DataItem dataItem = new UIDA_DataItem(crt);
                        dataItem.m_index = idx;
                        dataItem.grid = uiElement;
                        return dataItem;
                    }
                }
            }
        }

        /// <summary>
        /// Gets all groups in a DataGrid control.
        /// </summary>
        public UIDA_Group[] Groups
        {
            get
            {
                List<IUIAutomationElement> allGroups = this.FindAll(UIA_ControlTypeIds.UIA_GroupControlTypeId,
                    null, false, false, true);

                List<UIDA_Group> returnGroups = new List<UIDA_Group>();

                foreach (IUIAutomationElement group in allGroups)
                {
                    UIDA_Group returnGroup = new UIDA_Group(group);
                    returnGroups.Add(returnGroup);
                }

                return returnGroups.ToArray();
            }
        }

        /// <summary>
        /// Selects all items in a DataGrid control.
        /// </summary>
        public void SelectAll()
        {
            UIDA_DataItem[] items = this.Rows;

            foreach (UIDA_DataItem item in items)
            {
                try
                {
                    item.AddToSelection();
                }
                catch (Exception ex)
                { }
            }
        }

        /// <summary>
        /// Clears all selection in a DataGrid control.
        /// </summary>
        public void ClearAllSelection()
        {
            UIDA_DataItem[] selectedItems = this.SelectedRows;

            foreach (UIDA_DataItem selectedItem in selectedItems)
            {
                try
                {
                    selectedItem.RemoveFromSelection();
                }
                catch (Exception ex)
                { }
            }
        }

        /// <summary>
        /// Returns selected rows in the current DataGrid control.
        /// </summary>
        public UIDA_DataItem[] SelectedRows
        {
            get
            {
                IUIAutomationSelectionPattern selectionPattern = this.GetSelectionPattern();

                if (selectionPattern == null)
                {
                    Engine.TraceInLogFile(
                        "DataGrid.SelectedRows - SelectionPattern not supported");

                    throw new Exception(
                        "DataGrid.SelectedRows - SelectionPattern not supported");
                }

                try
                {
                    IUIAutomationElementArray selectedItems =
                        selectionPattern.GetCurrentSelection();

                    List<UIDA_DataItem> returnCollection = new List<UIDA_DataItem>();

                    for (int i = 0; i < selectedItems.Length; i++)
                    {
                        IUIAutomationElement selectedItem = selectedItems.GetElement(i);
                        UIDA_DataItem dataItem = new UIDA_DataItem(selectedItem);
                        dataItem.m_index = IndexOf(dataItem);
                        dataItem.grid = uiElement;

                        returnCollection.Add(dataItem);
                    }

                    return returnCollection.ToArray();
                }
                catch (Exception ex)
                {
                    Engine.TraceInLogFile(
                        "DataGrid.SelectedRows failed: " + ex.Message);

                    throw new Exception(
                        "DataGrid.SelectedRows failed: " + ex.Message);
                }
            }
        }
        
        private int IndexOf(UIDA_DataItem item)
        {
            UIDA_DataItem[] items = this.Rows;
            for (int i = 0; i < items.Length; i++)
            {
                if (item.IsEqual(items[0]))
                {
                    return i;
                }
            }
            return -1;
        }
        
        /// <summary>
        /// Scrolls the DataGrid vertically and horizontally using the specified percents.
        /// </summary>
        /// <param name="percentVertical">percentage to scroll vertically</param>
        /// <param name="percentHorizontal">percentage to scroll horizontally</param>
        public void Scroll(double percentVertical, double percentHorizontal = 0)
        {
            object objectPattern = uiElement.GetCurrentPattern(UIA_PatternIds.UIA_ScrollPatternId);
            IUIAutomationScrollPattern scrollPattern = objectPattern as IUIAutomationScrollPattern;
            
            if (scrollPattern != null)
            {
                try
                {
                    scrollPattern.SetScrollPercent(percentHorizontal, percentVertical);
                }
                catch (Exception ex)
                {
                    Engine.TraceInLogFile("Scroll(): " + ex.Message);
                }
            }
        }
        
        private IUIAutomationElement GetItemByText(string text)
        {
            object objectPattern = uiElement.GetCurrentPattern(UIA_PatternIds.UIA_ItemContainerPatternId);
            IUIAutomationItemContainerPattern itemContainerPattern = objectPattern as IUIAutomationItemContainerPattern;
            
            if (itemContainerPattern != null)
            {
                return itemContainerPattern.FindItemByProperty(null, 30005, text);
            }
            return null;
        }
        
        private IUIAutomationElement GetWPFDataItem(int index)
        {
            object objectPattern = uiElement.GetCurrentPattern(UIA_PatternIds.UIA_ItemContainerPatternId);
            IUIAutomationItemContainerPattern itemContainerPattern = objectPattern as IUIAutomationItemContainerPattern;
            
            if (itemContainerPattern == null)
            {
                return null;
            }
            
            IUIAutomationElement crt = null;
            do
            {
                crt = itemContainerPattern.FindItemByProperty(crt, 0, null);
                if (crt == null)
                {
                    Engine.TraceInLogFile("list item index too big");
                    break;
                }
                index--;
            }
            while (index != 0);
            
            if (index == 0)
            {
                return crt;
            }
            else
            {
                return null;
            }
        }
        
        /// <summary>
        /// Selects an item in a DataGrid by the item index. Other selected items will be deselected.
        /// </summary>
        /// <param name="index">item index, starts with 1</param>
        public void Select(int index)
        {
            if (uiElement.CurrentFrameworkId == "WPF")
            {
                IUIAutomationElement item = GetWPFDataItem(index);
                if (item != null)
                {
                    UIDA_DataItem dataItemWPF = new UIDA_DataItem(item);
                    dataItemWPF.Select();
                    return;
                }
            }
            
            UIDA_DataItem dataItem = DataItemAt(null, index, true);
            if (dataItem == null)
            {
                Engine.TraceInLogFile("Item not found");
                throw new Exception("Item not found");
            }
            dataItem.Select();
        }
        
        /// <summary>
        /// Selects an item (or more items) in a DataGrid by the item text. Other selected items will be deselected.
        /// </summary>
        /// <param name="itemText">Item text. Wildcards can be used.</param>
        /// <param name="selectAll">true to select all items matching the given text, false to select only the first item matching the given text</param>
        /// <param name="caseSensitive">true if the item text search is done case sensitive</param>
        /*public void Select(string itemText = null, bool selectAll = true, bool caseSensitive = true)
        {
            if (uiElement.CurrentFrameworkId == "WPF")
            {
                IUIAutomationElement item = GetItemByText(itemText);
                if (item != null)
                {
                    DataItem dataItem = new DataItem(item);
                    dataItem.Select();
                    return;
                }
                else
                {
                    Engine.TraceInLogFile("FindItemByProperty didn't find the item");
                }
            }
            
            if (selectAll == false)
            {
                DataItem dataItem = DataItem(itemText, true, caseSensitive);
                if (dataItem == null)
                {
                    Engine.TraceInLogFile("Item not found");
                    throw new Exception("Item not found");
                    return;
                }
                
                dataItem.Select();
            }
            else
            {
                DataItem[] items = DataItems(itemText, true, caseSensitive);
                if (items.Length == 0)
                {
                    return;
                }
                items[0].Select();
                for (int i = 1; i < items.Length; i++)
                {
                    items[i].AddToSelection();
                }
            }
        }*/
        
        /// <summary>
        /// Adds an item to selection in a DataGrid by the item index.
        /// </summary>
        /// <param name="index">item index, starts with 1</param>
        public void AddToSelection(int index)
        {
            if (uiElement.CurrentFrameworkId == "WPF")
            {
                IUIAutomationElement item = GetWPFDataItem(index);
                if (item != null)
                {
                    UIDA_DataItem dataItemWPF = new UIDA_DataItem(item);
                    dataItemWPF.AddToSelection();
                    return;
                }
            }
            
            UIDA_DataItem dataItem = DataItemAt(null, index, true);
            if (dataItem == null)
            {
                Engine.TraceInLogFile("Item not found");
                throw new Exception("Item not found");
            }
            dataItem.AddToSelection();
        }
        
        /// <summary>
        /// Adds an item (or more items) to selection in a DataGrid by the item text.
        /// </summary>
        /// <param name="itemText">Item text. Wildcards can be used.</param>
        /// <param name="selectAll">true to add to selection all items matching the given text, false to add to selection only the first item matching the given text</param>
        /// <param name="caseSensitive">true if the item text search is done case sensitive</param>
        /*public void AddToSelection(string itemText = null, bool selectAll = true, bool caseSensitive = true)
        {
            if (uiElement.CurrentFrameworkId == "WPF")
            {
                IUIAutomationElement item = GetItemByText(itemText);
                if (item != null)
                {
                    DataItem dataItem = new DataItem(item);
                    dataItem.AddToSelection();
                    return;
                }
                else
                {
                    Engine.TraceInLogFile("FindItemByProperty didn't find the item");
                }
            }
            
            if (selectAll == false)
            {
                DataItem dataItem = DataItem(itemText, true, caseSensitive);
                if (dataItem == null)
                {
                    Engine.TraceInLogFile("Item not found");
                    throw new Exception("Item not found");
                    return;
                }
                
                dataItem.AddToSelection();
            }
            else
            {
                DataItem[] items = DataItems(itemText, true, caseSensitive);

                foreach (DataItem item in items)
                {
                    item.AddToSelection();
                }
            }
        }*/
        
        /// <summary>
        /// Removes an item from selection in a DataGrid by the item index.
        /// </summary>
        /// <param name="index">item index, starts with 1</param>
        public void RemoveFromSelection(int index)
        {
            if (uiElement.CurrentFrameworkId == "WPF")
            {
                IUIAutomationElement item = GetWPFDataItem(index);
                if (item != null)
                {
                    UIDA_DataItem dataItemWPF = new UIDA_DataItem(item);
                    dataItemWPF.RemoveFromSelection();
                    return;
                }
            }
            
            UIDA_DataItem dataItem = DataItemAt(null, index, true);
            if (dataItem == null)
            {
                Engine.TraceInLogFile("Item not found");
                throw new Exception("Item not found");
            }
            dataItem.RemoveFromSelection();
        }
        
        /// <summary>
        /// Removes an item (or more items) from selection in a DataGrid by the item text.
        /// </summary>
        /// <param name="itemText">Item text. Wildcards can be used.</param>
        /// <param name="all">true to remove from selection all items matching the given text, false to remove from selection only the first item matching the given text</param>
        /// <param name="caseSensitive">true if the item text search is done case sensitive</param>
        /*public void RemoveFromSelection(string itemText = null, bool all = true, bool caseSensitive = true)
        {
            if (uiElement.CurrentFrameworkId == "WPF")
            {
                IUIAutomationElement item = GetItemByText(itemText);
                if (item != null)
                {
                    DataItem dataItem = new DataItem(item);
                    dataItem.RemoveFromSelection();
                    return;
                }
                else
                {
                    Engine.TraceInLogFile("FindItemByProperty didn't find the item");
                }
            }
            
            if (all == false)
            {
                DataItem dataItem = DataItem(itemText, true, caseSensitive);
                if (dataItem == null)
                {
                    Engine.TraceInLogFile("Item not found");
                    throw new Exception("Item not found");
                    return;
                }
                
                dataItem.RemoveFromSelection();
            }
            else
            {
                DataItem[] items = DataItems(itemText, true, caseSensitive);

                foreach (DataItem item in items)
                {
                    item.RemoveFromSelection();
                }
            }
        }*/

        private IUIAutomationSelectionPattern GetSelectionPattern()
        {
            object selectionPatternObj = uiElement.GetCurrentPattern(UIA_PatternIds.UIA_SelectionPatternId);
            IUIAutomationSelectionPattern selectionPattern = selectionPatternObj as IUIAutomationSelectionPattern;

            return selectionPattern;
        }

        private IUIAutomationGridPattern GetGridPattern()
        {
            object gridPatternObj = uiElement.GetCurrentPattern(UIA_PatternIds.UIA_GridPatternId);
            IUIAutomationGridPattern gridPattern = gridPatternObj as IUIAutomationGridPattern;

            return gridPattern;
        }
    }
}
