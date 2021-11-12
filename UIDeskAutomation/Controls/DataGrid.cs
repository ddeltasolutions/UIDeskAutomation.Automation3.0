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
    public class UIDA_DataGrid: DataGridBase //ElementBase
    {
        /// <summary>
        /// Creates a UIDA_DataGrid using an IUIAutomationElement
        /// </summary>
        /// <param name="el">UI Automation Element</param>
        public UIDA_DataGrid(IUIAutomationElement el): base(el)
        {
            //this.uiElement = el;
        }

        /// <summary>
        /// Gets the header of a datagrid control.
        /// </summary>
        public DataGridHeader Header
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

                DataGridHeader dataGridHeader = new DataGridHeader(header);
                return dataGridHeader;
            }
        }

        /// <summary>
        /// This class represents a header in datagrid control.
        /// </summary>
        public class DataGridHeader: DataGridBase
        {
            /// <summary>
            /// Creates a DataGridHeader using an IUIAutomationElement
            /// </summary>
            /// <param name="el">UI Automation Element</param>
            public DataGridHeader(IUIAutomationElement el): base(el)
            {
                //this.uiElement = el;
            }

            /// <summary>
            /// Gets the header items in a datagrid header.
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
        new public DataGridGroup[] Groups
        {
            get
            {
                List<IUIAutomationElement> allGroups = this.FindAll(UIA_ControlTypeIds.UIA_GroupControlTypeId,
                    null, false, false, true);

                List<DataGridGroup> returnGroups = new List<DataGridGroup>();

                foreach (IUIAutomationElement group in allGroups)
                {
                    DataGridGroup returnGroup = new DataGridGroup(group);
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

    /// <summary>
    /// This class represents a header item in a datagrid.
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
        /// Text of header item.
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

    /// <summary>
    /// This class represents a group.
    /// </summary>
    public class DataGridGroup : DataGridBase
    {
        /// <summary>
        /// Creates a DataGridGroup using an IUIAutomationElement
        /// </summary>
        /// <param name="el">UI Automation Element</param>
        public DataGridGroup(IUIAutomationElement el)
            : base(el)
        { }

        /// <summary>
        /// Expands a group inside a data grid.
        /// </summary>
        public void Expand()
        {
            IUIAutomationExpandCollapsePattern expandCollapsePattern = this.GetExpandCollapsePattern();

            if (expandCollapsePattern == null)
            {
                Engine.TraceInLogFile(
                    "Group.Expand() - ExpandCollapsePattern not supported");

                throw new Exception(
                    "Group.Expand() - ExpandCollapsePattern not supported");
            }

            try
            {
                if (expandCollapsePattern.CurrentExpandCollapseState !=
                    ExpandCollapseState.ExpandCollapseState_Expanded)
                {
                    expandCollapsePattern.Expand();
                }
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("Group.Expand() error: " + ex.Message);
                throw new Exception("Group.Expand() error: " + ex.Message);
            }
        }

        /// <summary>
        /// Collapses a group inside a data grid.
        /// </summary>
        public void Collapse()
        {
            IUIAutomationExpandCollapsePattern expandCollapsePattern =
                this.GetExpandCollapsePattern();

            if (expandCollapsePattern == null)
            {
                Engine.TraceInLogFile(
                    "Group.Collapse() - ExpandCollapsePattern not supported");

                throw new Exception(
                    "Group.Collapse() - ExpandCollapsePattern not supported");
            }

            try
            {
                if (expandCollapsePattern.CurrentExpandCollapseState !=
                    ExpandCollapseState.ExpandCollapseState_Collapsed)
                {
                    expandCollapsePattern.Collapse();
                }
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("Group.Collapse() error: " + ex.Message);
                throw new Exception("Group.Collapse() error: " + ex.Message);
            }
        }

        private IUIAutomationExpandCollapsePattern GetExpandCollapsePattern()
        {
            object expandCollapsePatternObj = uiElement.GetCurrentPattern(UIA_PatternIds.UIA_ExpandCollapsePatternId);
            
            IUIAutomationExpandCollapsePattern expandCollapsePattern =
                    expandCollapsePatternObj as IUIAutomationExpandCollapsePattern;

            return expandCollapsePattern;
        }
    }

    /// <summary>
    /// This class represents a data item.
    /// </summary>
    public class UIDA_DataItem : DataGridBase
    {
        internal int m_index = -1;
        internal IUIAutomationElement grid = null;
        
        /// <summary>
        /// Creates a UIDA_DataItem using an IUIAutomationElement
        /// </summary>
        /// <param name="el">UI Automation Element</param>
        public UIDA_DataItem(IUIAutomationElement el)
            : base(el)
        { }
        
        internal bool IsEqual(UIDA_DataItem item)
        {
            return Helper.CompareAutomationElements(uiElement, item.InnerElement);
        }

        /// <summary>
        /// Gets the selection state of a DataItem in a DataGrid.
        /// </summary>
        public bool IsSelected
        {
            get
            {
                IUIAutomationSelectionItemPattern selectionItemPattern = this.GetSelectionItemPattern();

                if (selectionItemPattern == null)
                {
                    Engine.TraceInLogFile(
                        "DataItem.IsSelected - SelectionItemPattern not supported");

                    throw new Exception(
                        "DataItem.IsSelected - SelectionItemPattern not supported");
                }

                try
                {
                    return (selectionItemPattern.CurrentIsSelected != 0);
                }
                catch (Exception ex)
                {
                    Engine.TraceInLogFile(
                        "DataItem.IsSelected failed: " + ex.Message);

                    throw new Exception(
                        "DataItem.IsSelected failed: " + ex.Message);
                }
            }
        }

        /// <summary>
        /// Select the current item and deselects all other selected items.
        /// </summary>
        public void Select()
        {
            IUIAutomationSelectionItemPattern selectionItemPattern = this.GetSelectionItemPattern();

            if (selectionItemPattern == null)
            {
                Engine.TraceInLogFile(
                    "DataItem.Select() - SelectionItemPattern not supported");

                throw new Exception(
                    "DataItem.Select() - SelectionItemPattern not supported");
            }

            try
            {
                selectionItemPattern.Select();
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("DataItem.Select() failed: " + ex.Message);
                throw new Exception("DataItem.Select() failed: " + ex.Message);
            }
        }

        /// <summary>
        /// Adds the current DataItem to the selected items.
        /// </summary>
        public void AddToSelection()
        {
            IUIAutomationSelectionItemPattern selectionItemPattern = this.GetSelectionItemPattern();

            if (selectionItemPattern == null)
            {
                Engine.TraceInLogFile(
                    "DataItem.AddToSelection() - SelectionItemPattern not supported");

                throw new Exception(
                    "DataItem.AddToSelection() - SelectionItemPattern not supported");
            }

            try
            {
                selectionItemPattern.AddToSelection();
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile(
                    "DataItem.AddToSelection() failed: " + ex.Message);

                throw new Exception(
                    "DataItem.AddToSelection() failed: " + ex.Message);
            }
        }

        /// <summary>
        /// Removes the current DataItem from selected items.
        /// </summary>
        public void RemoveFromSelection()
        {
            IUIAutomationSelectionItemPattern selectionItemPattern = this.GetSelectionItemPattern();

            if (selectionItemPattern == null)
            {
                Engine.TraceInLogFile(
                    "DataItem.RemoveFromSelection() - SelectionItemPattern not supported");

                throw new Exception(
                    "DataItem.RemoveFromSelection() - SelectionItemPattern not supported");
            }

            try
            {
                selectionItemPattern.RemoveFromSelection();
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile(
                    "DataItem.RemoveFromSelection() failed: " + ex.Message);

                throw new Exception(
                    "DataItem.RemoveFromSelection() failed: " + ex.Message);
            }
        }

        private IUIAutomationSelectionItemPattern GetSelectionItemPattern()
        {
            object selectionItemPatternObj = uiElement.GetCurrentPattern(UIA_PatternIds.UIA_SelectionItemPatternId);

            IUIAutomationSelectionItemPattern selectionItemPattern =
                selectionItemPatternObj as IUIAutomationSelectionItemPattern;

            return selectionItemPattern;
        }
        
        /// <summary>
        /// Gets the value at the specified column index.
        /// </summary>
        /// <param name="columnIndex">zero based column index</param>
        public string this[int columnIndex]
        {
            get
            {
                object objectPattern = null;
                if (grid != null)
                {
                    objectPattern = grid.GetCurrentPattern(UIA_PatternIds.UIA_GridPatternId);
                    IUIAutomationGridPattern gridPattern = objectPattern as IUIAutomationGridPattern;
                    
                    if (gridPattern != null && m_index >= 0)
                    {
                        //Engine.TraceInLogFile("columnIndex = " + columnIndex);
                        IUIAutomationElement el = gridPattern.GetItem(m_index, columnIndex);
                        return (new UIDA_Custom(el)).GetText();
                    }
                }
                
                objectPattern = uiElement.GetCurrentPattern(UIA_PatternIds.UIA_ItemContainerPatternId);
                IUIAutomationItemContainerPattern itemContainerPattern = objectPattern as IUIAutomationItemContainerPattern;
                
                if (itemContainerPattern == null)
                {
                    IUIAutomationElementArray collection = uiElement.FindAll(TreeScope.TreeScope_Children, Engine.uiAutomation.CreateTrueCondition());
                    return (new UIDA_Custom(collection.GetElement(columnIndex))).GetText();
                }
                else
                {
                    if (columnIndex < 0)
                    {
                        throw new Exception("Index cannot be negative");
                    }
                    IUIAutomationElement crt = null;
                    do
                    {
                        crt = itemContainerPattern.FindItemByProperty(crt, 0, null);
                        if (crt == null)
                        {
                            throw new Exception("Index too big");
                        }
                        columnIndex--;
                    }
                    while (columnIndex >= 0);
                    
                    return (new UIDA_Custom(crt)).GetText();
                }
            }
        }
        
        /// <summary>
        /// Gets the value at the specified column.
        /// </summary>
        /// <param name="columnName">column name</param>
        public string this[string columnName]
        {
            get
            {
                IUIAutomationTreeWalker tw = Engine.uiAutomation.ControlViewWalker;
                IUIAutomationElement gridEl = tw.GetParentElement(this.uiElement);
                UIDA_DataGrid grid = new UIDA_DataGrid(gridEl);
                
                UIDA_HeaderItem[] headerItems = grid.Header.Items;
                int columnIndex = -1;
                for (int i = 0; i < headerItems.Length; i++)
                {
                    if (columnName == headerItems[i].Text)
                    {
                        columnIndex = i;
                        break;
                    }
                }
                
                if (columnIndex >= 0)
                {
                    return this[columnIndex];
                }
                else
                {
                    throw new Exception("No column with this name");
                }
            }
        }
    }

    /// <summary>
    /// Represents base class for DataGrid control. This class cannot be instantiated.
    /// </summary>
    abstract public class DataGridBase: ElementBase
    {
        internal DataGridBase(IUIAutomationElement el)
        {
            this.uiElement = el;
        }

        /// <summary>
        /// Searches a header item in the current element
        /// </summary>
        /// <param name="name">text of header item</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is case sensitive, default true</param>
        /// <returns>UIDA_HeaderItem element</returns>
        public UIDA_HeaderItem HeaderItem(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_HeaderItemControlTypeId,
                name, searchDescendants, false, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("HeaderItem method - HeaderItem element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("HeaderItem method - HeaderItem element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_HeaderItem headerItem = new UIDA_HeaderItem(returnElement);
            return headerItem;
        }

        /// <summary>
        /// Searches for a HeaderItem with a specified text at a specified index.
        /// </summary>
        /// <param name="name">text of HeaderItem</param>
        /// <param name="index">index of HeaderItem</param>
        /// <param name="searchDescendants">true if search through descendants, false if search only through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_HeaderItem element</returns>
        public UIDA_HeaderItem HeaderItemAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("HeaderItemAt method - index cannot be negative");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("HeaderItemAt method - index cannot be negative");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_HeaderItemControlTypeId, name, index, searchDescendants,
                false, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("HeadetItemAt method - HeaderItem element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("HeaderItemAt method - HeaderItem element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("HeaderItemAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("HeaderItemAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_HeaderItem headerItem = new UIDA_HeaderItem(returnElement);
            return headerItem;
        }

        /// <summary>
        /// Searches a group item in the current element
        /// </summary>
        /// <param name="name">text of group item</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is case sensitive, default true</param>
        /// <returns>DataGridGroup element</returns>
        new public DataGridGroup Group(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_GroupControlTypeId,
                name, searchDescendants, false, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("Group method - Group element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("Group method - Group element not found");
                }
                else
                {
                    return null;
                }
            }

            DataGridGroup group = new DataGridGroup(returnElement);
            return group;
        }

        /// <summary>
        /// Searches for a Group item with a specified text at a specified index.
        /// </summary>
        /// <param name="name">text of group</param>
        /// <param name="index">index of group</param>
        /// <param name="searchDescendants">true if search through descendants, false if search only through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>DataGridGroup element</returns>
        new public DataGridGroup GroupAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("GroupAt method - index cannot be negative");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("GroupAt method - index cannot be negative");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_GroupControlTypeId, name, index, searchDescendants,
                false, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("GroupAt method - Group element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("GroupAt method - Group element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("GroupAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("GroupAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            DataGridGroup group = new DataGridGroup(returnElement);
            return group;
        }

        /// <summary>
        /// Searches a data item in the current element
        /// </summary>
        /// <param name="name">text of data item</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is case sensitive, default true</param>
        /// <returns>UIDA_DataItem element</returns>
        public UIDA_DataItem DataItem(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_DataItemControlTypeId,
                name, searchDescendants, false, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("DataItem method - DataItem element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("DataItem method - DataItem element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_DataItem dataItem = new UIDA_DataItem(returnElement);
            return dataItem;
        }

        /// <summary>
        /// Searches for a DataItem with a specified text at a specified index.
        /// </summary>
        /// <param name="name">text of DataItem</param>
        /// <param name="index">index of DataItem</param>
        /// <param name="searchDescendants">true if search through descendants, false if search only through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_DataItem element</returns>
        public UIDA_DataItem DataItemAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("DataItemAt method - index cannot be negative");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("DataItemAt method - index cannot be negative");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_DataItemControlTypeId, name, index, searchDescendants,
                false, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("DataItemAt method - DataItem element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("DataItemAt method - DataItem element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("DataItemAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("DataItemAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_DataItem dataItem = new UIDA_DataItem(returnElement);
            return dataItem;
        }
        
        /// <summary>
        /// Returns a collection of DataItem that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of DataItem elements, use null to return all DataItems</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_DataItem elements</returns>
        public UIDA_DataItem[] DataItems(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allDataItems = FindAll(UIA_ControlTypeIds.UIA_DataItemControlTypeId,
                name, searchDescendants, false, caseSensitive);

            List<UIDA_DataItem> dataitems = new List<UIDA_DataItem>();
            if (allDataItems != null)
            {
                foreach (IUIAutomationElement crtEl in allDataItems)
                {
                    dataitems.Add(new UIDA_DataItem(crtEl));
                }
            }
            return dataitems.ToArray();
        }
    }
}
