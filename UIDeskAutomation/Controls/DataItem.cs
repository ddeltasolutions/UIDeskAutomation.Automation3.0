using System;
using System.Collections.Generic;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
	/// <summary>
    /// This class represents a data item.
    /// </summary>
    public class UIDA_DataItem : ElementBase
    {
        internal int m_index = -1;
        internal IUIAutomationElement grid = null;
        
        /// <summary>
        /// Creates a UIDA_DataItem using an IUIAutomationElement
        /// </summary>
        /// <param name="el">UI Automation Element</param>
        public UIDA_DataItem(IUIAutomationElement el)
        {
			this.uiElement = el;
		}
        
        internal bool IsEqual(UIDA_DataItem item)
        {
            return Helper.CompareAutomationElements(uiElement, item.InnerElement);
        }

        /// <summary>
        /// Gets the selection state of a DataItem.
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
                
				UIDA_Header header = grid.Header;
				if (header == null)
				{
					throw new Exception("No header found");
				}
				
                UIDA_HeaderItem[] headerItems = header.Items;
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
}