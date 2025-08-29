using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a Generic object. Used by search functions like FindFirst() and FindAll().
    /// </summary>
    public class UIDA_Generic : ElementBase
    {
        public UIDA_Generic(IUIAutomationElement el)
        {
            this.uiElement = el;
        }

        public UIDA_Button AsButton()
        {
            return new UIDA_Button(uiElement);
        }

        public UIDA_Calendar AsCalendar()
        {
            return new UIDA_Calendar(uiElement);
        }

        public UIDA_CheckBox AsCheckBox()
        {
            return new UIDA_CheckBox(uiElement);
        }

        public UIDA_Custom AsCustom()
        {
            return new UIDA_Custom(uiElement);
        }

        public UIDA_DataGrid AsDataGrid()
        {
            return new UIDA_DataGrid(uiElement);
        }

        public UIDA_DataItem AsDataItem()
        {
            return new UIDA_DataItem(uiElement);
        }

        public UIDA_DatePicker AsDatePicker()
        {
            return new UIDA_DatePicker(uiElement);
        }

        public UIDA_Document AsDocument()
        {
            return new UIDA_Document(uiElement);
        }

        public UIDA_Edit AsEdit()
        {
            return new UIDA_Edit(uiElement);
        }

        public UIDA_Group AsGroup()
        {
            return new UIDA_Group(uiElement);
        }

        public UIDA_Header AsHeader()
        {
            return new UIDA_Header(uiElement);
        }

        public UIDA_HeaderItem AsHeaderItem()
        {
            return new UIDA_HeaderItem(uiElement);
        }

        public UIDA_HyperLink AsHyperLink()
        {
            return new UIDA_HyperLink(uiElement);
        }

        public UIDA_Image AsImage()
        {
            return new UIDA_Image(uiElement);
        }

        public UIDA_Text AsText()
        {
            return new UIDA_Text(uiElement);
        }

        public UIDA_List AsList()
        {
            return new UIDA_List(uiElement);
        }

        public UIDA_ListItem AsListItem()
        {
            return new UIDA_ListItem(uiElement);
        }

        public UIDA_MenuBar AsMenuBar()
        {
            return new UIDA_MenuBar(uiElement);
        }

        public UIDA_MenuItem AsMenuItem()
        {
            return new UIDA_MenuItem(uiElement);
        }

        public UIDA_Pane AsPane()
        {
            return new UIDA_Pane(uiElement);
        }

        public UIDA_ProgressBar AsProgressBar()
        {
            return new UIDA_ProgressBar(uiElement);
        }

        public UIDA_RadioButton AsRadioButton()
        {
            return new UIDA_RadioButton(uiElement);
        }

        public UIDA_ScrollBar AsScrollBar()
        {
            return new UIDA_ScrollBar(uiElement);
        }

        public UIDA_Separator AsSeparator()
        {
            return new UIDA_Separator(uiElement);
        }

        public UIDA_Slider AsSlider()
        {
            return new UIDA_Slider(uiElement);
        }

        public UIDA_Spinner AsSpinner()
        {
            return new UIDA_Spinner(uiElement);
        }

        public UIDA_SplitButton AsSplitButton()
        {
            return new UIDA_SplitButton(uiElement);
        }

        public UIDA_StatusBar AsStatusBar()
        {
            return new UIDA_StatusBar(uiElement);
        }

        public UIDA_TabCtrl AsTabCtrl()
        {
            return new UIDA_TabCtrl(uiElement);
        }

        public UIDA_TabItem AsTabItem()
        {
            return new UIDA_TabItem(uiElement);
        }

        public UIDA_Table AsTable()
        {
            return new UIDA_Table(uiElement);
        }

        public UIDA_Thumb AsThumb()
        {
            return new UIDA_Thumb(uiElement);
        }

        public UIDA_TitleBar AsTitleBar()
        {
            return new UIDA_TitleBar(uiElement);
        }

        public UIDA_ToolBar AsToolBar()
        {
            return new UIDA_ToolBar(uiElement);
        }

        public UIDA_ToolTip AsToolTip()
        {
            return new UIDA_ToolTip(uiElement);
        }

        public UIDA_Menu AsMenu()
        {
            return new UIDA_Menu(uiElement);
        }

        public UIDA_Tree AsTree()
        {
            return new UIDA_Tree(uiElement);
        }

        public UIDA_TreeItem AsTreeItem()
        {
            return new UIDA_TreeItem(uiElement);
        }
    }
}