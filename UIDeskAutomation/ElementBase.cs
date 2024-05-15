using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Base class for all UI controls. This class cannot be instantiated.
    /// </summary>
    abstract public partial class ElementBase
    {
        internal ElementBase() {}

        /// <summary>
        /// inner MS UI Automation element
        /// </summary>
        internal IUIAutomationElement uiElement = null;

        //wait period
        static internal int waitPeriod = 100;
		
		/// <summary>
        /// Gets the inner IUIAutomationElement of this element
        /// </summary>
		public IUIAutomationElement InnerElement
		{
			get
			{
				return this.uiElement;
			}
		}

        /// <summary>
        /// Gets a general description for this element
        /// </summary>
        /// <returns>description for the current element</returns>
        public string GetDescription()
        {
            if (this.uiElement == null)
            {
                Engine.TraceInLogFile("GetDescription method - Null element");
                return "Null element";
            }

            string sName = "";

            try
            {
                sName = this.uiElement.CurrentName;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

            string localizedCtrlType = "";
            try
            {
                localizedCtrlType = this.uiElement.CurrentLocalizedControlType;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

            return localizedCtrlType + ": " + sName;
        }

        /// <summary>
        /// Highlights an element for a second
        /// </summary>
        public void Highlight()
        {
            if (this.uiElement == null)
            {
                Engine.TraceInLogFile("Highlight method - Null element");
                throw new Exception("Highlight method - Null element");
            }

            tagRECT rect;
            try
            {
                rect = this.uiElement.CurrentBoundingRectangle;
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("Highlight method threw an exception - " +
                    ex.Message);

                return;
            }

            int left = rect.left;
            int top = rect.top;
            int width = rect.right - rect.left;
            int height = rect.bottom - rect.top;

            int thickness = 5;

            GraphicsPath path = new GraphicsPath();

            path.AddRectangle(new System.Drawing.Rectangle(0, 0, thickness, height)); //add left
            path.AddRectangle(new System.Drawing.Rectangle(thickness, height - thickness, 
                width - 2 * thickness, thickness)); //add bottom
            path.AddRectangle(new System.Drawing.Rectangle(width - thickness, 0, 
                thickness, height)); //add right
            path.AddRectangle(new System.Drawing.Rectangle(thickness, 0, 
                width - 2 * thickness, thickness)); //add top

            System.Drawing.Region region = new System.Drawing.Region(path);

            try
            {
                FrmDummy dummyForm = new FrmDummy();

                dummyForm.Left = left;
                dummyForm.Top = top;

                dummyForm.Width = width;
                dummyForm.Height = height;

                dummyForm.Region = region;

                dummyForm.ShowInTaskbar = false;
                dummyForm.TopMost = true;
                dummyForm.Visible = true;

                Thread.Sleep(200);
                dummyForm.Visible = false;
                Thread.Sleep(200);
                dummyForm.Visible = true;
                Thread.Sleep(200);
                dummyForm.Visible = false;
            }
            catch (Exception ex)
            { }
        }

        /// <summary>
        /// Tests if underlying UI object is still available. 
        /// For example if container window closes, the current UI element is not available to user anymore.
        /// </summary>
        public bool IsAlive
        {
            get
            {
                bool isAlive = true;

                try
                {
                    int processId = this.uiElement.CurrentProcessId;
                }
                catch //ElementNotAvailableException
                {
                    isAlive = false;
                }

                return isAlive;
            }
        }

        /// <summary>
        /// Gets the text of this element.
        /// </summary>
        /// <returns>text of the element</returns>
        public string GetText()
        {
            IUIAutomationValuePattern valuePattern = Helper.GetValuePattern(this.uiElement);

            if (valuePattern != null)
            {
                try
                {
                    return valuePattern.CurrentValue;
                }
                catch (Exception ex)
                {
                    Engine.TraceInLogFile("IUIAutomationValuePattern.CurrentValue: " + ex.Message);
                }
            }

			string name = null;
            try
            {
                name = this.uiElement.CurrentName;
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("AutomationElement.Name: " + ex.Message);
            }

            if (name != null)
            {
                return name;
            }

            IntPtr windowHandle = this.GetWindow();

            if (windowHandle != IntPtr.Zero)
            {
                UIDA_Window window = new UIDA_Window(windowHandle);
                return window.GetWindowText();
            }
            else
            {
                Engine.TraceInLogFile("GetText error - cannot get element text");
                throw new Exception("GetText error - cannot get element text");
            }
        }

        /// <summary>
        /// Gets the Id of the associated process
        /// </summary>
        public int ProcessId
        {
            get
            {
                IntPtr hWnd = this.GetWindow();
                if (hWnd == IntPtr.Zero)
                {
                    return 0;
                }

                uint processId = 0;
                UnsafeNativeFunctions.GetWindowThreadProcessId(hWnd, out processId);
                return (int)processId;
            }
        }
		
		/// <summary>
        /// Gets the UI Automation Class Name of this element
        /// </summary>
        public string ClassName
        {
            get
            {
                return this.uiElement.CurrentClassName;
            }
        }
		
		/// <summary>
        /// Gets the Left screen coordinate of the element or -1 if the element doesn't have a visual representation.
        /// </summary>
		public int Left
		{
			get
			{
				try
				{
					return this.uiElement.CurrentBoundingRectangle.left;
				}
				catch { }

				return -1;
			}
		}
		
		/// <summary>
        /// Gets the Top screen coordinate of the element or -1 if the element doesn't have a visual representation.
        /// </summary>
		public int Top
		{
			get
			{
				try
				{
					return this.uiElement.CurrentBoundingRectangle.top;
				}
				catch { }
		
				return -1;
			}
		}
		
		/// <summary>
        /// Gets the Width of the element bounding rectangle or -1 if the element doesn't have a visual representation.
        /// </summary>
		public int Width
		{
			get
			{
				try
				{
					tagRECT rect = this.uiElement.CurrentBoundingRectangle;
					return (rect.right - rect.left);
				}
				catch { }

				return -1;
			}
		}
		
		/// <summary>
        /// Gets the Height of the element bounding rectangle or -1 if the element doesn't have a visual representation.
        /// </summary>
		public int Height
		{
			get
			{
				try
				{
					tagRECT rect = this.uiElement.CurrentBoundingRectangle;
					return (rect.bottom - rect.top);
				}
				catch { }

				return -1;
			}
		}
		
		/// <summary>
        /// Gets the parent of the element. You need to cast the parent to the proper type. For example, if the current 
		/// element is a list item you can write: "UIDA_List list = listItem.Parent as UIDA_List;"
        /// </summary>
		public ElementBase Parent
		{
			get
			{
				IUIAutomationTreeWalker tw = Engine.uiAutomation.ControlViewWalker;
				IUIAutomationElement parent = tw.GetParentElement(this.uiElement);
				
				if (parent == null)
				{
					return null;
				}
				
				if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_ButtonControlTypeId)
				{
					return new UIDA_Button(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_CalendarControlTypeId)
				{
					return new UIDA_Calendar(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_CheckBoxControlTypeId)
				{
					return new UIDA_CheckBox(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_ComboBoxControlTypeId)
				{
					return new UIDA_ComboBox(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_HyperlinkControlTypeId)
				{
					return new UIDA_HyperLink(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_ImageControlTypeId)
				{
					return new UIDA_Image(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_ListItemControlTypeId)
				{
					return new UIDA_ListItem(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_ListControlTypeId)
				{
					return new UIDA_List(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_MenuControlTypeId)
				{
					return new UIDA_TopLevelMenu(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_MenuBarControlTypeId)
				{
					return new UIDA_MenuBar(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_MenuItemControlTypeId)
				{
					return new UIDA_MenuItem(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_ProgressBarControlTypeId)
				{
					return new UIDA_ProgressBar(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_RadioButtonControlTypeId)
				{
					return new UIDA_RadioButton(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_ScrollBarControlTypeId)
				{
					return new UIDA_ScrollBar(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_SliderControlTypeId)
				{
					return new UIDA_Slider(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_SpinnerControlTypeId)
				{
					return new UIDA_Spinner(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_StatusBarControlTypeId)
				{
					return new UIDA_StatusBar(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_TabControlTypeId)
				{
					return new UIDA_TabCtrl(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_TabItemControlTypeId)
				{
					return new UIDA_TabItem(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_TextControlTypeId)
				{
					return new UIDA_Label(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_ToolBarControlTypeId)
				{
					return new UIDA_ToolBar(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_ToolTipControlTypeId)
				{
					return new UIDA_ToolTip(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_TreeControlTypeId)
				{
					return new UIDA_Tree(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_TreeItemControlTypeId)
				{
					return new UIDA_TreeItem(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_CustomControlTypeId)
				{
					return new UIDA_Custom(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_GroupControlTypeId)
				{
					return new UIDA_Group(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_ThumbControlTypeId)
				{
					return new UIDA_Thumb(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_DataGridControlTypeId)
				{
					return new UIDA_DataGrid(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_DataItemControlTypeId)
				{
					return new UIDA_DataItem(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_DocumentControlTypeId)
				{
					return new UIDA_Document(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_SplitButtonControlTypeId)
				{
					return new UIDA_SplitButton(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_WindowControlTypeId)
				{
					return new UIDA_Window(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_PaneControlTypeId)
				{
					return new UIDA_Pane(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_HeaderControlTypeId)
				{
					return new UIDA_Header(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_HeaderItemControlTypeId)
				{
					return new UIDA_HeaderItem(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_TableControlTypeId)
				{
					return new UIDA_Table(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_TitleBarControlTypeId)
				{
					return new UIDA_TitleBar(parent);
				}
				else if (parent.CurrentControlType == UIA_ControlTypeIds.UIA_SeparatorControlTypeId)
				{
					return new UIDA_Separator(parent);
				}
				
				return null;
			}
		}
		
		/// <summary>
        /// Gets a boolean to determine if the UI element is enabled.
        /// </summary>
		public bool IsEnabled
		{
			get
			{
				try
				{
					return (this.uiElement.CurrentIsEnabled != 0);
				}
				catch 
				{
					return true;
				}
			}
		}
    }
}
