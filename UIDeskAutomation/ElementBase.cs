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
            {
                return;
            }
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
    }
}
