using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Class that represents a Window ui element.
    /// </summary>
    public class UIDA_Window: ElementBase
    {
        private IntPtr hWnd = IntPtr.Zero;

        /// <summary>
        /// Creates a UIDA_Window using a window handle
        /// </summary>
        /// <param name="hwnd">Window handle</param>
        public UIDA_Window(IntPtr hwnd)
        {
            this.hWnd = hwnd;

            try
            {
                base.uiElement = Engine.uiAutomation.ElementFromHandle(this.hWnd);
            }
            catch { }
        }

        /// <summary>
        /// Creates a UIDA_Window using an IUIAutomationElement
        /// </summary>
        /// <param name="el">UI Automation Element</param>
        public UIDA_Window(IUIAutomationElement el)
        {
            base.uiElement = el;
            try
            {
            	this.hWnd = el.CurrentNativeWindowHandle;
            }
            catch {}
        }

        /// <summary>
        /// Gets a window description
        /// </summary>
        /// <returns></returns>
        public string WindowDescription()
        {
            if (this.hWnd == IntPtr.Zero)
            {
                Engine.TraceInLogFile("WindowDescription method - null window handle");
                return "null window handle";
            }

            StringBuilder sBuilderClassName = new StringBuilder(256);
            UnsafeNativeFunctions.GetClassName(this.hWnd, sBuilderClassName, 256);

            StringBuilder sBuilderWindowText = new StringBuilder(256);
            UnsafeNativeFunctions.GetWindowText(this.hWnd, sBuilderWindowText, 256);

            return "Class name: \"" + sBuilderClassName.ToString() + 
                "\", Window text: \"" + sBuilderWindowText.ToString() + 
                "\", hWnd=0x" + this.hWnd.ToString("X");
        }

        /// <summary>
        /// Shows a window. It acts like Restore(). This function doesn't bring the window in foreground. 
		/// If you want the window in foreground call BringToForeground().
        /// </summary>
        public void Show()
        {
            if (hWnd == IntPtr.Zero)
            {
                Engine.TraceInLogFile("Show method - null window handle");
                return;
            }

            UnsafeNativeFunctions.ShowWindow(hWnd, WindowShowStyle.ShowNormal);
        }

        /// <summary>
        /// Minimizes a window
        /// </summary>
        public void Minimize()
        {
            if (hWnd == IntPtr.Zero)
            {
                Engine.TraceInLogFile("Minimize method - null window handle");
                return;
            }

            UnsafeNativeFunctions.ShowWindow(hWnd, WindowShowStyle.Minimize);
        }

        /// <summary>
        /// Maximizes a window
        /// </summary>
        public void Maximize()
        {
            if (hWnd == IntPtr.Zero)
            {
                Engine.TraceInLogFile("Maximize method - null window handle");
                return;
            }

            UnsafeNativeFunctions.ShowWindow(hWnd, WindowShowStyle.Maximize);
        }

        /// <summary>
        /// Restores a window in restore position.
        /// </summary>
        public void Restore()
        {
            if (hWnd == IntPtr.Zero)
            {
                Engine.TraceInLogFile("Restore method - null window handle");
                return;
            }

            UnsafeNativeFunctions.ShowWindow(hWnd, WindowShowStyle.Restore);
        }

        /// <summary>
        /// Closes a window
        /// </summary>
        public void Close()
        {
            if (hWnd == IntPtr.Zero)
            {
                Engine.TraceInLogFile("Close method - null window handle");
                return;
            }

            UnsafeNativeFunctions.SendMessage(hWnd, WindowMessages.WM_CLOSE, 
                IntPtr.Zero, IntPtr.Zero);

            int timeOut = Engine.GetInstance().Timeout;

            while (timeOut > 0)
            {
                if (!UnsafeNativeFunctions.IsWindow(hWnd))
                {
                    break;
                }

                timeOut -= ElementBase.waitPeriod;
                Thread.Sleep(ElementBase.waitPeriod);
            }

            if (timeOut <= 0)
            {
                Engine.TraceInLogFile("Window.Close() method - window could not be closed, maybe a confirmation is required");
            }
        }

        /// <summary>
        /// Determines if a window is minimized
        /// </summary>
        public bool IsMinimized
        { 
            get
            {
                if ((hWnd == IntPtr.Zero) || !UnsafeNativeFunctions.IsWindow(hWnd))
                {
                    throw new Exception("IsMinimized() - not a valid window");
                }

                return UnsafeNativeFunctions.IsIconic(hWnd);
            }
        }

        /// <summary>
        /// Determines if a window is maximized
        /// </summary>
        public bool IsMaximized
        {
            get
            {
                if ((hWnd == IntPtr.Zero) || !UnsafeNativeFunctions.IsWindow(hWnd))
                {
                    throw new Exception("IsMaximized() - not a valid window");
                }

                WINDOWPLACEMENT placement = new WINDOWPLACEMENT();
                placement.length = System.Runtime.InteropServices.Marshal.SizeOf(placement);
                if (UnsafeNativeFunctions.GetWindowPlacement(hWnd, ref placement) == false)
                {
                    return false;
                }

                return (placement.showCmd == WindowPlacementConstants.SW_SHOWMAXIMIZED);
            }
        }

        /// <summary>
        /// Brings window to foreground
        /// </summary>
        new public void BringToForeground()
        {
            if (hWnd == IntPtr.Zero)
            {
                Engine.TraceInLogFile("BringToForeground method - null window handle");

                return;
            }

            UnsafeNativeFunctions.SetForegroundWindow(hWnd);
        }

        /// <summary>
        /// Moves this window to specified x and y screen coordinates.
        /// </summary>
        /// <param name="x">x coordinate</param>
        /// <param name="y">y coordinate</param>
        public void Move(int x, int y)
        {
            if (this.hWnd == IntPtr.Zero)
            {
                Engine.TraceInLogFile("Move method - null window handle");
                throw new Exception("Move method - null window handle");
            }

            RECT rect = new RECT(0, 0, 0, 0);
            UnsafeNativeFunctions.GetWindowRect(this.hWnd, out rect);

            int width = rect.Right - rect.Left;
            int height = rect.Bottom - rect.Top;

            UnsafeNativeFunctions.MoveWindow(this.hWnd, x, y, width, height, true);
        }

        /// <summary>
        /// Moves a window relatively with horizontal and vertical offsets
        /// </summary>
        /// <param name="horizontalOffset">horizontal offset</param>
        /// <param name="verticalOffset">vertical offset</param>
        public void MoveOffset(int horizontalOffset, int verticalOffset)
        {
			UIDA_Point pt = GetPosition();
			int newX = pt.X + horizontalOffset;
			int newY = pt.Y + verticalOffset;

            this.Move(newX, newY);
        }

        /// <summary>
        /// Sets both width and height of this window
        /// </summary>
        /// <param name="width">width in pixels</param>
        /// <param name="height">height in pixels</param>
        public void Resize(int width, int height)
        {
            if (this.hWnd == IntPtr.Zero)
            {
                Engine.TraceInLogFile("Resize method - null window handle");
                throw new Exception("Resize method - null window handle");
            }

            RECT rect = new RECT(0, 0, 0, 0);
            UnsafeNativeFunctions.GetWindowRect(this.hWnd, out rect);

            int x = rect.Left;
            int y = rect.Top;

            UnsafeNativeFunctions.MoveWindow(this.hWnd, x, y, width, height, true);
        }

        /// <summary>
        /// Gets upper-left corner position of this window in screen coordinates
        /// </summary>
        /// <returns>Point position</returns>
        public UIDA_Point GetPosition()
        {
            UIDA_Point point = new UIDA_Point(0, 0);

            if (this.hWnd == IntPtr.Zero)
            {
                Engine.TraceInLogFile("GetPosition method - null window handle");
                throw new Exception("GetPosition method - null window handle");
            }

            RECT rect = new RECT(0, 0, 0, 0);
            UnsafeNativeFunctions.GetWindowRect(this.hWnd, out rect);

            point.X = rect.Left;
            point.Y = rect.Top;

            return point;
        }

        /// <summary>
        /// gets/sets the width of this window
        /// </summary>
        public int WindowWidth
        {
            get
            {
                if (this.hWnd == IntPtr.Zero)
                {
                    Engine.TraceInLogFile("Width property (get) - null window handle");
                    throw new Exception("Width property (get) - null window handle");
                }

                RECT rect = new RECT(0, 0, 0, 0);
                UnsafeNativeFunctions.GetWindowRect(this.hWnd, out rect);

                int width = rect.Right - rect.Left;
                return width;
            }

            set
            {
                if (this.hWnd == IntPtr.Zero)
                {
                    Engine.TraceInLogFile("Width property (set) - null window handle");
                    throw new Exception("Width property (set) - null window handle");
                }

                RECT rect = new RECT(0, 0, 0, 0);
                UnsafeNativeFunctions.GetWindowRect(this.hWnd, out rect);

                int x = rect.Left;
                int y = rect.Top;

                int height = rect.Bottom - rect.Top;

                UnsafeNativeFunctions.MoveWindow(hWnd, x, y, value, height, true);
            }
        }

        /// <summary>
        /// gets/sets the window height
        /// </summary>
        public int WindowHeight
        {
            get
            {
                if (this.hWnd == IntPtr.Zero)
                {
                    Engine.TraceInLogFile("Height property (get) - null window handle");
                    throw new Exception("Height property (get) - null window handle");
                }

                RECT rect = new RECT(0, 0, 0, 0);
                UnsafeNativeFunctions.GetWindowRect(this.hWnd, out rect);

                int height = rect.Bottom - rect.Top;
                return height;
            }

            set
            {
                if (this.hWnd == IntPtr.Zero)
                {
                    Engine.TraceInLogFile("Height property (set) - null window handle");
                    throw new Exception("Height property (set) - null window handle");
                }

                RECT rect = new RECT(0, 0, 0, 0);
                UnsafeNativeFunctions.GetWindowRect(this.hWnd, out rect);

                int x = rect.Left;
                int y = rect.Top;

                int width = rect.Right - rect.Left;

                UnsafeNativeFunctions.MoveWindow(this.hWnd, x, y, width, value, true);
            }
        }

        /// <summary>
        /// Gets the window text
        /// </summary>
        /// <returns>window text</returns>
        public string GetWindowText()
        {
            StringBuilder sBuilderWindowText = new StringBuilder(256);

            int res = UnsafeNativeFunctions.GetWindowText(this.hWnd, 
                sBuilderWindowText, 256);

            if (res == 0)
            {
                IntPtr textLengthPtr = UnsafeNativeFunctions.SendMessage(this.hWnd,
                    WindowMessages.WM_GETTEXTLENGTH, IntPtr.Zero, IntPtr.Zero);

                if (textLengthPtr.ToInt32() > 0)
                {
                    int textLength = textLengthPtr.ToInt32() + 1;

                    StringBuilder text = new StringBuilder(textLength);

                    UnsafeNativeFunctions.SendMessage(this.hWnd,
                        WindowMessages.WM_GETTEXT, textLength, text);

                    return text.ToString();
                }
                else
                {
                    return string.Empty;
                }
            }
            else
            {
                return sBuilderWindowText.ToString();
            }
        }
    }
	
	public class UIDA_Point
	{
		public int X { get; set; }
		public int Y { get; set; }
		
		public UIDA_Point()
		{
			X = 0;
			Y = 0;
		}
		
		public UIDA_Point(int x, int y)
		{
			X = x;
			Y = y;
		}
	}
}
