using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using System.Drawing;
using System.Drawing.Imaging;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
	/// <summary>
    /// Namespace for all classes.
    /// </summary>
	public static class NamespaceDoc
	{}

    /// <summary>
    /// Main entry point class. You need only one instance of this class. 
    /// Information on how to use this library can be found at 
    /// <a href="http://automationspy.freecluster.eu/uideskautomation_unmanaged.html">http://automationspy.freecluster.eu/uideskautomation_unmanaged.html</a>.
    /// </summary>
    public class Engine
    {
        private static int wait = 10000; //default wait 10 seconds
        private static string logFileName = string.Empty;
        private static bool throwExceptionsWhenSearch = false;

        private static Engine instance = null;

        /// <summary>
        /// Constructor
        /// </summary>
        public Engine()
        {
            Engine.instance = this;
            return;
        }
        
        internal static CUIAutomation uiAutomation = new CUIAutomation();
		//public static bool IsCancelled = false;

        /// <summary>
        /// Gets/Sets Timeout period (in milliseconds)
        /// </summary>
        public int Timeout
        {
            get
            {
                return Engine.wait;
            }
            set
            {
                Engine.wait = value;
            }
        }

        /// <summary>
        /// Gets/Sets Log File path
        /// </summary>
        public string LogFile
        { 
            get
            {
                return Engine.logFileName;
            }
            set
            {
                Engine.logFileName = value;

                try
                {
                    if (File.Exists(Engine.logFileName))
                    {
                        //clear log file content
                        File.WriteAllText(Engine.logFileName, string.Empty);
                    }
                }
                catch { }
            }
        }

        internal static bool ThrowExceptionsWhenSearch
        {
            get
            {
                return Engine.throwExceptionsWhenSearch;
            }
            set
            {
                Engine.throwExceptionsWhenSearch = value;
            }
        }

        /// <summary>
        /// Gets/Sets whether exceptions should be raised when elements are not found.
        /// </summary>
        public bool ThrowExceptionsForSearchFunctions
        {
            get
            {
                return Engine.throwExceptionsWhenSearch;
            }
            set
            {
                Engine.throwExceptionsWhenSearch = value;
            }
        }

        /// <summary>
        /// Writes a message in the log file
        /// </summary>
        /// <param name="sMessage">message</param>
        internal static void TraceInLogFile(string sMessage)
        {
            if (Engine.logFileName.Length == 0)
            {
                return;
            }

            try
            {
                string sLineToWrite = DateTime.Now.ToString("G") + ": " + sMessage + Environment.NewLine;
                File.AppendAllText(Engine.logFileName, sLineToWrite);
            }
            catch { }
        }

        /// <summary>
        /// Writes a message in the log file
        /// </summary>
        /// <param name="sMessage">message to write in log file</param>
        public void WriteInLogFile(string sMessage)
        {
            Engine.TraceInLogFile(sMessage);
        }

        /// <summary>
        /// Gets a top-level window element given its window text and/or class name
        /// </summary>
        /// <param name="windowText">Window text (optional), wildcards can be used</param>
        /// <param name="className">Window class name (optional)</param>
        /// <param name="index">index (optional)</param>
        /// <param name="caseSensitive">searches the window text case sensitive, default true</param>
        /// <returns>UIDA_Window</returns>
        public UIDA_Window GetTopLevel(string windowText, string className = null, int index = 0,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("GetTopLevel method - index must be positive number");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("GetTopLevel method - index must be positive number");
                }
                else
                {
                    return null;
                }
            }

            /*if (index == 0)
            {
                index = 1;
            }*/

            int nWaitMs = Engine.wait;
            IntPtr hwnd = IntPtr.Zero;

            Errors error = Errors.None;

			while (true)
            {
                error = Helper.GetWindowAt(IntPtr.Zero, className, windowText, 
                    index, caseSensitive, out hwnd);

                if ((hwnd != IntPtr.Zero) && (error == Errors.None))
                {
                    break; //found!
                }
				
				/*if (Engine.IsCancelled == true)
				{
					Engine.IsCancelled = false;
					break;
				}*/

                nWaitMs -= ElementBase.waitPeriod; //wait 100 milliseconds
				if (nWaitMs <= 0)
				{
					break;
				}
                Thread.Sleep(ElementBase.waitPeriod);
            }

            if (hwnd == IntPtr.Zero)
            {
                if (error == Errors.IndexTooBig)
                {
                    Engine.TraceInLogFile("GetTopLevel method - Index too big");

                    if (Engine.ThrowExceptionsWhenSearch == true)
                    {
                        throw new Exception("GetTopLevel method - Index too big");
                    }
                    else
                    {
                        return null;
                    }
                }
                else
                {
                    Engine.TraceInLogFile("GetTopLevel method - Window not found");

                    if (Engine.ThrowExceptionsWhenSearch == true)
                    {
                        throw new Exception("GetTopLevel method - Window not found");
                    }
                    else
                    {
                        return null;
                    }
                }
            }

            return new UIDA_Window(hwnd);
        }

        /// <summary>
        /// Gets the top level window of the specified process with the specified 
        /// window text, class name and index
        /// </summary>
        /// <param name="processName">process name (ex: "myprocess.exe")</param>
        /// <param name="windowText">Window text, wildcards can be used</param>
        /// <param name="className">Window class name</param>
        /// <param name="index">index</param>
        /// <param name="caseSensitive">searches the window text case sensitive, default true</param>
        /// <returns>Top level UIDA_Window</returns>
        public UIDA_Window GetTopLevelByProcName(string processName, 
            string windowText = null, string className = null, int index = 0, bool caseSensitive = true)
        {
            int nWaitMs = Engine.wait;
            IntPtr windowHandle = IntPtr.Zero;

            if (index < 0)
            { 
                Engine.TraceInLogFile("GetTopLevelByProcName failed: negative index");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("GetTopLevelByProcName failed - negative index");
                }
                else
                {
                    return null;
                }
            }

            if (index == 0)
            {
                index = 1;
            }

			while (true)
            {
                windowHandle = Helper.GetTopLevelByProcName(processName, 
                    className, windowText, index, caseSensitive);

                if (windowHandle != IntPtr.Zero)
                {
                    break; //found!
                }
				
				/*if (Engine.IsCancelled == true)
				{
					Engine.IsCancelled = false;
					break;
				}*/

                nWaitMs -= ElementBase.waitPeriod; //wait 100 milliseconds
				if (nWaitMs <= 0)
				{
					break;
				}
                Thread.Sleep(ElementBase.waitPeriod);
            }

            if (windowHandle == IntPtr.Zero)
            {
                Engine.TraceInLogFile("GetTopLevelByProcName - window not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("GetTopLevelByProcName method - window not found");
                }
                else
                {
                    return null;
                }
            }

            return new UIDA_Window(windowHandle);
        }

        /// <summary>
        /// Get the toplevel windows with the specified window text and/or class name.
        /// </summary>
        /// <param name="windowText">Window text, wildcards can be used</param>
        /// <param name="className">Window class name</param>
        /// <param name="caseSensitive">searches the windows text case sensitive, default true</param>
        /// <returns>UIDA_Window collection</returns>
        public UIDA_Window[] GetTopLevelWindows(
            string windowText = null, string className = null, bool caseSensitive = true)
        {
            IntPtr[] handles = Helper.GetWindows(IntPtr.Zero, className, 
                windowText, caseSensitive);

            List<UIDA_Window> windows = new List<UIDA_Window>();

            foreach (IntPtr handle in handles)
            {
                UIDA_Window wnd = new UIDA_Window(handle);
                windows.Add(wnd);
            }

            return windows.ToArray();
        }

        /// <summary>
        /// Get the toplevel window of the specified process with the specified window text and/or class name.
        /// </summary>
        /// <param name="processId">Process Id</param>
        /// <param name="windowText">Window text, wildcards can be used</param>
        /// <param name="className">Window class name</param>
        /// <param name="caseSensitive">searches the window text case sensitive, default true</param>
        /// <returns>UIDA_Window</returns>
        public UIDA_Window GetTopLevelByProcId(int processId, 
            string windowText = null, string className = null, bool caseSensitive = true)
        {
            Process process = null;

            try
            {
                process = Process.GetProcessById(processId);
            }
            catch { }

            if (process == null)
            {
                Engine.TraceInLogFile("GetTopLevelByProcId method - process with this id doesn't exist");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("GetTopLevelByProcId method - process with this id doesn't exist");
                }
                else
                {
                    return null;
                }
            }

            int nWaitMs = Engine.wait;
            IntPtr windowHandle = IntPtr.Zero;

			while (true)
            {
                windowHandle = Helper.GetTopLevelByProcId(processId, className, 
                    windowText, caseSensitive);

                if (windowHandle != IntPtr.Zero)
                {
                    break; //found!
                }
				
				/*if (Engine.IsCancelled == true)
				{
					Engine.IsCancelled = false;
					break;
				}*/

                nWaitMs -= ElementBase.waitPeriod; //wait 100 milliseconds
				if (nWaitMs <= 0)
				{
					break;
				}
                Thread.Sleep(ElementBase.waitPeriod);
            }

            if (windowHandle == IntPtr.Zero)
            {
                Engine.TraceInLogFile("GetTopLevelByProcId method - window not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("GetTopLevelByProcId method - window not found");
                }
                else
                {
                    return null;
                }
            }

            return new UIDA_Window(windowHandle);
        }

        private IntPtr FindWindow(string sClassName, string sText, int nIndex = 0)
        {
            IntPtr hwnd = IntPtr.Zero;

            if (nIndex <= 1)
            {
                hwnd = UnsafeNativeFunctions.FindWindowEx(IntPtr.Zero, IntPtr.Zero, sClassName, sText);
            }
            else
            {
                IntPtr hwndTemp = IntPtr.Zero;

                while (nIndex >= 1)
                {
                    hwnd = UnsafeNativeFunctions.FindWindowEx(IntPtr.Zero, hwndTemp, sClassName, sText);

                    if (hwnd == IntPtr.Zero)
                    {
                        return IntPtr.Zero;
                    }

                    hwndTemp = hwnd;
                    nIndex--;
                }
            }

            return hwnd;
        }

        /// <summary>
        /// Gets a Pane ui element representing the main desktop.
        /// </summary>
        /// <returns>a UIDA_Pane element that represents the desktop</returns>
        public UIDA_Pane GetDesktopPane()
        {
            IUIAutomationElement desktopPane = uiAutomation.GetRootElement();
            return new UIDA_Pane(desktopPane);
        }

        /// <summary>
        /// Sends keystrokes to the window which currently has the focus.
		/// This function behaves like the .NET 'SendKeys.Send(string)' function.
		/// More details can be found here: https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.sendkeys.send?view=netframework-4.8
        /// </summary>
        /// <param name="text">keys to send</param>
        public void SendKeys(string text)
        {
            System.Windows.Forms.SendKeys.SendWait(text);
        }
        
        /// <summary>
        /// This function presses a key (without releasing it).
        /// </summary>
        /// <param name="key">key to press</param>
        public void KeyDown(VirtualKeys key)
        {
            SendInputClass.KeyDown(key);
        }
        
        /// <summary>
        /// This function presses and releases a key.
        /// </summary>
        /// <param name="key">key to press and release</param>
        public void KeyPress(VirtualKeys key)
        {
            SendInputClass.KeyDown(key);
            SendInputClass.KeyUp(key);
        }
        
        /// <summary>
        /// This function presses and releases multiple keys.
        /// </summary>
        /// <param name="keys">keys to press and release</param>
        public void KeysPress(VirtualKeys[] keys)
        {
            foreach (VirtualKeys key in keys)
            {
                SendInputClass.KeyDown(key);
                SendInputClass.KeyUp(key);
            }
        }
        
        /// <summary>
        /// This function releases a key.
        /// </summary>
        /// <param name="key">key to release</param>
        public void KeyUp(VirtualKeys key)
        {
            SendInputClass.KeyUp(key);
        }

        /// <summary>
        /// Simulates keystrokes. 
		/// This function behaves like the .NET 'SendKeys.Send(string)' function.
        /// </summary>
        /// <param name="text">keys to send</param>
        public void SimulateSendKeys(string text)
        {
            /////// Get focused window from foreground window
            IntPtr hwndForeground = UnsafeNativeFunctions.GetForegroundWindow();
            if (hwndForeground == IntPtr.Zero)
            {
                Engine.TraceInLogFile("SimulateSendKeys failed: Cannot get foreground window.");
                throw new Exception("SimulateSendKeys failed: Cannot get foreground window.");
            }

            uint processId = 0;
            uint foregroundThreadId = 
                UnsafeNativeFunctions.GetWindowThreadProcessId(hwndForeground, out processId);

            uint currentThreadId = UnsafeNativeFunctions.GetCurrentThreadId();

            bool bResult = UnsafeNativeFunctions.AttachThreadInput(
                currentThreadId, foregroundThreadId, true);
            Debug.Assert(bResult == true);

            IntPtr hwnd = UnsafeNativeFunctions.GetFocus();

            bResult = UnsafeNativeFunctions.AttachThreadInput(
                currentThreadId, foregroundThreadId, false);
            Debug.Assert(bResult == true);
            ///////

            if (hwnd == IntPtr.Zero)
            {
                hwnd = hwndForeground;
            }

            if (hwnd == IntPtr.Zero)
            {
                Engine.TraceInLogFile("SimulateSendKeys - Cannot get the active window");
                throw new Exception("SimulateSendKeys - Cannot get the active window");
            }

            Helper.SimulateSendKeys(text, hwnd);
        }
        

        // x, y - screen coordinates  
        // mouse button, 1 - left button, 2 - right button, 3 - middle button
        //public void ClickScreenCoordinatesAt(int x, int y, int mouseButton = 1)
        /// <summary>
        /// Left mouse click at screen coordinates
        /// </summary>
        /// <param name="x">x - screen coordinate</param>
        /// <param name="y">y - screen coordinate</param>
        /// <param name="keys">keys pressed, 0 - None, 1 - Control pressed, 
        /// 2 - Shift pressed, 3 - Both Control and Shift pressed</param>
        public void ClickScreenCoordinatesAt(int x, int y, int keys = 0)
        {
            // left click
            SendInputClass.ClickLeftMouseButton(x, y, keys);
        }

        // x, y - screen coordinates  
        /// <summary>
        /// Right mouse button click at specified screen coordinates
        /// </summary>
        /// <param name="x">x - screen coordinate</param>
        /// <param name="y">y - screen coordinate</param>
        /// <param name="keys">keys pressed, 0 - None, 1 - Control pressed, 
        /// 2 - Shift pressed, 3 - Both Control and Shift pressed</param>
        public void RightClickScreenCoordinatesAt(int x, int y, int keys = 0)
        {
            // right click
            SendInputClass.ClickRightMouseButton(x, y, keys);
        }

        // x, y - screen coordinates  
        /// <summary>
        /// Middle mouse button click at specified screen coordinates
        /// </summary>
        /// <param name="x">x - screen coordinate</param>
        /// <param name="y">y - screen coordinate</param>
        /// <param name="keys">keys pressed, 0 - None, 1 - Control pressed, 
        /// 2 - Shift pressed, 3 - Both Control and Shift pressed</param>
        public void MiddleClickScreenCoordinatesAt(int x, int y, int keys = 0)
        {
            // middle button click
            SendInputClass.ClickMiddleMouseButton(x, y, keys);
        }

        // x, y - screen coordinates 
        /// <summary>
        /// Left mouse button double click at specified screen coordinates
        /// </summary>
        /// <param name="x">x - screen coordinate</param>
        /// <param name="y">y - screen coordinate</param>
        /// <param name="keys">keys pressed, 0 - None, 1 - Control pressed, 
        /// 2 - Shift pressed, 3 - Both Control and Shift pressed</param>
        public void DoubleClickAt(int x, int y, int keys = 0)
        {
            SendInputClass.DoubleClick(x, y, keys);
        }
		
		/// <summary>
        /// Presses the left mouse button, without releasing it. 
		/// It can be used in combination with MoveMouse and LeftMouseButtonUp for drag and drop operations.
        /// </summary>
        public void LeftMouseButtonDown()
        {
            SendInputClass.LeftMouseButtonDown();
			Sleep(100);
        }
		
		/// <summary>
        /// Releases the left mouse button
        /// </summary>
        public void LeftMouseButtonUp()
        {
            SendInputClass.LeftMouseButtonUp();
			Sleep(100);
        }
		
		/// <summary>
        /// Presses the right mouse button, without releasing it
        /// </summary>
        public void RightMouseButtonDown()
        {
            SendInputClass.RightMouseButtonDown();
			Sleep(100);
        }
		
		/// <summary>
        /// Releases the right mouse button
        /// </summary>
        public void RightMouseButtonUp()
        {
            SendInputClass.RightMouseButtonUp();
			Sleep(100);
        }

        // simulate click functions

        /// <summary>
        /// Simulates left mouse button click
        /// </summary>
        /// <param name="x">x screen coordinate</param>
        /// <param name="y">y screen coordinate</param>
        /// <param name="keys">keys pressed, 0 - None, 1 - Control pressed, 
        /// 2 - Shift pressed, 3 - Both Control and Shift are pressed</param>
        public void SimulateClickAt(int x, int y, int keys = 0)
        {
            Helper.SimulateClickAt(1, x, y, keys, IntPtr.Zero);
        }

        /// <summary>
        /// Simulates right mouse button click
        /// </summary>
        /// <param name="x">x screen coordinate</param>
        /// <param name="y">y screen coordinate</param>
        /// <param name="keys">keys pressed, 0 - None, 1 - Control pressed, 
        /// 2 - Shift pressed, 3 - Both Control and Shift are pressed</param>
        public void SimulateRightClickAt(int x, int y, int keys = 0)
        {
            Helper.SimulateClickAt(2, x, y, keys, IntPtr.Zero);
        }

        /// <summary>
        /// Simulates middle mouse button click
        /// </summary>
        /// <param name="x">x screen coordinate</param>
        /// <param name="y">y screen coordinate</param>
        /// <param name="keys">keys pressed, 0 - None, 1 - Control pressed, 
        /// 2 - Shift pressed, 3 - Both Control and Shift are pressed</param>
        public void SimulateMiddleClickAt(int x, int y, int keys = 0)
        {
            Helper.SimulateClickAt(3, x, y, keys, IntPtr.Zero);
        }

        /// <summary>
        /// Simulates left mouse button double click
        /// </summary>
        /// <param name="x">x screen coordinate</param>
        /// <param name="y">y screen coordinate</param>
        /// <param name="keys">keys pressed, 0 - None, 1 - Control pressed, 
        /// 2 - Shift pressed, 3 - Both Control and Shift are pressed</param>
        public void SimulateDoubleClickAt(int x, int y, int keys = 0)
        {
            Helper.SimulateClickAt(4, x, y, keys, IntPtr.Zero);
        }
		
		// x, y - screen coordinates 
        /// <summary>
        /// Moves mouse pointer at the specified screen coordinates
        /// </summary>
        /// <param name="x">x - screen coordinate</param>
        /// <param name="y">y - screen coordinate</param>
        /// <param name="keys">keys pressed, 0 - None, 1 - Control pressed, 
        /// 2 - Shift pressed, 3 - Both Control and Shift pressed</param>
        public void MoveMouse(int x, int y, int keys = 0)
        {
            SendInputClass.MoveMousePointer(x, y, keys);
        }
		
		/// <summary>
        /// Moves mouse pointer relatively to the current mouse position
        /// </summary>
        /// <param name="dx">dx - horizontal offset</param>
        /// <param name="dy">dy - vertical offset</param>
        /// <param name="keys">keys pressed, 0 - None, 1 - Control pressed, 
        /// 2 - Shift pressed, 3 - Both Control and Shift pressed</param>
        public void MoveMouseOffset(int dx, int dy, int keys = 0)
        {
			int x = System.Windows.Forms.Cursor.Position.X;
			int y = System.Windows.Forms.Cursor.Position.Y;
            SendInputClass.MoveMousePointer(x + dx, y + dy, keys);
        }
		
		/// <summary>
        /// Scrolls mouse wheel up with the specified number of wheel ticks
        /// </summary>
        /// <param name="wheelTicks">number of wheel ticks</param>
        public void MouseScrollUp(uint wheelTicks)
        {
            SendInputClass.MouseScroll(wheelTicks);
        }
		
		/// <summary>
        /// Scrolls mouse wheel down with the specified number of wheel ticks
        /// </summary>
        /// <param name="wheelTicks">number of wheel ticks</param>
        public void MouseScrollDown(uint wheelTicks)
        {
            SendInputClass.MouseScroll((uint)((-1) * wheelTicks));
        }

        /// <summary>
        /// Gets the currently Engine instance. Should not be used by user.
        /// It is for internal purposes.
        /// </summary>
        /// <returns></returns>
        internal static Engine GetInstance()
        {
            return Engine.instance;
        }

        /// <summary>
        /// Hangs execution untill a certain property of an element reaches the specified value.
		/// For example, if you want to wait until a window closes: engine.WaitUntil(window, "IsAlive", "==", false);
        /// </summary>
        /// <param name="element">Element</param>
        /// <param name="property">Property name. Can be: IsAlive, Text and Value</param>
        /// <param name="comparison">Comparison. Can be ==, !=, &lt; (less than), &gt; (greater than), &lt;= (less or equal), &gt;= (greater or equal)</param>
        /// <param name="value">Value. Depends on the property name. Can be boolean (for IsAlive), a text (for Text) and a real number (for Value)</param>
        /// <param name="timeOut">time out for waiting. Default 5 minutes.</param>
        /// <returns>true if timed out, false otherwise</returns>
        public bool WaitUntil(ElementBase element, string property, 
            string comparison, object value, int timeOut = 300000)
        {
            bool timedOut = false;
            int timeoutTemp = 0;

            if (property == "IsAlive")
            {
                #region IsAlive property
                bool valueBool = false;

                try
                {
                    valueBool = Convert.ToBoolean(value);
                }
                catch 
                {
                    Engine.TraceInLogFile("WaitUntil - value should be a boolean");
                    throw new Exception("WaitUntil - value should be a boolean");
                }

                if (comparison == "==")
                {
                    do
                    {
                        Thread.Sleep(100);
                        timeoutTemp += 100;

                        if (timeoutTemp >= timeOut)
                        {
                            Engine.TraceInLogFile("WaitUntil - timeout reached");
                            timedOut = true;

                            break;
                        }
                    }
                    while (element.IsAlive != valueBool);
                }
                else if (comparison == "!=")
                {
                    do
                    {
                        Thread.Sleep(100);
                        timeoutTemp += 100;

                        if (timeoutTemp >= timeOut)
                        {
                            Engine.TraceInLogFile("WaitUntil - timeout reached");
                            timedOut = true;

                            break;
                        }
                    }
                    while (element.IsAlive == valueBool);
                }
                #endregion
            }
            else if (property == "Text")
            {
                #region Text property
                string valueText = value as string;

                if (valueText == null)
                {
                    Engine.TraceInLogFile("WaitUntil - value must be a string");
                    throw new Exception("WaitUntil - value must be a string");
                }

                if (comparison == "==")
                {
                    do
                    {
                        Thread.Sleep(100);
                        timeoutTemp += 100;

                        if (timeoutTemp >= timeOut)
                        {
                            Engine.TraceInLogFile("WaitUntil - timeout reached");
                            timedOut = true;

                            break;
                        }
                    }
                    while (element.GetText() != valueText);
                }
                else if (comparison == "!=")
                {
                    do
                    {
                        Thread.Sleep(100);
                        timeoutTemp += 100;

                        if (timeoutTemp >= timeOut)
                        {
                            Engine.TraceInLogFile("WaitUntil - timeout reached");
                            timedOut = true;

                            break;
                        }
                    }
                    while (element.GetText() == valueText);
                }
                #endregion
            }
            else if (property == "Value")
            {
                #region Value property
                IUIAutomationRangeValuePattern rangeValuePattern =
                    Helper.GetRangeValuePattern(element.uiElement);

                if (rangeValuePattern == null)
                {
                    Engine.TraceInLogFile("WaitUntil - Element does not support RangeValuePattern");
                    throw new Exception("WaitUntil - Element does not support RangeValuePattern");
                }

                double valueDouble = 0.0;

                try
                {
                    valueDouble = Convert.ToDouble(value);
                }
                catch
                {
                    Engine.TraceInLogFile("WaitUntil - value should be a number");
                    throw new Exception("WaitUntil - value should be a number");
                }

                if (comparison == "==")
                {
                    while (rangeValuePattern.CurrentValue != valueDouble)
                    {
                        Thread.Sleep(100);
                        timeoutTemp += 100;

                        if (timeoutTemp >= timeOut)
                        {
                            Engine.TraceInLogFile("WaitUntil - timeout reached");
                            timedOut = true;

                            break;
                        }
                    }
                }
                else if (comparison == "!=")
                {
                    while (rangeValuePattern.CurrentValue == valueDouble)
                    {
                        Thread.Sleep(100);
                        timeoutTemp += 100;

                        if (timeoutTemp >= timeOut)
                        {
                            Engine.TraceInLogFile("WaitUntil - timeout reached");
                            timedOut = true;

                            break;
                        }
                    }
                }
                else if (comparison == "<")
                {
                    while (rangeValuePattern.CurrentValue >= valueDouble)
                    {
                        Thread.Sleep(100);
                        timeoutTemp += 100;

                        if (timeoutTemp >= timeOut)
                        {
                            Engine.TraceInLogFile("WaitUntil - timeout reached");
                            timedOut = true;

                            break;
                        }
                    }
                }
                else if (comparison == "<=")
                {
                    while (rangeValuePattern.CurrentValue > valueDouble)
                    {
                        Thread.Sleep(100);
                        timeoutTemp += 100;

                        if (timeoutTemp >= timeOut)
                        {
                            Engine.TraceInLogFile("WaitUntil - timeout reached");
                            timedOut = true;

                            break;
                        }
                    }
                }
                else if (comparison == ">")
                {
                    while (rangeValuePattern.CurrentValue <= valueDouble)
                    {
                        Thread.Sleep(100);
                        timeoutTemp += 100;

                        if (timeoutTemp >= timeOut)
                        {
                            Engine.TraceInLogFile("WaitUntil - timeout reached");
                            timedOut = true;

                            break;
                        }
                    }
                }
                else if (comparison == ">=")
                {
                    while (rangeValuePattern.CurrentValue < valueDouble)
                    {
                        Thread.Sleep(100);
                        timeoutTemp += 100;

                        if (timeoutTemp >= timeOut)
                        {
                            Engine.TraceInLogFile("WaitUntil - timeout reached");
                            timedOut = true;

                            break;
                        }
                    }
                }
                else
                {
                    Engine.TraceInLogFile("WaitUntil - comparison sign not recognized");
                    throw new Exception("WaitUntil - comparison sign not recognized");
                }
                #endregion
            }
            else
            {
                Engine.TraceInLogFile("WaitUntil - property not supported");
                throw new Exception("WaitUntil - property not supported");
            }

            return timedOut;
        }

        /// <summary>
        /// Starts a new process
        /// </summary>
		/// <param name="processName">name of the executable file</param>
        /// <returns>Id of the new started process</returns>
		public int StartProcess(string processName)
		{
			try
			{
				Process proc = Process.Start(processName);
				return proc.Id;
			}
			catch (Exception ex)
			{
				Engine.TraceInLogFile("Cannot start process: " + ex.Message);
				throw new Exception("Cannot start process: " + ex.Message);
			}
		}
		
		/// <summary>
        /// Starts a new process and waits for input idle
        /// </summary>
		/// <param name="processName">name of the executable file</param>
        /// <returns>Id of the new started process</returns>
		public int StartProcessAndWaitForInputIdle(string processName)
		{
			try
			{
				Process proc = Process.Start(processName);
				proc.WaitForInputIdle();
				
				for (int i = 0; i < 10; i++)
				{
					Thread.Sleep(100);
				
					try
					{
						if (proc.MainWindowHandle != IntPtr.Zero)
						{
							break;
						}
					}
					catch
					{
						break;
					}
				}
				
				return proc.Id;
			}
			catch (Exception ex)
			{
				Engine.TraceInLogFile("Cannot start process: " + ex.Message);
				throw new Exception("Cannot start process: " + ex.Message);
			}
		}
		
		/// <summary>
        /// Starts a new process with arguments
        /// </summary>
		/// <param name="processName">name of the executable file</param>
		/// <param name="arguments">command line arguments</param>
        /// <returns>Id of the new started process</returns>
		public int StartProcess(string processName, string arguments)
		{
			try
			{
				Process proc = Process.Start(processName, arguments);
				return proc.Id;
			}
			catch (Exception ex)
			{
				Engine.TraceInLogFile("Cannot start process: " + ex.Message);
				throw new Exception("Cannot start process: " + ex.Message);
			}
		}
		
		/// <summary>
        /// Starts a new process with arguments and waits for input idle
        /// </summary>
		/// <param name="processName">name of the executable file</param>
		/// <param name="arguments">command line arguments</param>
        /// <returns>Id of the new started process</returns>
		public int StartProcessAndWaitForInputIdle(string processName, string arguments)
		{
			try
			{
				Process proc = Process.Start(processName, arguments);
				proc.WaitForInputIdle();
				return proc.Id;
			}
			catch (Exception ex)
			{
				Engine.TraceInLogFile("Cannot start process: " + ex.Message);
				throw new Exception("Cannot start process: " + ex.Message);
			}
		}
		
		/// <summary>
        /// Blocks the calling thread for a specified period of time
        /// </summary>
		/// <param name="milliseconds">number of milliseconds</param>
        public void Sleep(int milliseconds)
		{
			try
			{
				Thread.Sleep(milliseconds);
			}
			catch (Exception ex)
			{
				Engine.TraceInLogFile("Sleep failed: " + ex.Message);
				throw new Exception("Sleep failed: " + ex.Message);
			}
		}
		
		/// <summary>
        /// Captures the screen and saves the image in an System.Drawing.Bitmap object. 
		/// You can capture a part of the screen using a crop rectangle.
        /// </summary>
		/// <param name="cropRect">Coordinates of the rectangle to crop. Don't specify it if you want to capture the entire screen.</param>
		/// <returns>System.Drawing.Bitmap that contains the screen or the captured portion of the screen</returns>
        public Bitmap CaptureScreenToBitmap(UIDA_Rect cropRect = null)
        {
            return Helper.CaptureScreen(cropRect);
		}
		
		/// <summary>
        /// Captures the screen and saves the image in the specified file. Supported formats: BMP, JPG/JPEG and PNG. 
		/// You can capture a part of the screen using a crop rectangle.
        /// </summary>
        /// <param name="fileName">Image file path. If only the file name is provided, without a full path, then the image will be saved in the current directory of the application which loaded this library.</param>
		/// <param name="cropRect">Coordinates of the rectangle to crop. Don't specify it if you want to capture the entire screen.</param>
        public void CaptureScreenToFile(string fileName, UIDA_Rect cropRect = null)
        {
            Bitmap bitmap = Helper.CaptureScreen(cropRect);
            
            if (bitmap != null)
            {
                ImageFormat format = ImageFormat.Bmp;
                string extension = System.IO.Path.GetExtension(fileName).ToLower();
                if (extension == "jpg" || extension == "jpeg")
                {
                    format = ImageFormat.Jpeg;
                }
                else if (extension == "png")
                {
                    format = ImageFormat.Png;
                }
                
                bitmap.Save(fileName, format);
                bitmap.Dispose();
            }
        }
		
		/// <summary>
        /// Shows the desktop, minimizes all windows. If you call this function again it will not restore your desktop.
        /// </summary>
		public void ShowDesktop()
		{
			IUIAutomationElement desktopPane = uiAutomation.GetRootElement();
			TreeScope scope = TreeScope.TreeScope_Children;
			IUIAutomationCondition trueCondition = Engine.uiAutomation.CreateTrueCondition();
			
			IUIAutomationElementArray collection = desktopPane.FindAll(scope, trueCondition);
			
			for (int i = 0; i < collection.Length; i++)
			{
				IUIAutomationElement el = collection.GetElement(i);
				UIDA_Window window = new UIDA_Window(el);
				window.Minimize();
			}
			
			Thread.Sleep(100);
		}
		
		/*public void CancelSearch()
		{
			Engine.IsCancelled = true;
		}*/
		
		/// <summary>
        /// Opens a new Google Chrome browser window.
        /// </summary>
		/// <returns>A new Google Chrome browser automation object</returns>
		public ChromeBrowser NewChromeBrowser()
		{
			ChromeBrowser browser = new ChromeBrowser();
			browser.StartBrowser();
			return browser;
		}
		
		/// <summary>
        /// Opens a new Microsoft Edge browser window.
        /// </summary>
		/// <returns>A new Microsoft Edge browser automation object</returns>
		public EdgeBrowser NewEdgeBrowser()
		{
			EdgeBrowser browser = new EdgeBrowser();
			browser.StartBrowser();
			return browser;
		}
    }

    internal enum Errors
    {
        None,
        ElementNotFound,
        IndexTooBig,
        NegativeIndex
    }

    internal enum MouseButtons
    {
        LeftButton = 1, RightButton, Wheel
    }
}
