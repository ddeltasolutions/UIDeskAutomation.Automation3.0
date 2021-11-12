using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Drawing;
using UIAutomationClient;
using System.Windows;

namespace UIDeskAutomationLib
{
    internal class Helper
    {
        public static string WildcardToRegex(string pattern)
        {
            return "^" + Regex.Escape(pattern)
                              .Replace(@"\*", ".*")
                              .Replace(@"\?", ".")
                       + "$";
        }

        public static IntPtr[] GetWindowsWithClass(IntPtr parentHandle, string className)
        {
            List<IntPtr> windowsList = new List<IntPtr>();
            IntPtr currentWindowHandle = IntPtr.Zero;

            do
            {
                currentWindowHandle = UnsafeNativeFunctions.FindWindowEx(parentHandle, currentWindowHandle,
                    className, null);

                if (currentWindowHandle != IntPtr.Zero)
                {
                    windowsList.Add(currentWindowHandle);
                }
            }
            while (currentWindowHandle != IntPtr.Zero);

            return windowsList.ToArray();
        }

        public static IntPtr[] GetWindows(IntPtr parentHandle, string className, 
            string windowText, bool caseSensitive)
        {
            List<IntPtr> windows = new List<IntPtr>();
            IntPtr[] windowsWithClass = Helper.GetWindowsWithClass(parentHandle, className);

            if (windowText == null)
            {
                return windowsWithClass;
            }

            if (caseSensitive == false)
            {
                windowText = windowText.ToLower();
            }

            string windowTextRegEx = Helper.WildcardToRegex(windowText);
            Regex regex = new Regex(windowTextRegEx);

            foreach (IntPtr hwnd in windowsWithClass)
            {
                StringBuilder windowTextBuilder = new StringBuilder(256);
                UnsafeNativeFunctions.GetWindowText(hwnd, windowTextBuilder, 256);

                string currentWindowText = windowTextBuilder.ToString();

                if (caseSensitive == false)
                {
                    currentWindowText = currentWindowText.ToLower();
                }

                Match match = null;

                try
                {
                    match = regex.Match(currentWindowText);
                }
                catch
                {
                    continue;
                }

                if ((match.Success == true) && (match.Value == currentWindowText))
                {
                    windows.Add(hwnd);
                }
            }

            return windows.ToArray();
        }

        //Gets a toplevel window at the specified index using wildcards in windowText
        public static Errors GetWindowAt(IntPtr parentHandle, string className, 
            string windowText, int index, bool caseSensitive, 
            out IntPtr childHandle)
        {
            if (index < 0)
            {
                childHandle = IntPtr.Zero;
                return Errors.NegativeIndex;
            }

            /*if (index == 0)
            {
                index = 1;
            }*/

            IntPtr[] windows = Helper.GetWindows(/*IntPtr.Zero*/ 
                parentHandle, className, windowText, caseSensitive);

            if ((windows == null) || (windows.Length == 0))
            {
                childHandle = IntPtr.Zero;
                return Errors.ElementNotFound;
            }

            if (index > windows.Length)
            {
                childHandle = IntPtr.Zero;
                return Errors.IndexTooBig;
            }

            if (parentHandle == IntPtr.Zero)
            {
                //sort windows list by creation time of starting process for toplevel windows
                List<WindowWithInfo> sortedWindowsList = new List<WindowWithInfo>();

                foreach (IntPtr hwnd in windows)
                {
                    WindowWithInfo windowWithInfo = new WindowWithInfo(hwnd);
                    sortedWindowsList.Add(windowWithInfo);
                }

                sortedWindowsList = sortedWindowsList.OrderBy(x => x.creationDate).ToList();
				if (index == 0)
				{
					// if index is not specified then get the most recent window
					childHandle = sortedWindowsList[sortedWindowsList.Count - 1].hwnd;
				}
				else
				{
					childHandle = sortedWindowsList[index - 1].hwnd;
				}
            }
            else
            {
                //child windows, controls
				if (index == 0)
				{
					childHandle = windows[windows.Length - 1];
				}
				else
				{
					childHandle = windows[index - 1];
				}
            }

            return Errors.None;
        }

        public static List<IUIAutomationElement> MatchStrings(
            IUIAutomationElementArray collection, string name, 
            bool bSearchByLabel, bool caseSensitive)
        {
            List<IUIAutomationElement> returnList = new List<IUIAutomationElement>();
            
            if (name == null)
            {
                for (int i = 0; i < collection.Length; i++)
                {
                    IUIAutomationElement el = collection.GetElement(i);
                    returnList.Add(el);
                }
                return returnList;
            }

            if (caseSensitive == false)
            {
                name = name.ToLower();
            }

            string regExName = Helper.WildcardToRegex(name);
            Regex regex = new Regex(regExName, RegexOptions.Multiline);

            for (int i = 0; i < collection.Length; i++)
            {
                IUIAutomationElement el = collection.GetElement(i);
                string currentName = null;
                
				try
				{
					// AutomationId ??
					currentName = el.CurrentName;
				}
				catch { }

                if (currentName == null)
                {
                    continue;
                }

                if (caseSensitive == false)
                {
                    currentName = currentName.ToLower();
                }

                Match match = null;

                try
                {
                    match = regex.Match(currentName);
                }
                catch
                {
                    continue;
                }

                if ((match.Success == true) && (currentName.Contains(match.Value)))
                {
                    returnList.Add(el);
                }
            }

            return returnList;
        }

        public static IntPtr GetTopLevelByProcId(int processId, 
            string className, string windowText, bool caseSensitive)
        {
            IntPtr windowHandle = IntPtr.Zero;
            uint currentProcessId = 0;

            Regex regex = null;

            if (windowText != null)
            {
                if (caseSensitive == false)
                {
                    windowText = windowText.ToLower();
                }

                string regExText = Helper.WildcardToRegex(windowText);
                regex = new Regex(regExText);
            }

            do
            {
                windowHandle = UnsafeNativeFunctions.FindWindowEx(IntPtr.Zero, windowHandle, className, null);
                UnsafeNativeFunctions.GetWindowThreadProcessId(windowHandle, out currentProcessId);

                StringBuilder windowTextBuilder = new StringBuilder(256);
                UnsafeNativeFunctions.GetWindowText(windowHandle, windowTextBuilder, 256);

                string currentWindowText = windowTextBuilder.ToString();

                if (caseSensitive == false)
                {
                    currentWindowText = currentWindowText.ToLower();
                }

                if ((int)currentProcessId == processId)
                {
                    if (windowText != null)
                    {
                        Match match = null;

                        try
                        {
                            match = regex.Match(currentWindowText);
                        }
                        catch
                        {
                            continue;
                        }

                        if ((match.Success == true) && (match.Value == currentWindowText))
                        {
                            break;
                        }
                    }
                    else
                    {
                        break;
                    }
                }
            }
            while (windowHandle != IntPtr.Zero);

            return windowHandle;
        }

        #region Text functions
        public static string GetElementName(IUIAutomationElement element)
        {
            if (element == null)
            {
                throw new Exception("Invalid Automation Element");
            }

            int type = -1;
            try
            {
                type = element.CurrentControlType;
            }
            catch { }

            string name = null;

            try
            {
                name = element.CurrentName;
            }
            catch { }
            
            //for ComboBox, Document and Edit ignore the AutomationElement.Name property
            if (type == UIA_ControlTypeIds.UIA_ComboBoxControlTypeId || type == UIA_ControlTypeIds.UIA_DocumentControlTypeId || type == UIA_ControlTypeIds.UIA_EditControlTypeId)
            {
                return null;
            }

            return name;
        }

        public static string GetElementText(IUIAutomationElement element)
        {
            if (element == null)
            {
                throw new Exception("Invalid Automation Element");
            }

            IUIAutomationValuePattern valuePattern = element.GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId) as IUIAutomationValuePattern;
            if (valuePattern == null)
            {
                // ValuePattern not supported
                string text = null;
                try
                {
                    text = element.CurrentName;
                }
                catch { }

                return text;
            }

            // ValuePattern supported
            string value = null;
            try
            {
                value = valuePattern.CurrentValue;
            }
            catch { }

            return value;
        }
        #endregion

        // xScreen, yScreen - screen coordinates
        /// <summary>
        /// Simulates click
        /// </summary>
        /// <param name="mouseButton">1 - Left click, 2 - Right click, 
        /// 3 - Middle click, 4 - Double click</param>
        /// <param name="xScreen"></param>
        /// <param name="yScreen"></param>
        /// <param name="keys"></param>
        /// <param name="hwnd"></param>
        public static void SimulateClickAt(int mouseButton, int xScreen, int yScreen, 
            int keys, IntPtr hwnd)
        {
            POINT ptScreen = new POINT(xScreen, yScreen);

            if (hwnd == IntPtr.Zero)
            {
                hwnd = UnsafeNativeFunctions.WindowFromPoint(ptScreen);
            }

            if (UnsafeNativeFunctions.ScreenToClient(hwnd, ref ptScreen) == false)
            {
                Engine.TraceInLogFile("Helper::SimulateClickAt - Error.");
                throw new Exception("Helper::SimulateClickAt - Error.");
            }

            // not ptScreen has client coordinates
            int coordinates = (ptScreen.Y << 16) + ptScreen.X;

            uint mouseDownMessage = 0;
            uint mouseUpMessage = 0;
            int mouseButtonConstant = 0;

            if (mouseButton == 1) // left mouse button
            {
                mouseDownMessage = WindowMessages.WM_LBUTTONDOWN;
                mouseUpMessage = WindowMessages.WM_LBUTTONUP;
                mouseButtonConstant = Win32Constants.MK_LBUTTON;
            }
            else if (mouseButton == 2) // right mouse button
            {
                mouseDownMessage = WindowMessages.WM_RBUTTONDOWN;
                mouseUpMessage = WindowMessages.WM_RBUTTONUP;
                mouseButtonConstant = Win32Constants.MK_RBUTTON;
            }
            else if (mouseButton == 3) // middle mouse button
            {
                mouseDownMessage = WindowMessages.WM_MBUTTONDOWN;
                mouseUpMessage = WindowMessages.WM_MBUTTONUP;
                mouseButtonConstant = Win32Constants.MK_MBUTTON;
            }

            if (keys == 0) // no key pressed
            {
                if (mouseButton == 4) // double click
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_LBUTTONDOWN,
                        new IntPtr(Win32Constants.MK_LBUTTON), new IntPtr(coordinates));
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_LBUTTONUP,
                        new IntPtr(Win32Constants.MK_LBUTTON), new IntPtr(coordinates));
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_LBUTTONDBLCLK,
                        new IntPtr(Win32Constants.MK_LBUTTON), new IntPtr(coordinates));
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_LBUTTONUP,
                        new IntPtr(Win32Constants.MK_LBUTTON), new IntPtr(coordinates));
                }
                else
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, mouseDownMessage,
                        new IntPtr(mouseButtonConstant), new IntPtr(coordinates));
                    UnsafeNativeFunctions.PostMessage(hwnd, mouseUpMessage,
                        new IntPtr(mouseButtonConstant), new IntPtr(coordinates));
                }
            }
            else if (keys == 1) // control is pressed
            {
                if (mouseButton == 4) // double click
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_LBUTTONDOWN,
                        new IntPtr(Win32Constants.MK_LBUTTON | Win32Constants.MK_CONTROL), 
                        new IntPtr(coordinates));
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_LBUTTONUP,
                        new IntPtr(Win32Constants.MK_LBUTTON | Win32Constants.MK_CONTROL), 
                        new IntPtr(coordinates));
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_LBUTTONDBLCLK,
                        new IntPtr(Win32Constants.MK_LBUTTON | Win32Constants.MK_CONTROL), 
                        new IntPtr(coordinates));
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_LBUTTONUP,
                        new IntPtr(Win32Constants.MK_LBUTTON | Win32Constants.MK_CONTROL), 
                        new IntPtr(coordinates));
                }
                else
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, mouseDownMessage,
                        new IntPtr(mouseButtonConstant | Win32Constants.MK_CONTROL),
                        new IntPtr(coordinates));
                    UnsafeNativeFunctions.PostMessage(hwnd, mouseUpMessage,
                        new IntPtr(mouseButtonConstant | Win32Constants.MK_CONTROL),
                        new IntPtr(coordinates));
                }
            }
            else if (keys == 2) // shift is pressed
            {
                if (mouseButton == 4) // double click
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_LBUTTONDOWN,
                        new IntPtr(Win32Constants.MK_LBUTTON | Win32Constants.MK_SHIFT), 
                        new IntPtr(coordinates));
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_LBUTTONUP,
                        new IntPtr(Win32Constants.MK_LBUTTON | Win32Constants.MK_SHIFT), 
                        new IntPtr(coordinates));
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_LBUTTONDBLCLK,
                        new IntPtr(Win32Constants.MK_LBUTTON | Win32Constants.MK_SHIFT), 
                        new IntPtr(coordinates));
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_LBUTTONUP,
                        new IntPtr(Win32Constants.MK_LBUTTON | Win32Constants.MK_SHIFT), 
                        new IntPtr(coordinates));
                }
                else
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, mouseDownMessage,
                        new IntPtr(mouseButtonConstant | Win32Constants.MK_SHIFT),
                        new IntPtr(coordinates));
                    UnsafeNativeFunctions.PostMessage(hwnd, mouseUpMessage,
                        new IntPtr(mouseButtonConstant | Win32Constants.MK_SHIFT),
                        new IntPtr(coordinates));
                }
            }
            else if (keys == 3) // both control and shift are pressed
            {
                if (mouseButton == 4)
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_LBUTTONDOWN,
                        new IntPtr(Win32Constants.MK_LBUTTON | Win32Constants.MK_CONTROL | Win32Constants.MK_SHIFT), 
                        new IntPtr(coordinates));
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_LBUTTONUP,
                        new IntPtr(Win32Constants.MK_LBUTTON | Win32Constants.MK_CONTROL | Win32Constants.MK_SHIFT), 
                        new IntPtr(coordinates));
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_LBUTTONDBLCLK,
                        new IntPtr(Win32Constants.MK_LBUTTON | Win32Constants.MK_CONTROL | Win32Constants.MK_SHIFT), 
                        new IntPtr(coordinates));
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_LBUTTONUP,
                        new IntPtr(Win32Constants.MK_LBUTTON | Win32Constants.MK_CONTROL | Win32Constants.MK_SHIFT), 
                        new IntPtr(coordinates));
                }
                else
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, mouseDownMessage,
                        new IntPtr(mouseButtonConstant | Win32Constants.MK_CONTROL | Win32Constants.MK_SHIFT),
                        new IntPtr(coordinates));
                    UnsafeNativeFunctions.PostMessage(hwnd, mouseUpMessage,
                        new IntPtr(mouseButtonConstant | Win32Constants.MK_CONTROL | Win32Constants.MK_SHIFT),
                        new IntPtr(coordinates));
                }
            }
            else
            {
                Engine.TraceInLogFile("Helper::SimulateClickAt - invalid argument 'keys'");
                throw new Exception("Helper::SimulateClickAt - invalid argument 'keys'");
            }
        }

        internal static string[] GetKeys(string text)
        {
            bool insideBrackets = false;
            List<string> keys = new List<string>();
            string currentKey = string.Empty;

            for (int i = 0; i < text.Length; i++)
            {
                if ((text[i] == '{') && (insideBrackets == false))
                {
                    insideBrackets = true;
                    currentKey += text[i];

                    continue;
                }

                if (text[i] == '}')
                {
                    if ((i < (text.Length - 1)) && (text[i + 1] == '}'))
                    {
                        currentKey += "}";
                        i++;
                    }
                    insideBrackets = false;

                    currentKey += text[i];
                    keys.Add(currentKey);
                    currentKey = string.Empty;

                    continue;
                }

                if (insideBrackets == false)
                {
                    currentKey = new string(text[i], 1);
                    keys.Add(currentKey);
                    currentKey = string.Empty;
                }
                else
                {
                    currentKey += text[i];
                }
            }

            return keys.ToArray();
        }

        internal static void SimulateSendKeys(string text, IntPtr hwnd)
        {
            string[] keys = Helper.GetKeys(text);

            bool bInsideBrackets = false;
            bool bShiftIsPressed = false;
            bool bCtrlIsPressed = false;
            bool bAltIsPressed = false;

            foreach (string key in keys)
            {
                if ((key == "(") && (bInsideBrackets == false))
                {
                    bInsideBrackets = true;
                    continue;
                }
                if ((key == ")") && (bInsideBrackets == true))
                {
                    // Release Control, Alt or Shift if they are pressed
                    if (bShiftIsPressed == true)
                    {
                        // Shift up
                        UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYUP,
                            new IntPtr((int)VirtualKeys.Shift), IntPtr.Zero);
                        bShiftIsPressed = false;
                    }
                    if (bCtrlIsPressed == true)
                    {
                        // Ctrl up
                        UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYUP,
                            new IntPtr((int)VirtualKeys.Control), IntPtr.Zero);
                        bCtrlIsPressed = false;
                    }
                    if (bAltIsPressed == true)
                    {
                        // Alt up
                        UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYUP,
                            new IntPtr((int)VirtualKeys.Menu), IntPtr.Zero);
                        bAltIsPressed = false;
                    }

                    bInsideBrackets = false;
                    continue;
                }

                // Press key
                if (key == "%")
                {
                    // Alt key down
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.Menu), IntPtr.Zero);
                    bAltIsPressed = true;
                }
                else if (key == "+")
                {
                    // Shift down
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.Shift), IntPtr.Zero);
                    bShiftIsPressed = true;
                }
                else if (key == "^")
                {
                    // Ctrl down
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.Control), IntPtr.Zero);
                    bCtrlIsPressed = true;
                }
                else if ((key == "{BACKSPACE}") || (key == "{BS}") || (key == "{BKSP}"))
                {
                    // Backspace
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.Back), IntPtr.Zero);
                }
                else if (key == "{BREAK}") // Ctrl-Break
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.Cancel), IntPtr.Zero);
                }
                else if ((key == "{DELETE}") || (key == "{DEL}"))
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.Delete), IntPtr.Zero);
                }
                else if (key == "{DOWN}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.Down), IntPtr.Zero);
                }
                else if (key == "{END}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.End), IntPtr.Zero);
                }
                else if (key == "{ENTER}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.Return), IntPtr.Zero);
                }
                else if (key == "{ESC}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.Escape), IntPtr.Zero);
                }
                else if (key == "{HELP}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.Help), IntPtr.Zero);
                }
                else if (key == "{HOME}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.Home), IntPtr.Zero);
                }
                else if ((key == "{INS}") || (key == "{INSERT}"))
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.Insert), IntPtr.Zero);
                }
                else if (key == "{LEFT}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.Left), IntPtr.Zero);
                }
                else if (key == "{PGDN}") // page down
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.Next), IntPtr.Zero);
                }
                else if (key == "{PGUP}") // page up
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.Prior), IntPtr.Zero);
                }
                else if (key == "{PRTSC}") // print screen
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.Snapshot), IntPtr.Zero);
                }
                else if (key == "{RIGHT}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.Right), IntPtr.Zero);
                }
                else if (key == "{TAB}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.Tab), IntPtr.Zero);
                }
                else if (key == "{UP}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.Up), IntPtr.Zero);
                }
                else if (key == "{F1}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.F1), IntPtr.Zero);
                }
                else if (key == "{F2}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.F2), IntPtr.Zero);
                }
                else if (key == "{F3}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.F3), IntPtr.Zero);
                }
                else if (key == "{F4}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.F4), IntPtr.Zero);
                }
                else if (key == "{F5}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.F5), IntPtr.Zero);
                }
                else if (key == "{F6}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.F6), IntPtr.Zero);
                }
                else if (key == "{F7}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.F7), IntPtr.Zero);
                }
                else if (key == "{F8}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.F8), IntPtr.Zero);
                }
                else if (key == "{F9}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.F9), IntPtr.Zero);
                }
                else if (key == "{F10}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.F10), IntPtr.Zero);
                }
                else if (key == "{F11}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.F11), IntPtr.Zero);
                }
                else if (key == "{F12}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.F12), IntPtr.Zero);
                }
                else if (key == "{F13}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.F13), IntPtr.Zero);
                }
                else if (key == "{F14}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.F14), IntPtr.Zero);
                }
                else if (key == "{F15}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.F15), IntPtr.Zero);
                }
                else if (key == "{F16}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.F16), IntPtr.Zero);
                }
                else if (key == "{ADD}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.Add), IntPtr.Zero);
                }
                else if (key == "{SUBTRACT}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.Subtract), IntPtr.Zero);
                }
                else if (key == "{MULTIPLY}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.Multiply), IntPtr.Zero);
                }
                else if (key == "{DIVIDE}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr((int)VirtualKeys.Divide), IntPtr.Zero);
                }
                else if (key.Length == 1) // one character string
                {
                    // Post WM_CHAR message
                    Char ch = key[0];
                    if (ch == '~')
                    {
                        // Enter
                        UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                            new IntPtr((int)VirtualKeys.Return), IntPtr.Zero);
                    }
                    else
                    {
                        UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_CHAR,
                            new IntPtr(ch), IntPtr.Zero);
                    }
                }

                // Release Alt, Shift or Control if outside brackets ()
                bool bIsAltShiftOrCtrl = (key == "%") || (key == "+") || (key == "^");
                if ((bInsideBrackets == false) && (bIsAltShiftOrCtrl == false))
                {
                    if (bShiftIsPressed == true)
                    {
                        // Shift up
                        UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYUP,
                            new IntPtr((int)VirtualKeys.Shift), IntPtr.Zero);
                        bShiftIsPressed = false;
                    }
                    if (bCtrlIsPressed == true)
                    {
                        // Ctrl up
                        UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYUP,
                            new IntPtr((int)VirtualKeys.Control), IntPtr.Zero);
                        bCtrlIsPressed = false;
                    }
                    if (bAltIsPressed == true)
                    {
                        // Alt up
                        UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYUP,
                            new IntPtr((int)VirtualKeys.Menu), IntPtr.Zero);
                        bAltIsPressed = false;
                    }
                }
            }
        }

        private static void PressKey(IntPtr chVirtKeyCode, IntPtr hwnd)
        {
            UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                chVirtKeyCode, IntPtr.Zero);

            UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYUP,
                chVirtKeyCode, IntPtr.Zero);
        }

        internal static IntPtr GetTopLevelByProcName(string processName, 
            string className, string windowText, int index, bool caseSensitive)
        {
            processName = processName.ToLower();

            if (caseSensitive == false)
            {
                windowText = windowText.ToLower();
            }

            IntPtr windowHandle = IntPtr.Zero;
            uint currentProcessId = 0;

            Regex regex = null;

            if (windowText != null)
            {
                string regExText = Helper.WildcardToRegex(windowText);
                regex = new Regex(regExText);
            }

            do
            {
                windowHandle = UnsafeNativeFunctions.FindWindowEx(IntPtr.Zero, windowHandle, className, null);

                if (windowHandle == IntPtr.Zero)
                {
                    break;
                }

                if (UnsafeNativeFunctions.IsWindowVisible(windowHandle) == false)
                { 
                    // continue to next sibling window if current window is not visible
                    continue;
                }

                UnsafeNativeFunctions.GetWindowThreadProcessId(windowHandle, out currentProcessId);

                Process currentProcess = Process.GetProcessById((int)currentProcessId);
                string currentProcessName = currentProcess.ProcessName.ToLower() + ".exe";

                StringBuilder windowTextBuilder = new StringBuilder(256);
                UnsafeNativeFunctions.GetWindowText(windowHandle, windowTextBuilder, 256);

                string currentWindowText = windowTextBuilder.ToString();

                if (caseSensitive == false)
                {
                    currentWindowText = currentWindowText.ToLower();
                }

                if (currentProcessName == processName)
                {
                    if (windowText != null)
                    {
                        Match match = null;

                        try
                        {
                            match = regex.Match(currentWindowText);
                        }
                        catch 
                        {
                            continue;
                        }

                        if ((match.Success == true) && (match.Value == currentWindowText))
                        {
                            index--;
                        }
                    }
                    else
                    {
                        index--;
                    }

                    if (index == 0)
                    {
                        break;
                    }
                }
            }
            while (windowHandle != IntPtr.Zero);

            if (index == 0)
            {
                return windowHandle;
            }
            else
            {
                // not found
                return IntPtr.Zero;
            }
        }

        //Returns true is elements are identical, false otherwise.
        internal static bool CompareAutomationElements(IUIAutomationElement el1,
            IUIAutomationElement el2)
        {
            Array runtimeId1 = null;
            Array runtimeId2 = null;

            try
            {
                runtimeId1 = el1.GetRuntimeId();
                runtimeId2 = el2.GetRuntimeId();
            }
            catch { }

            if ((runtimeId1 == null) || (runtimeId2 == null))
            {
                return (el1 == el2);
            }

            if (runtimeId1.Length != runtimeId2.Length)
            {
                return false;
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

            return !different;
        }

        internal static IUIAutomationValuePattern GetValuePattern(IUIAutomationElement uiElement)
        {
            object valuePatternObj = uiElement.GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId);
            IUIAutomationValuePattern valuePattern = valuePatternObj as IUIAutomationValuePattern;

            return valuePattern;
        }

        internal static IUIAutomationRangeValuePattern GetRangeValuePattern(IUIAutomationElement uiElement)
        {
            object rangeValuePatternObj = uiElement.GetCurrentPattern(UIA_PatternIds.UIA_RangeValuePatternId);
            IUIAutomationRangeValuePattern rangeValuePattern = rangeValuePatternObj as IUIAutomationRangeValuePattern;

            return rangeValuePattern;
        }
        
        private static Bitmap CreateWindowSnapshot(IntPtr hWnd, IUIAutomationElement element)
        {
            RECT rect;
            if (UnsafeNativeFunctions.GetWindowRect(hWnd, out rect) == false)
            {
                return null;
            }
            
            int width = rect.Right - rect.Left;
            int height = rect.Bottom - rect.Top;
            int left = 0;
            int top = 0;

            IntPtr hdc = UnsafeNativeFunctions.GetWindowDC(hWnd);
            //IntPtr hdc = UnsafeNativeFunctions.GetDC(hWnd);
            if (hdc == IntPtr.Zero)
            {
                return null;
            }
            IntPtr hdcMem = UnsafeNativeFunctions.CreateCompatibleDC(hdc);
            if (hdcMem == IntPtr.Zero)
            {
                return null;
            }
            IntPtr hbmp = UnsafeNativeFunctions.CreateCompatibleBitmap(hdc, width, height);
            if (hbmp == IntPtr.Zero)
            {
                return null;
            }
            if (UnsafeNativeFunctions.SelectObject(hdcMem, hbmp) == IntPtr.Zero)
            {
                return null;
            }
            if (UnsafeNativeFunctions.BitBlt(hdcMem, 0, 0, width, height, hdc, left, top, 
                TernaryRasterOperations.SRCCOPY | TernaryRasterOperations.CAPTUREBLT) == false)
            {
                return null;
            }

            Bitmap bitmap = System.Drawing.Image.FromHbitmap(hbmp);

            UnsafeNativeFunctions.DeleteObject(hbmp);
            UnsafeNativeFunctions.DeleteDC(hdcMem);
            UnsafeNativeFunctions.ReleaseDC(hWnd, hdc);
            return bitmap;
        }
        
        private static Bitmap CreateSnapshot(IntPtr hWnd, int left, int top, int right, int bottom)
        {
            POINT pt;
            pt.X = left;
            pt.Y = top;
            if (hWnd != IntPtr.Zero && UnsafeNativeFunctions.ScreenToClient(hWnd, ref pt) == false)
            {
                return null;
            }
            if (pt.X < 0 || pt.Y < 0)
            {
                Engine.TraceInLogFile("Out of client area");
                return null;
            }
            int width = right - left;
            int height = bottom - top;

            //HDC hdc = GetWindowDC(hWnd);
            IntPtr hdc = UnsafeNativeFunctions.GetDC(hWnd);
            if (hdc == IntPtr.Zero)
            {
                return null;
            }
            IntPtr hdcMem = UnsafeNativeFunctions.CreateCompatibleDC(hdc);
            if (hdcMem == IntPtr.Zero)
            {
                return null;
            }
            IntPtr hbmp = UnsafeNativeFunctions.CreateCompatibleBitmap(hdc, width, height);
            if (hbmp == IntPtr.Zero)
            {
                return null;
            }
            if (UnsafeNativeFunctions.SelectObject(hdcMem, hbmp) == IntPtr.Zero)
            {
                return null;
            }
            if (UnsafeNativeFunctions.BitBlt(hdcMem, 0, 0, width, height, hdc, pt.X, pt.Y, 
                TernaryRasterOperations.SRCCOPY | TernaryRasterOperations.CAPTUREBLT) == false)
            {
                return null;
            }

            Bitmap bitmap = System.Drawing.Image.FromHbitmap(hbmp);

            UnsafeNativeFunctions.DeleteObject(hbmp);
            UnsafeNativeFunctions.DeleteDC(hdcMem);
            UnsafeNativeFunctions.ReleaseDC(hWnd, hdc);
            return bitmap;
        }
        
        private static IUIAutomationElement GetParentWindow(IUIAutomationElement element)
        {
            IUIAutomationTreeWalker tw = Engine.uiAutomation.ControlViewWalker;
            IUIAutomationElement root = Engine.uiAutomation.GetRootElement();
            IUIAutomationElement crtElement = element;
            
            if (crtElement == null || CompareAutomationElements(crtElement, root) == true)
            {
                return null;
            }
            
            do
            {
                crtElement = tw.GetParentElement(crtElement);
                if (crtElement == null || CompareAutomationElements(crtElement, root) == true)
                {
                    return null;
                }
                
                IntPtr hWnd = crtElement.CurrentNativeWindowHandle;
                if (hWnd != IntPtr.Zero)
                {
                    return crtElement;
                }
            }
            while (true);
            
            return null;
        }
        
        internal static Bitmap CaptureElement(IUIAutomationElement element)
        {
            if (element.CurrentControlType == UIA_ControlTypeIds.UIA_MenuControlTypeId || 
                element.CurrentControlType == UIA_ControlTypeIds.UIA_MenuItemControlTypeId)
            {
                return null;
            }
            
            if (element.CurrentControlType == UIA_ControlTypeIds.UIA_SpinnerControlTypeId && element.CurrentFrameworkId == "WinForm")
            {
                element = Engine.uiAutomation.ControlViewWalker.GetParentElement(element);
            }
            
			Bitmap bitmap = null;
			IntPtr hWnd = element.CurrentNativeWindowHandle;
			if (hWnd != IntPtr.Zero)
			{
				bitmap = CreateWindowSnapshot(hWnd, element);
				if (bitmap != null)
				{
					return bitmap;
				}
			}
			IUIAutomationElement crtParent = element;
			if (hWnd == IntPtr.Zero)
			{
				crtParent = GetParentWindow(element);
				if (crtParent == null)
				{
					return null;
				}
				hWnd = crtParent.CurrentNativeWindowHandle;
			}
			
			tagRECT rect = element.CurrentBoundingRectangle;
			
			while (bitmap == null)
			{
				bitmap = CreateSnapshot(hWnd, rect.left, rect.top, rect.right, rect.bottom);
				if (bitmap == null)
				{
					crtParent = GetParentWindow(crtParent);
					if (crtParent == null)
					{
						return null;
					}
					hWnd = crtParent.CurrentNativeWindowHandle;
				}
			}
			
			return bitmap;
        }
		
		internal static Bitmap CaptureVisibleElement(IUIAutomationElement element, UIDA_Rect cropRect = null)
		{
			tagRECT rect = element.CurrentBoundingRectangle;
			
			System.Drawing.Size size = new System.Drawing.Size(rect.right - rect.left, rect.bottom - rect.top);
			if (cropRect != null)
			{
				size = new System.Drawing.Size(cropRect.Width, cropRect.Height);
			}
			
			IntPtr hwnd = element.CurrentNativeWindowHandle;
			Bitmap bitmap = null;
			if (hwnd == IntPtr.Zero)
			{
				bitmap = new Bitmap(size.Width, size.Height);
			}
			else
			{
				Graphics gFromHwnd = Graphics.FromHwnd(hwnd);
				bitmap = new Bitmap(size.Width, size.Height, gFromHwnd);
			}
			
			using (Graphics g = Graphics.FromImage(bitmap))
			{
				if (cropRect != null)
				{
					g.CopyFromScreen(rect.left + cropRect.Left, rect.top + cropRect.Top, 0, 0, size);
				}
				else
				{
					g.CopyFromScreen(rect.left, rect.top, 0, 0, size);
				}
			}
			
			return bitmap;
		}
		
		internal static Bitmap CropImage(Bitmap image, Rectangle cropRect)
		{
			Bitmap result = new Bitmap(cropRect.Width, cropRect.Height);

			using(Graphics g = Graphics.FromImage(result))
			{
				g.DrawImage(image, new Rectangle(0, 0, result.Width, result.Height), 
						cropRect, GraphicsUnit.Pixel);
			}
			
			return result;
		}
		
		internal static Bitmap CaptureScreen(UIDA_Rect cropRect = null)
		{
			System.Drawing.Size size = new System.Drawing.Size(
				System.Windows.Forms.SystemInformation.VirtualScreen.Width, 
				System.Windows.Forms.SystemInformation.VirtualScreen.Height);
			if (cropRect != null)
			{
				size = new System.Drawing.Size(cropRect.Width, cropRect.Height);
			}
			
			Bitmap bitmap = new Bitmap(size.Width, size.Height);
			using (Graphics g = Graphics.FromImage(bitmap))
			{
				if (cropRect != null)
				{
					g.CopyFromScreen(cropRect.Left, cropRect.Top, 0, 0, size);
				}
				else
				{
					g.CopyFromScreen(0, 0, 0, 0, size);
				}
			}
			
			return bitmap;
		}
    }

    internal class WindowWithInfo
    {
        public IntPtr hwnd = IntPtr.Zero;

        public DateTime creationDate = DateTime.MinValue;

        public WindowWithInfo(IntPtr windowHandle)
        {
            this.hwnd = windowHandle;

            uint processId = 0;

            UnsafeNativeFunctions.GetWindowThreadProcessId(this.hwnd, out processId);

            Process process = Process.GetProcessById(Convert.ToInt32(processId));

            this.creationDate = process.StartTime;
        }
    }
}
