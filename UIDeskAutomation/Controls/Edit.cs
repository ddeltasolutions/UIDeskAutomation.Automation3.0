using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents an EditBox UI element
    /// </summary>
    public class UIDA_Edit : ElementBase
    {
        internal UIDA_Edit()
        {

        }

        /// <summary>
        /// Creates a UIDA_Edit using an IUIAutomationElement
        /// </summary>
        /// <param name="el">UI Automation Element</param>
        public UIDA_Edit(IUIAutomationElement el)
        {
            base.uiElement = el;
        }

        /// <summary>
        /// Sets text to this EditBox element
        /// </summary>
        /// <param name="text">text to set</param>
        public void SetText(string text)
        {
            if (this.IsAlive == false)
            {
                Engine.TraceInLogFile("This UI element is not available anymore.");
                throw new Exception("This UI element is not available anymore.");
            }

            object valuePatternObj = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId);
            IUIAutomationValuePattern valuePattern = valuePatternObj as IUIAutomationValuePattern;
            
            if (valuePattern != null)
            {
                if (valuePattern.CurrentIsReadOnly != 0)
                {
                    Engine.TraceInLogFile("Edit control is read-only.");
                    throw new Exception("Edit control is read-only");
                }
                else
                {
                    valuePattern.SetValue(text);
                }
                return; // text successfully set
            }

            // try with native Win32 function SetWindowText
            IntPtr hwnd = IntPtr.Zero;
            try
            {
                hwnd = base.uiElement.CurrentNativeWindowHandle;
            }
            catch { }

            if (hwnd != IntPtr.Zero)
            {
                IntPtr textPtr = IntPtr.Zero;
                try
                {
                    textPtr = Marshal.StringToBSTR(text);
                }
                catch { }

                try
                {
                    if (textPtr != IntPtr.Zero)
                    {
                        if (UnsafeNativeFunctions.SendMessage(hwnd,
                            WindowMessages.WM_SETTEXT, IntPtr.Zero, textPtr) ==
                            Win32Constants.TRUE)
                        {
                            return; // text successfully set
                        }
                    }
                }
                catch { }
                finally
                {
                    Marshal.FreeBSTR(textPtr);
                }
            }

            //simulate send chars
            foreach (char ch in text)
            {
                // send WM_CHAR for each character
                UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_CHAR,
                    new IntPtr(ch), new IntPtr(1));
            }
        }

        /// <summary>
        /// Gets the text of this EditBox element
        /// </summary>
        /// <returns>the text of this EditBox element</returns>
        public new string GetText()
        {
            if (this.IsAlive == false)
            {
                Engine.TraceInLogFile("This UI element is not available to the user anymore.");
                throw new Exception("This UI element is not available to the user anymore.");
            }

            object valuePatternObj = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId);
            IUIAutomationValuePattern valuePattern = valuePatternObj as IUIAutomationValuePattern;

            if (valuePattern != null)
            {
                return valuePattern.CurrentValue;
            }

            object textPatternObject = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_TextPatternId);
            IUIAutomationTextPattern textPattern = textPatternObject as IUIAutomationTextPattern;

            if (textPattern != null)
            {
                return textPattern.DocumentRange.GetText(-1);
            }

            // try with native Win32 function SetWindowText
            IntPtr hwnd = IntPtr.Zero;
            try
            {
                hwnd = base.uiElement.CurrentNativeWindowHandle;
            }
            catch { }

            if (hwnd != IntPtr.Zero)
            {
                //StringBuilder text = new StringBuilder(256);
                //UnsafeNativeFunctions.GetWindowText(hwnd, text, 256);
                IntPtr textLengthPtr = UnsafeNativeFunctions.SendMessage(hwnd,
                    WindowMessages.WM_GETTEXTLENGTH, IntPtr.Zero, IntPtr.Zero);

                //textLengthPtr += 1;
                if (textLengthPtr.ToInt32() > 0)
                {
                    int textLength = textLengthPtr.ToInt32() + 1;
                    StringBuilder text = new StringBuilder(textLength);

                    UnsafeNativeFunctions.SendMessage(hwnd,
                        WindowMessages.WM_GETTEXT, textLength, text);

                    return text.ToString();
                }
            }

            return string.Empty;
        }

        /// <summary>
        /// Gets or sets the control's text
        /// </summary>
        public string Text
        {
            get
            {
                return this.GetText();
            }
            set
            {
                this.SetText(value);
            }
        }

        /// <summary>
        /// Clears the text of an edit control.
        /// </summary>
        public void ClearText()
        {
            this.SetText("");
        }

        /// <summary>
        /// Selects a text in an edit control.
        /// </summary>
        /// <param name="text">text to select</param>
        /// <param name="backwards">true if search backwards in edit text, default false</param>
        /// <param name="ignoreCase">true if case is ignored in search, default true</param>
        public void SelectText(string text, bool backwards = false, bool ignoreCase = true)
        {
            object textPatternObject = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_TextPatternId);
            IUIAutomationTextPattern textPattern = textPatternObject as IUIAutomationTextPattern;
            
            if (textPattern == null)
            {
                return;
            }

            if (textPattern.SupportedTextSelection == SupportedTextSelection.SupportedTextSelection_None)
            {
                Engine.TraceInLogFile("SelectText method: selection not supported");
                throw new Exception("SelectText method: selection not supported");
            }

            IUIAutomationTextRange document = null;
            try
            {
                document = textPattern.DocumentRange;
            }
            catch { }

            if (document == null)
            {
                return;
            }

            IUIAutomationTextRange textRange = document.FindText(text, (backwards == true ? 1 : 0), (ignoreCase == true ? 1 : 0));
            if (textRange == null)
            {
                Engine.TraceInLogFile("SelectText method: Cannot find text");
                return;
            }

            try
            {
                textRange.Select();
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("SelectText method: " + ex.Message);
                throw new Exception("SelectText method: " + ex.Message);
            }
        }
		
		/// <summary>
        /// Selects all text in an edit control.
        /// </summary>
        public void SelectAll()
        {
            object textPatternObject = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_TextPatternId);
            IUIAutomationTextPattern textPattern = textPatternObject as IUIAutomationTextPattern;
            
            if (textPattern == null)
            {
                return;
            }

            if (textPattern.SupportedTextSelection == SupportedTextSelection.SupportedTextSelection_None)
            {
                Engine.TraceInLogFile("SelectAll method: selection not supported");
                throw new Exception("SelectAll method: selection not supported");
            }

            IUIAutomationTextRange document = null;
            try
            {
                document = textPattern.DocumentRange;
            }
            catch { }

            if (document == null)
            {
                return;
            }
            
            try
            {
                document.Select();
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("SelectAll method: " + ex.Message);
                throw new Exception("SelectAll method: " + ex.Message);
            }
		}

        /// <summary>
        /// Clears any selected text in an edit control.
        /// </summary>
        public void ClearSelection()
        {
            object textPatternObject = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_TextPatternId);
            IUIAutomationTextPattern textPattern = textPatternObject as IUIAutomationTextPattern;
            
            if (textPattern == null)
            {
                return;
            }

            IUIAutomationTextRangeArray selections = null;
            try
            {
                selections = textPattern.GetSelection();
            }
            catch { }

            if (selections == null)
            {
                return;
            }

            for (int i = 0; i < selections.Length; i++)
            {
                IUIAutomationTextRange selection = selections.GetElement(i);
                try
                {
                    selection.RemoveFromSelection();
                }
                catch (Exception ex)
                {
                    Engine.TraceInLogFile("ClearSelection method: operation not supported");
                    throw new Exception("ClearSelection method: operation not supported");
                }
            }
        }
    }
}
