using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents an Edit UI element
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
        /// Sets text to this Edit element
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
        /// Gets the text of this Edit element
        /// </summary>
        /// <returns>the text of this Edit element</returns>
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
				// try Ctrl+A
				this.KeyDown(VirtualKeys.Control);
				this.KeyPress(VirtualKeys.A);
				this.KeyUp(VirtualKeys.Control);
				//this.SendKeys("^a");
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
		
		/// <summary>
        /// Gets the selected text from the edit control. Returns null if cannot get the selected text.
        /// </summary>
		/// <returns>the selected text of this Edit element</returns>
        public string GetSelectedText()
		{
			IUIAutomationTextPattern textPattern = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_TextPatternId) as IUIAutomationTextPattern;
            if (textPattern == null)
			{
				return null;
			}
			
			IUIAutomationTextRangeArray selections = null;
            try
            {
                selections = textPattern.GetSelection();
            }
            catch { }
			if (selections == null || selections.Length == 0)
            {
                return null;
            }
			
			try
			{
				IUIAutomationTextRange selection = selections.GetElement(0);
				return selection.GetText(-1);
			}
			catch {}
			return null;
		}
		
		private UIA_AutomationEventHandler UIA_TextChangedEventHandler = null;
		private UIA_AutomationEventHandler UIA_TextSelectionChangedEventHandler = null;
		private UIA_AutomationPropertyChangedEventHandler UIA_TextChangedEventHandlerWin32 = null;
		
		/// <summary>
        /// Delegate for Text Changed event
        /// </summary>
		/// <param name="sender">The edit control that sent the event</param>
		/// <param name="text">the text of the edit control</param>
		public delegate void TextChanged(UIDA_Edit sender, string text);
		internal TextChanged TextChangedHandler = null;
		
		/// <summary>
        /// Delegate for Text Selection Changed event
        /// </summary>
		/// <param name="sender">The edit control that sent the event</param>
		/// <param name="selectedText">the selected text</param>
		public delegate void TextSelectionChanged(UIDA_Edit sender, string selectedText);
		internal TextSelectionChanged TextSelectionChangedHandler = null;
		
		/// <summary>
        /// Attaches/detaches a handler to text changed event
        /// </summary>
		public event TextChanged TextChangedEvent
		{
			add
			{
				try
				{
					if (this.TextChangedHandler == null)
					{
						string cfid = base.uiElement.CurrentFrameworkId;
						if (cfid == "Win32" || cfid == "WinForm")
						{
							this.UIA_TextChangedEventHandlerWin32 = new UIA_AutomationPropertyChangedEventHandler(this);
						
							Engine.uiAutomation.AddPropertyChangedEventHandler(base.uiElement, TreeScope.TreeScope_Element, 
								null, this.UIA_TextChangedEventHandlerWin32, new int[] { UIA_PropertyIds.UIA_ValueValuePropertyId });
						}
						else
						{
							this.UIA_TextChangedEventHandler = new UIA_AutomationEventHandler(this);
						
							Engine.uiAutomation.AddAutomationEventHandler(UIA_EventIds.UIA_Text_TextChangedEventId, 
								base.uiElement, TreeScope.TreeScope_Element, null, this.UIA_TextChangedEventHandler);
						}
					}
					
					this.TextChangedHandler += value;
				}
				catch {}
			}
			remove
			{
				try
				{
					this.TextChangedHandler -= value;
				
					if (this.TextChangedHandler == null)
					{
						string cfid = base.uiElement.CurrentFrameworkId;
						if (cfid == "Win32" || cfid == "WinForm")
						{
							RemoveEventHandlerWin32();
						}
						else
						{
							RemoveEventHandler();
						}
					}
				}
				catch {}
			}
		}
		
		/// <summary>
        /// Attaches/detaches a handler to text selection changed event
        /// </summary>
		public event TextSelectionChanged TextSelectionChangedEvent
		{
			add
			{
				try
				{
					if (this.TextSelectionChangedHandler == null)
					{
						this.UIA_TextSelectionChangedEventHandler = new UIA_AutomationEventHandler(this);
					
						Engine.uiAutomation.AddAutomationEventHandler(UIA_EventIds.UIA_Text_TextSelectionChangedEventId, 
							base.uiElement, TreeScope.TreeScope_Element, null, this.UIA_TextSelectionChangedEventHandler);
					}
					
					this.TextSelectionChangedHandler += value;
				}
				catch {}
			}
			remove
			{
				try
				{
					this.TextSelectionChangedHandler -= value;
				
					if (this.TextSelectionChangedHandler == null)
					{
						if (this.UIA_TextSelectionChangedEventHandler == null)
						{
							return;
						}
						
						System.Threading.Tasks.Task.Run(() => 
						{
							try
							{
								Engine.uiAutomation.RemoveAutomationEventHandler(UIA_EventIds.UIA_Text_TextSelectionChangedEventId, 
									base.uiElement, this.UIA_TextSelectionChangedEventHandler);
								UIA_TextSelectionChangedEventHandler = null;
							}
							catch { }
						}).Wait(5000);
					}
				}
				catch {}
			}
		}
		
		private void RemoveEventHandlerWin32()
		{
			if (this.UIA_TextChangedEventHandlerWin32 == null)
			{
				return;
			}
			
			System.Threading.Tasks.Task.Run(() => 
			{
				try
				{
					Engine.uiAutomation.RemovePropertyChangedEventHandler(base.uiElement, 
						this.UIA_TextChangedEventHandlerWin32);
					UIA_TextChangedEventHandlerWin32 = null;
				}
				catch { }
			}).Wait(5000);
		}
		
		private void RemoveEventHandler()
		{
			if (this.UIA_TextChangedEventHandler == null)
			{
				return;
			}
			
			System.Threading.Tasks.Task.Run(() => 
			{
				try
				{
					Engine.uiAutomation.RemoveAutomationEventHandler(UIA_EventIds.UIA_Text_TextChangedEventId, 
						base.uiElement, this.UIA_TextChangedEventHandler);
					UIA_TextChangedEventHandler = null;
				}
				catch { }
			}).Wait(5000);
		}
    }
}
