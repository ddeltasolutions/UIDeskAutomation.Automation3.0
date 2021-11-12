using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Generic base class for Spinner, Slider, ProgressBar and Scrollbar.
    /// This class is used only internally and cannot be instantiated.
    /// </summary>
    public abstract class GenericSpinner: ElementBase
    {
        internal GenericSpinner(IUIAutomationElement el)
        {
            this.uiElement = el;
        }

        /// <summary>
        /// Gets/Sets the value of the current element spinner/slider/progressbar.
        /// Value of a progress bar cannot be set, only read.
        /// </summary>
        internal double Value
        {
            get
            {
                string localizedControlType = "current element";

                try
                {
                    localizedControlType = this.uiElement.CurrentLocalizedControlType;
                }
                catch { }

                try
                {
                    if (this.uiElement.CurrentControlType == UIA_ControlTypeIds.UIA_SpinnerControlTypeId)
                    {
                        double? val = GetSetWindowsFormsSpinnerValue();

                        if (val.HasValue)
                        {
                            return val.Value;
                        }
                    }

                }
                catch { }

                object rangeValuePatternObj = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_RangeValuePatternId);
                IUIAutomationRangeValuePattern rangeValuePattern = rangeValuePatternObj as IUIAutomationRangeValuePattern;

                if (rangeValuePattern == null)
                {
                    object valuePatternObj = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId);
                    IUIAutomationValuePattern valuePattern = valuePatternObj as IUIAutomationValuePattern;
                    
                    if (valuePattern == null)
                    {
                        Engine.TraceInLogFile("Cannot get value of " + localizedControlType);
                        throw new Exception("Cannot get value of " + localizedControlType);
                    }
                    
                    try
                    {
                        return Convert.ToDouble(valuePattern.CurrentValue);
                    }
                    catch (Exception ex)
                    {
                        Engine.TraceInLogFile("Cannot get " + 
                            localizedControlType + " value: " + ex.Message);
                        throw new Exception("Cannot get " + 
                            localizedControlType + " value: " + ex.Message);
                    }
                }

                try
                {
                    return rangeValuePattern.CurrentValue;
                }
                catch (Exception ex)
                {
                    Engine.TraceInLogFile("Cannot get " + 
                        localizedControlType + " value: " + ex.Message);
                    throw new Exception("Cannot get " + 
                        localizedControlType + " value: " + ex.Message);
                }
            }

            set
            {
                try
                {
                    // Don't let the user to set the value of progress bar control.
                    if (this.uiElement.CurrentControlType == UIA_ControlTypeIds.UIA_ProgressBarControlTypeId)
                    {
                        Engine.TraceInLogFile("Cannot set the value of a progress bar");
                        throw new Exception("Cannot set the value of a progress bar");
                    }
                }
                catch (Exception ex)
                {
                    Engine.TraceInLogFile("Set value error: " + ex.Message);
                    throw new Exception("Set value error: " + ex.Message);
                }

                string localizedControlType = "current element";
                try
                {
                    localizedControlType = this.uiElement.CurrentLocalizedControlType;
                }
                catch { }

                string message = null;
                object rangeValuePatternObj = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_RangeValuePatternId);
                IUIAutomationRangeValuePattern rangeValuePattern = rangeValuePatternObj as IUIAutomationRangeValuePattern;

                if (rangeValuePattern != null)
                {
                    try
                    {
                        rangeValuePattern.SetValue(value);
                        return;
                    }
                    catch (Exception ex)
                    {
                        message = "Cannot set value of " + localizedControlType + ": " + ex.Message;
                    }
                }
                else
                {
                    object valuePatternObj = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId);
                    IUIAutomationValuePattern valuePattern = valuePatternObj as IUIAutomationValuePattern;
                    
                    if (valuePattern != null)
                    {
                        try
                        {
                            valuePattern.SetValue(value.ToString());
                            return;
                        }
                        catch (Exception ex)
                        {
                            message = "Cannot set value of " + localizedControlType + ": " + ex.Message;
                        }
                    }
                }
				
				try
                {
                    if (this.uiElement.CurrentControlType == UIA_ControlTypeIds.UIA_SpinnerControlTypeId)
                    {
                        if (GetSetWindowsFormsSpinnerValue(value, true) != null)
						{
							return;
						}
                    }
                }
                catch { }

				if (message != null)
				{
					Engine.TraceInLogFile(message);
					throw new Exception(message);
				}
				else
				{
					Engine.TraceInLogFile("Cannot set value of " + localizedControlType);
					throw new Exception("Cannot set value of " + localizedControlType);
				}
            }
        }

        // for set = false, 'value' parameter is ignored
        private double? GetSetWindowsFormsSpinnerValue(double value = 0.0, 
            bool set = false)
        {
            IntPtr windowHandle = this.GetWindow();

            if (windowHandle == IntPtr.Zero)
            {
                return null;
            }

            StringBuilder className = new StringBuilder(256);
            UnsafeNativeFunctions.GetClassName(windowHandle, className, 256);

            if (className.ToString().StartsWith("WindowsForms"))
            {
                // Windows Forms spinner
                IntPtr hwndEdit = IntPtr.Zero;

                IntPtr hwndChild = UnsafeNativeFunctions.FindWindowEx(
                    windowHandle, IntPtr.Zero, null, null);

                while (hwndChild != IntPtr.Zero)
                {
                    StringBuilder childClassName = new StringBuilder(256);
                    UnsafeNativeFunctions.GetClassName(
                        hwndChild, childClassName, 256);
                    if (childClassName.ToString().ToLower().Contains("edit"))
                    {
                        hwndEdit = hwndChild;
                        break;
                    }

                    hwndChild = UnsafeNativeFunctions.FindWindowEx(windowHandle,
                        hwndChild, null, null);
                }

                if (hwndEdit != IntPtr.Zero)
                {
                    if (set == false)
                    {
                        //get window text
                        IntPtr textLengthPtr = UnsafeNativeFunctions.SendMessage(hwndEdit,
                        WindowMessages.WM_GETTEXTLENGTH, IntPtr.Zero, IntPtr.Zero);

                        string windowText = string.Empty;

                        if (textLengthPtr.ToInt32() > 0)
                        {
                            int textLength = textLengthPtr.ToInt32() + 1;
                            StringBuilder text = new StringBuilder(textLength);

                            UnsafeNativeFunctions.SendMessage(hwndEdit,
                                WindowMessages.WM_GETTEXT, textLength, text);

                            windowText = text.ToString();
                        }

                        double valueDouble = 0.0;
                        if (double.TryParse(windowText, out valueDouble) == true)
                        {
                            return valueDouble;
                        }
                    }
                    else
                    { 
                        // set spinner value
                        string valueString = string.Empty;
                        try
                        {
                            valueString = Convert.ToInt32(value).ToString();
                        }
                        catch 
                        {
                            return null;
                        }

                        IntPtr textPtr = Marshal.StringToBSTR(valueString);
                        if (textPtr != IntPtr.Zero)
                        {
                            UnsafeNativeFunctions.SendMessage(hwndEdit,
                                WindowMessages.WM_SETTEXT, IntPtr.Zero, textPtr);
							return 0;
                        }
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Gets the minimum value the current element spinner/slider/progressbar can get.
        /// </summary>
        /// <returns></returns>
        internal double GetMinimum()
        {
            string localizedControlType = "current element";

            try
            {
                localizedControlType = this.uiElement.CurrentLocalizedControlType;
            }
            catch { }

            object rangeValuePatternObj = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_RangeValuePatternId);
            IUIAutomationRangeValuePattern rangeValuePattern = rangeValuePatternObj as IUIAutomationRangeValuePattern;

            if (rangeValuePattern == null)
            {
                Engine.TraceInLogFile("Cannot get minimum of " + 
                    localizedControlType);

                throw new Exception("Cannot get minimum of " + 
                    localizedControlType);
            }

            try
            {
                return rangeValuePattern.CurrentMinimum;
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("Cannot get minimum of " + 
                    localizedControlType + ": " + ex.Message);

                throw new Exception("Cannot get minimum of " + 
                    localizedControlType + ": " + ex.Message);
            }
        }

        /// <summary>
        /// Gets the maximum value the current element spinner/slider/progressbar can get.
        /// </summary>
        /// <returns></returns>
        internal double GetMaximum()
        {
            string localizedControlType = "current element";

            try
            {
                localizedControlType = this.uiElement.CurrentLocalizedControlType;
            }
            catch { }

            object rangeValuePatternObj = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_RangeValuePatternId);
            IUIAutomationRangeValuePattern rangeValuePattern = rangeValuePatternObj as IUIAutomationRangeValuePattern;

            if (rangeValuePattern == null)
            {
                Engine.TraceInLogFile("Cannot get maximum of " + localizedControlType);
                throw new Exception("Cannot get maximum of " + localizedControlType);
            }

            try
            {
                return rangeValuePattern.CurrentMaximum;
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("Cannot get maximum of " + 
                    localizedControlType + ": " + ex.Message);

                throw new Exception("Cannot get maximum of " + 
                    localizedControlType + ": " + ex.Message);
            }
        }

        internal double GetSmallChange()
        {
            object rangeValuePatternObj = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_RangeValuePatternId);
            IUIAutomationRangeValuePattern rangeValuePattern = rangeValuePatternObj as IUIAutomationRangeValuePattern;

            if (rangeValuePattern == null)
            {
                Engine.TraceInLogFile("GetSmallChange() method: RangeValuePattern not supported");
                throw new Exception("GetSmallChange() method: RangeValuePattern not supported");
            }

            try
            {
                double smallChange = rangeValuePattern.CurrentSmallChange;

                return smallChange;
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("GetSmallChange() method failed: " + ex.Message);
                throw new Exception("GetSmallChange() method failed: " + ex.Message);
            }
        }

        internal double GetLargeChange()
        {
            object rangeValuePatternObj = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_RangeValuePatternId);
            IUIAutomationRangeValuePattern rangeValuePattern = rangeValuePatternObj as IUIAutomationRangeValuePattern;

            if (rangeValuePattern == null)
            {
                Engine.TraceInLogFile("GetLargeChange() method: RangeValuePattern not supported");
                throw new Exception("GetLargeChange() method: RangeValuePattern not supported");
            }

            try
            {
                double largeChange = rangeValuePattern.CurrentLargeChange;

                return largeChange;
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("GetLargeChange() method failed: " + ex.Message);
                throw new Exception("GetLargeChange() method failed: " + ex.Message);
            }
        }
    }
}
