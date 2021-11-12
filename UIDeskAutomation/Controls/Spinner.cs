using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a spinner UI control.
    /// </summary>
    public class UIDA_Spinner: GenericSpinner
    {
        public UIDA_Spinner(IUIAutomationElement el): base(el)
        {
            //this.uiElement = el;
        }

        /// <summary>
        /// Increment the value of spinner. Is like pressing the up arrow.
        /// </summary>
        public void Increment()
        {
            UIDA_Button forwardButton = null;

            try
            {
                forwardButton = this.ButtonAt(null, 1/*, true*/);
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("Spinner.Increment method failed: " + ex.Message);
                throw ex;
            }

            if (forwardButton != null)
            {
                if (uiElement.CurrentFrameworkId == "WinForm" || uiElement.CurrentFrameworkId == "Win32")
                {
                    forwardButton.SimulateClick();
                    return;
                }
                
                try
                {
                    forwardButton.Invoke();
                }
                catch (Exception ex)
                {
                    Engine.TraceInLogFile("Spinner.Increment Invoke failed: " + ex.Message);
                    throw ex;
                }
            }
        }

        /// <summary>
        /// Decrements the value of spinner. Is like pressing the down arrow.
        /// </summary>
        public void Decrement()
        {
            UIDA_Button backwardButton = null;
            try
            {
                backwardButton = this.ButtonAt(null, 2/*, true*/);
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("Spinner.Decrement method failed: " + ex.Message);
                throw ex;
            }

            if (backwardButton != null)
            {
                if (uiElement.CurrentFrameworkId == "WinForm" || uiElement.CurrentFrameworkId == "Win32")
                {
                    backwardButton.SimulateClick();
                    return;
                }
                
                try
                {
                    backwardButton.Invoke();
                }
                catch (Exception ex)
                {
                    Engine.TraceInLogFile("Spinner.Decrement Invoke failed: " + ex.Message);
                    throw ex;
                }
            }
        }
        
        /// <summary>
        /// Gets the minimum value the current spinner can get.
        /// </summary>
        /// <returns>The minimum value of the spinner</returns>
        new public double GetMinimum()
        {
            return base.GetMinimum();
        }
        
        /// <summary>
        /// Gets the maximum value the current spinner can get.
        /// </summary>
        /// <returns>The maximum value of the spinner</returns>
        new public double GetMaximum()
        {
            return base.GetMaximum();
        }
        
        /// <summary>
        /// Gets/Sets the value of the current spinner.
        /// </summary>
        new public double Value
        {
            get 
            {
                if (uiElement.CurrentFrameworkId == "WinForm")
                {
                    IUIAutomationTreeWalker tw = Engine.uiAutomation.ControlViewWalker;
                    IUIAutomationElement parent = tw.GetParentElement(uiElement);
                    UIDA_ComboBox combo = new UIDA_ComboBox(parent);
                    UIDA_Edit edit = combo.Edit();
                    return Convert.ToDouble(edit.Text);
                }
                else if (uiElement.CurrentFrameworkId == "Win32")
                {
                    BringToForeground();
                    tagRECT rect = uiElement.CurrentBoundingRectangle;
                    tagPOINT pt;
                    pt.x = rect.left - 5;
                    pt.y = (rect.top + rect.bottom) / 2;
                    IUIAutomationElement editEl = Engine.uiAutomation.ElementFromPoint(pt);
                    UIDA_Edit edit = new UIDA_Edit(editEl);
                    return Convert.ToDouble(edit.Text);
                }
                else
                {
                    return base.Value; 
                }
            }
            set 
            {
                if (uiElement.CurrentFrameworkId == "WinForm")
                {
                    IUIAutomationTreeWalker tw = Engine.uiAutomation.ControlViewWalker;
                    IUIAutomationElement parent = tw.GetParentElement(uiElement);
                    UIDA_ComboBox combo = new UIDA_ComboBox(parent);
                    UIDA_Edit edit = combo.Edit();
                    edit.Text = value.ToString();
                }
                else if (uiElement.CurrentFrameworkId == "Win32")
                {
                    BringToForeground();
                    tagRECT rect = uiElement.CurrentBoundingRectangle;
                    tagPOINT pt;
                    pt.x = rect.left - 5;
                    pt.y = (rect.top + rect.bottom) / 2;
                    IUIAutomationElement editEl = Engine.uiAutomation.ElementFromPoint(pt);
                    UIDA_Edit edit = new UIDA_Edit(editEl);
                    edit.Text = value.ToString();
                }
                else
                {
                    base.Value = value; 
                }
            }
        }
    }
}
