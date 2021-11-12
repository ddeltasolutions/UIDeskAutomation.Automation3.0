using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a ScrollBar control.
    /// </summary>
    public class UIDA_ScrollBar: GenericSpinner
    {
        public UIDA_ScrollBar(IUIAutomationElement el): base(el)
        {
            //this.uiElement = el;
        }

        /// <summary>
        /// Increments the value of ScrollBar. 
        /// Is like pressing the right button/arrow (for horizontal scrollbar) or 
        /// the bottom button/down arrow (for vertical scrollbar) of the ScrollBar control.
        /// </summary>
        public void SmallIncrement()
        {
            double smallChange = 0.0;

            try
            {
                smallChange = base.GetSmallChange();
                base.Value += smallChange;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Increments scrollbar control with a larger value.
        /// Is like pressing the region between the thumb and the right arrow at 
        /// a horizontal scrollbar or pressing the region between the thumb and the 
        /// down arrow for a vertical scrollbar.
        /// </summary>
        public void LargeIncrement()
        {
            try
            {
                double largeChange = base.GetLargeChange();
                base.Value += largeChange;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Decrements the value of ScrollBar. 
        /// Is like pressing the left button (arrow) for a horizontal scrollbar or 
        /// the top button (up arrow) for a vertical scrollbar.
        /// </summary>
        public void SmallDecrement()
        {
            try
            {
                double smallChange = base.GetSmallChange();
                base.Value -= smallChange;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Decrements the value of a scrollbar with a larger value.
        /// Is like pressing the region between left arrow and the thumb for a 
        /// horizontal scrollbar or like pressing the region between up arrow and the 
        /// thumb for a vertical scrollbar.
        /// </summary>
        public void LargeDecrement()
        {
            try
            {
                double largeChange = base.GetLargeChange();
                base.Value -= largeChange;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Gets the minimum value the current scrollbar can get.
        /// </summary>
        /// <returns>The minimum value of the scrollbar</returns>
        new public double GetMinimum()
        {
            return base.GetMinimum();
        }
        
        /// <summary>
        /// Gets the maximum value the current scrollbar can get.
        /// </summary>
        /// <returns>The maximum value of the scrollbar</returns>
        new public double GetMaximum()
        {
            return base.GetMaximum();
        }
        
        /// <summary>
        /// Gets/Sets the value of the current scrollbar.
        /// </summary>
        new public double Value
        {
            get { return base.Value; }
            set { base.Value = value; }
        }
    }
}
