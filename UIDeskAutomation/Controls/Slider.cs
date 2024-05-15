using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Class that represents a Slider UI element.
    /// </summary>
    public class UIDA_Slider: GenericSpinner
    {
        public UIDA_Slider(IUIAutomationElement el)
            : base(el)
        { }
		
		/// <summary>
        /// Increments the value of Slider. Is like pressing the arrow "Right" key.
        /// </summary>
        public void SmallIncrement()
        {
            this.KeyPress(VirtualKeys.Right);
        }

        /// <summary>
        /// Increments the value of Slider. Is like pressing in the right side of the thumb or "Page Up" key.
        /// </summary>
        public void LargeIncrement()
        {
            UIDA_Button increaseButton = null;

            try
            {
				if (this.uiElement.CurrentFrameworkId == "WinForm")
				{
					increaseButton = this.Button("Page right");
				}
				else
				{
					increaseButton = this.ButtonAt(null, 2);
				}
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("Slider.Increment method failed");
                throw new Exception("Slider.Increment method failed");
            }

            if (increaseButton != null)
            {
                increaseButton.Invoke();
            }
        }
		
		/// <summary>
        /// Decrements the value of Slider. Is like pressing the arrow "Left" key.
        /// </summary>
        public void SmallDecrement()
        {
            this.KeyPress(VirtualKeys.Left);
        }

        /// <summary>
        /// Decrements the value of Slider. Is like pressing in the left side of the thumb or "Page Down" key.
        /// </summary>
        public void LargeDecrement()
        {
            UIDA_Button decreaseButton = null;
            try
            {
				if (this.uiElement.CurrentFrameworkId == "WinForm")
				{
					decreaseButton = this.Button("Page left");
				}
				else
				{
					decreaseButton = this.ButtonAt(null, 1);
				}
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("Slider.Decrement method failed");
                throw new Exception("Slider.Decrement method failed");
            }

            if (decreaseButton != null)
            {
                decreaseButton.Invoke();
            }
        }
        
        /// <summary>
        /// Gets the minimum value the current slider can get.
        /// </summary>
        /// <returns>The minimum value of the slider</returns>
        new public double GetMinimum()
        {
            return base.GetMinimum();
        }
        
        /// <summary>
        /// Gets the maximum value the current slider can get.
        /// </summary>
        /// <returns>The maximum value of the slider</returns>
        new public double GetMaximum()
        {
            return base.GetMaximum();
        }
        
        /// <summary>
        /// Gets/Sets the value of the slider.
        /// </summary>
        new public double Value
        {
            get { return base.Value; }
            set { base.Value = value; }
        }
		
		/// <summary>
        /// Attaches/detaches a handler to value changed event. You can cast the first parameter (sender - of type GenericSpinner) to an UIDA_Slider object.
		/// The second parameter (of type double) is the new value of the slider.
        /// </summary>
		public event ValueChanged ValueChangedEvent
		{
			add
			{
				base.ValueChangedEvent += value;
			}
			remove
			{
				base.ValueChangedEvent -= value;
			}
		}
    }
}
