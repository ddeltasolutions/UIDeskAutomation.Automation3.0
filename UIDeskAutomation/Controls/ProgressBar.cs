using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents the class for a Progress Bar UI control.
    /// </summary>
    public class UIDA_ProgressBar: GenericSpinner
    {
        public UIDA_ProgressBar(IUIAutomationElement el): base(el)
        { }
        
        /// <summary>
        /// Gets the value of the current progressbar.
        /// </summary>
        new public double Value
        {
            get { return base.Value; }
        }
        
        /// <summary>
        /// Gets the minimum value the current progressbar can get.
        /// </summary>
        /// <returns>The minimum value of the progressbar</returns>
        new public double GetMinimum()
        {
            return base.GetMinimum();
        }
        
        /// <summary>
        /// Gets the maximum value the current progressbar can get.
        /// </summary>
        /// <returns>The maximum value of the progressbar</returns>
        new public double GetMaximum()
        {
            return base.GetMaximum();
        }
		
		/// <summary>
        /// Attaches/detaches a handler to value changed event. You can cast the first parameter (sender - of type GenericSpinner) to an UIDA_ProgressBar object.
		/// The second parameter (of type double) is the new value of the progress bar.
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
