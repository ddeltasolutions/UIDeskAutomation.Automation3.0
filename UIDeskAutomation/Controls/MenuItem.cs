using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Class that represents a Menu Item ui element
    /// </summary>
    public class UIDA_MenuItem : ElementBase
    {
        public UIDA_MenuItem(IUIAutomationElement el)
        {
            base.uiElement = el;
        }
        
        /// <summary>
        /// Accesses the menu item, like clicking on it
        /// </summary>
        public void AccessMenu()
        {
            base.Invoke();
        }

        /// <summary>
        /// Expands this menu item
        /// </summary>
        public void Expand()
        {
            if (this.IsAlive == false)
            {
                throw new Exception("This UI element is not available anymore.");
            }

            object objectPattern = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_ExpandCollapsePatternId);
            IUIAutomationExpandCollapsePattern expandCollapsePattern = objectPattern as IUIAutomationExpandCollapsePattern;

            if (expandCollapsePattern == null)
            {
                Engine.TraceInLogFile("Expand method - ExpandCollapse pattern not supported, Try Invoke");
                this.Invoke();
                return;
            }

            try
            {
                expandCollapsePattern.Expand();
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("Expand method - " + ex.Message);
                throw new Exception("Expand method - " + ex.Message);
            }
        }

        /// <summary>
        /// Collapses this menu item
        /// </summary>
        public void Collapse()
        {
            if (this.IsAlive == false)
            {
                throw new Exception("This UI element is not available anymore.");
            }

            object objectPattern = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_ExpandCollapsePatternId);
            IUIAutomationExpandCollapsePattern expandCollapsePattern = objectPattern as IUIAutomationExpandCollapsePattern;

            if (expandCollapsePattern == null)
            {
                Engine.TraceInLogFile("Collapse method - ExpandCollapse pattern not supported, Try Invoke");
                this.Invoke();
                return;
            }

            try
            {
                expandCollapsePattern.Collapse();
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("Collapse method - " + ex.Message);
                throw new Exception("Collapse method - " + ex.Message);
            }
        }

        /// <summary>
        /// Toggles a menu item (checked/unchecked)
        /// </summary>
        public void Toggle()
        {
            if (this.IsAlive == false)
            {
                throw new Exception("This UI element is not available to the user anymore.");
            }

            object objectPattern = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_TogglePatternId);
            IUIAutomationTogglePattern togglePattern = objectPattern as IUIAutomationTogglePattern;
            if (togglePattern == null)
            {
                //Engine.TraceInLogFile("Toggle method - Toggle pattern not supported");
				this.Invoke();
				return;
            }

            try
            {
                togglePattern.Toggle();
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("Toggle method - " + ex.Message);
                throw new Exception("Toggle method - " + ex.Message);
            }
        }

        /// <summary>
        /// Gets or Sets the checked state of a menu item 
        /// </summary>
        public bool IsChecked
        {
            get
            {
                if (this.IsAlive == false)
                {
                    throw new Exception("This UI element is not available anymore.");
                }

				//if (frameworkid == "Win32")
				//{
					object objectPattern = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_TogglePatternId);
                    IUIAutomationTogglePattern togglePattern = objectPattern as IUIAutomationTogglePattern;

					if (togglePattern == null)
					{
						//return false;
                        Engine.TraceInLogFile("Cannot get the checked state of the menu item");
                        throw new Exception("Cannot get the checked state of the menu item");
					}

					try
					{
						if (togglePattern.CurrentToggleState == ToggleState.ToggleState_On)
						{
							return true;
						}
						else
						{
							return false;
						}
						
                        /*else if (togglePattern.CurrentToggleState == ToggleState.ToggleState_Off)
                        {
                            return false;
                        }
                        else
                        {
                            return null;
                        }*/
					}
					catch (Exception ex)
					{
						Engine.TraceInLogFile("Checked (get) property - " + ex.Message);
						throw new Exception("Checked (get) property - " + ex.Message);
					}
				//}
            }

            set
            {
                if (this.IsAlive == false)
                {
                    throw new Exception("This UI element is not available anymore.");
                }
				
				object objectPattern = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_TogglePatternId);
                IUIAutomationTogglePattern togglePattern = objectPattern as IUIAutomationTogglePattern;
				if (togglePattern == null)
				{
					string frameworkid = "";
					try
					{
						frameworkid = this.uiElement.CurrentFrameworkId;
					}
					catch {}
					
					if (frameworkid == "Win32")
					{
						//current element is unchecked because it doesn't support TogglePattern
						if (value == true)
						{
							this.Invoke();
						}
						return;
					}
					else
					{
						Engine.TraceInLogFile("Checked (set) property - Toggle pattern not supported");
						throw new Exception("Checked (set) property - Toggle pattern not supported");
					}
				}

				try
				{
					/*if (((value == true) && (togglePattern.CurrentToggleState != ToggleState.ToggleState_On)) ||
						((value == false) && (togglePattern.CurrentToggleState == ToggleState.ToggleState_On)))*/
						
					if (((value == true) && (togglePattern.CurrentToggleState != ToggleState.ToggleState_On)) ||
						((value == false) && (togglePattern.CurrentToggleState != ToggleState.ToggleState_Off)))
					{
						togglePattern.Toggle();
					}
				}
				catch (Exception ex)
				{
					Engine.TraceInLogFile("Checked (set) property - " + ex.Message);
					throw new Exception("Checked (set) property - " + ex.Message);
				}
            }
        }
		
		/// <summary>
        /// Gets the text of the menu item
        /// </summary>
		public string Text
		{
			get
			{
				return this.GetText();
			}
		}
    }
}
