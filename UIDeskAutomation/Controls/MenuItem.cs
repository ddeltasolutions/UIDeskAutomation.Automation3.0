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
			try
			{
				base.Invoke();
			}
			catch
			{
				base.Click();
			}
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
		
		private UIA_AutomationEventHandler UIA_MenuOpenedEventHandler = null;
		private UIA_AutomationPropertyChangedEventHandler UIA_ExpandedEventHandler = null;
		
		/// <summary>
        /// Delegate for Expanded event
        /// </summary>
		/// <param name="sender">The menu item that sent the event.</param>
		public delegate void Expanded(UIDA_MenuItem sender);
		internal Expanded ExpandedHandler = null;
		
		/// <summary>
        /// Attaches/detaches a handler to expanded event
        /// </summary>
		public event Expanded ExpandedEvent
		{
			add
			{
				try
				{
					if (this.ExpandedHandler == null)
					{
						if (base.uiElement.CurrentFrameworkId == "WinForm")
						{
							this.UIA_MenuOpenedEventHandler = new UIA_AutomationEventHandler(this);
		
							Engine.uiAutomation.AddAutomationEventHandler(UIA_EventIds.UIA_MenuOpenedEventId, 
								base.uiElement, TreeScope.TreeScope_Subtree, null, this.UIA_MenuOpenedEventHandler);
						}
						else
						{
							this.UIA_ExpandedEventHandler = new UIA_AutomationPropertyChangedEventHandler(this);
				
							Engine.uiAutomation.AddPropertyChangedEventHandler(base.uiElement, TreeScope.TreeScope_Element, 
								null, this.UIA_ExpandedEventHandler, new int[] { UIA_PropertyIds.UIA_ExpandCollapseExpandCollapseStatePropertyId });
						}
					}
					
					this.ExpandedHandler += value;
				}
				catch {}
			}
			remove
			{
				try
				{
					this.ExpandedHandler -= value;
				
					if (this.ExpandedHandler == null)
					{
						if (base.uiElement.CurrentFrameworkId == "WinForm")
						{
							if (this.UIA_MenuOpenedEventHandler == null)
							{
								return;
							}
							
							System.Threading.Tasks.Task.Run(() => 
							{
								try
								{
									Engine.uiAutomation.RemoveAutomationEventHandler(UIA_EventIds.UIA_MenuOpenedEventId, 
										base.uiElement, this.UIA_MenuOpenedEventHandler);
									UIA_MenuOpenedEventHandler = null;
								}
								catch { }
							}).Wait(5000);
						}
						else
						{
							if (this.UIA_ExpandedEventHandler == null)
							{
								return;
							}
							
							System.Threading.Tasks.Task.Run(() => 
							{
								try
								{
									Engine.uiAutomation.RemovePropertyChangedEventHandler(base.uiElement, 
										this.UIA_ExpandedEventHandler);
									UIA_ExpandedEventHandler = null;
								}
								catch { }
							}).Wait(5000);
						}
					}
				}
				catch {}
			}
		}
		
		private UIA_AutomationEventHandler UIA_ClickedEventHandler = null;
		
		/// <summary>
        /// Delegate for Clicked event
        /// </summary>
		/// <param name="sender">The menu item that sent the event.</param>
		public delegate void Clicked(UIDA_MenuItem sender);
		internal Clicked ClickedHandler = null;
		
		/// <summary>
        /// Attaches/detaches a handler to click event
        /// </summary>
		public event Clicked ClickedEvent
		{
			add
			{
				try
				{
					if (this.ClickedHandler == null)
					{
						this.UIA_ClickedEventHandler = new UIA_AutomationEventHandler(this);
		
						Engine.uiAutomation.AddAutomationEventHandler(UIA_EventIds.UIA_Invoke_InvokedEventId, 
							base.uiElement, TreeScope.TreeScope_Element, null, this.UIA_ClickedEventHandler);
					}
					
					this.ClickedHandler += value;
				}
				catch {}
			}
			remove
			{
				try
				{
					this.ClickedHandler -= value;
				
					if (this.ClickedHandler == null)
					{
						if (this.UIA_ClickedEventHandler == null)
						{
							return;
						}
						
						System.Threading.Tasks.Task.Run(() => 
						{
							try
							{
								Engine.uiAutomation.RemoveAutomationEventHandler(UIA_EventIds.UIA_Invoke_InvokedEventId, 
									base.uiElement, this.UIA_ClickedEventHandler);
								UIA_ClickedEventHandler = null;
							}
							catch { }
						}).Wait(5000);
					}
				}
				catch {}
			}
		}
    }
}
