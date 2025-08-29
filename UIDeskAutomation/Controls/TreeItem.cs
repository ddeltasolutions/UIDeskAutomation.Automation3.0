using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Class that represents a TreeView Item.
    /// </summary>
    public class UIDA_TreeItem: ElementBase
    {
        public UIDA_TreeItem(IUIAutomationElement el)
        {
            this.uiElement = el;
        }

        /// <summary>
        /// Gets a collection of subitems of this TreeItem.
        /// </summary>
        /// <returns>UIDA_TreeItem array</returns>
        public UIDA_TreeItem[] SubItems
        {
            get
            {
                List<UIDA_TreeItem> tvItems = new List<UIDA_TreeItem>();

                List<IUIAutomationElement> subItems =
                    this.FindAll(UIA_ControlTypeIds.UIA_TreeItemControlTypeId, null, false, false, true);

                foreach (IUIAutomationElement el in subItems)
                {
                    UIDA_TreeItem tvItem = new UIDA_TreeItem(el);
                    tvItems.Add(tvItem);
                }

                return tvItems.ToArray();
            }
        }

        /// <summary>
        /// Expands a TreeItem.
        /// </summary>
        public void Expand()
        {
            object expandCollapsePatternObj = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_ExpandCollapsePatternId);
            IUIAutomationExpandCollapsePattern expandCollapsePattern = expandCollapsePatternObj as IUIAutomationExpandCollapsePattern;

            if (expandCollapsePattern == null)
            {
                Engine.TraceInLogFile("TreeItem.Expand method failed");
                throw new Exception("TreeItem.Expand method failed");
            }

            expandCollapsePattern.Expand();
        }

        /// <summary>
        /// Collapses a TreeItem.
        /// </summary>
        public void Collapse()
        {
            object expandCollapsePatternObj = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_ExpandCollapsePatternId);
            IUIAutomationExpandCollapsePattern expandCollapsePattern = expandCollapsePatternObj as IUIAutomationExpandCollapsePattern;
 
            if (expandCollapsePattern == null)
            {
                Engine.TraceInLogFile("TreeItem.Collapse method failed");
                throw new Exception("TreeItem.Collapse method failed");
            }

            expandCollapsePattern.Collapse();
        }
		
		/// <summary>
        /// Gets the ExpandCollapseState of this TreeItem.
        /// </summary>
        /// <returns>ExpandCollapseState</returns>
		public ExpandCollapseState ExpandCollapseState
		{
			get
			{
				object expandCollapsePatternObj = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_ExpandCollapsePatternId);
                IUIAutomationExpandCollapsePattern expandCollapsePattern = expandCollapsePatternObj as IUIAutomationExpandCollapsePattern;
                
                if (expandCollapsePattern == null)
                {
                    Engine.TraceInLogFile("TreeItem.ExpandCollapseState property failed");
                    throw new Exception("TreeItem.ExpandCollapseState property failed");
                }
                
                return expandCollapsePattern.CurrentExpandCollapseState;
			}
		}
		
		/// <summary>
        /// Returns true if the TreeItem is expanded.
        /// </summary>
        /// <returns>boolean</returns>
		public bool IsExpanded
		{
			get
			{
				object expandCollapsePatternObj = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_ExpandCollapsePatternId);
                IUIAutomationExpandCollapsePattern expandCollapsePattern = expandCollapsePatternObj as IUIAutomationExpandCollapsePattern;
                
                if (expandCollapsePattern == null)
                {
                    Engine.TraceInLogFile("TreeItem.IsExpanded property failed");
                    throw new Exception("TreeItem.IsExpanded property failed");
                }
                
                return expandCollapsePattern.CurrentExpandCollapseState == ExpandCollapseState.ExpandCollapseState_Expanded;
			}
		}

        /// <summary>
        /// Cycles through the check states (checked, unchecked, indeterminate).
        /// </summary>
        public void Toggle()
        {
			object togglePatternObj = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_TogglePatternId);
            IUIAutomationTogglePattern togglePattern = togglePatternObj as IUIAutomationTogglePattern;
            
            if (togglePattern == null)
            {
                Engine.TraceInLogFile("TreeItem.Toggle() failed");
                throw new Exception("TreeItem.Toggle() failed");
            }
            
            try
            {
                togglePattern.Toggle();
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("TreeItem.Toggle() error: " + ex.Message);
                throw new Exception("TreeItem.Toggle() error: " + ex.Message);
            }
        }

        /// <summary>
        /// Selects the current tree item and deselects all other selected tree items.
        /// </summary>
        public void Select()
        {
			this.BringToForeground();
			try
			{
				this.BringIntoView();
			}
			catch { }
			
			// click the tree item to make sure all events are raised
			string fwkid = this.uiElement.CurrentFrameworkId;
			if (fwkid == "WPF")
			{
				UIDA_Text label = this.Text();
				if (label != null)
				{
					label.Click();
				}
				else
				{
					this.Click();
				}
			}
			else
			{
				this.Click();
			}
			
            /*object selectionItemPatternObj = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_SelectionItemPatternId);
            IUIAutomationSelectionItemPattern selectionItemPattern = selectionItemPatternObj as IUIAutomationSelectionItemPattern;

            if (selectionItemPattern == null)
            {
                Engine.TraceInLogFile("TreeItem.Select() method failed");
                throw new Exception("TreeItem.Select() method failed");
            }

            selectionItemPattern.Select();*/
        }

        /// <summary>
        /// Brings the current Tree Item into viewable area of the parent Tree control.
        /// </summary>
        public void BringIntoView()
        {
            object scrollItemPatternObj = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_ScrollItemPatternId);
            IUIAutomationScrollItemPattern scrollItemPattern = scrollItemPatternObj as IUIAutomationScrollItemPattern;

            if (scrollItemPattern == null)
            {
                Engine.TraceInLogFile("TreeItem.BringIntoView method failed");
                throw new Exception("TreeItem.BringIntoView method failed");
            }

            try
            {
                scrollItemPattern.ScrollIntoView();
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("TreeItem.BringIntoView method failed: " + 
                    ex.Message);
                throw new Exception("TreeItem.BringIntoView method failed: " +
                    ex.Message);
            }
        }
		
		/// <summary>
        /// Gets or sets the checked state of the current tree item if supported.
        /// </summary>
		/// <returns>true if tree item is checked, false if unchecked, null if it's in an indeterminate state</returns>
        public bool? IsChecked
        {
            get
            {
                object togglePatternObj = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_TogglePatternId);
                IUIAutomationTogglePattern togglePattern = togglePatternObj as IUIAutomationTogglePattern;

                if (togglePattern == null)
                {
                    Engine.TraceInLogFile("TreeItem.Checked failed");
                    throw new Exception("TreeItem.Checked failed");
                }

                try
                {
                    ToggleState toggleState = togglePattern.CurrentToggleState;
                    if (toggleState == ToggleState.ToggleState_On)
                    {
                        return true;
                    }
                    else if (toggleState == ToggleState.ToggleState_Off)
                    {
                        return false;
                    }
                    else // ToggleState_Indeterminate
                    {
                        return null;
                    }
                }
                catch (Exception ex)
                { 
                    Engine.TraceInLogFile("TreeItem.Checked failed: " + ex.Message);
                    throw new Exception("TreeItem.Checked failed: " + ex.Message);
                }
            }

            set
            {
                object togglePatternObj = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_TogglePatternId);
                IUIAutomationTogglePattern togglePattern = togglePatternObj as IUIAutomationTogglePattern;

                if (togglePattern == null)
                {
                    if (value == true)
                    {
                        Engine.TraceInLogFile("Cannot check tree item");
                        throw new Exception("Cannot check tree item");
                    }
                    else if (value == false)
                    {
                        Engine.TraceInLogFile("Cannot uncheck tree item");
                        throw new Exception("Cannot uncheck tree item");
                    }
					else
					{
						Engine.TraceInLogFile("Cannot set tree item checked state");
                        throw new Exception("Cannot set tree item checked state");
					}
                }
                
                ToggleState desiredState = ToggleState.ToggleState_Indeterminate;
                if (value == true)
                {
                    desiredState = ToggleState.ToggleState_On;
                }
                else if (value == false)
                {
                    desiredState = ToggleState.ToggleState_Off;
                }
                
                int tries = 3;
                while (tries > 0)
                {
                    if (togglePattern.CurrentToggleState == desiredState)
                    {
                        break;
                    }
                    togglePattern.Toggle();
                    tries--;
                }
                
                if (tries == 0)
                {
                    if (value == true)
                    {
                        Engine.TraceInLogFile("Cannot check tree item");
                        throw new Exception("Cannot check tree item");
                    }
                    else if (value == false)
                    { 
                        Engine.TraceInLogFile("Cannot uncheck tree item");
                        throw new Exception("Cannot uncheck tree item");
                    }
					else
					{
						Engine.TraceInLogFile("Cannot set tree item checked state");
                        throw new Exception("Cannot set tree item checked state");
					}
                }
            }
        }
		
		private UIA_AutomationPropertyChangedEventHandler UIA_ExpandedEventHandler = null;
		
		/// <summary>
        /// Delegate for Expanded event
        /// </summary>
		/// <param name="sender">The tree item that sent the event.</param>
		/// <param name="isExpanded">true if expanded, false if not.</param>
		public delegate void Expanded(UIDA_TreeItem sender, bool isExpanded);
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
						this.UIA_ExpandedEventHandler = new UIA_AutomationPropertyChangedEventHandler(this);
			
						Engine.uiAutomation.AddPropertyChangedEventHandler(base.uiElement, TreeScope.TreeScope_Element, 
							null, this.UIA_ExpandedEventHandler, new int[] { UIA_PropertyIds.UIA_ExpandCollapseExpandCollapseStatePropertyId });
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
				catch {}
			}
		}
    }
}
