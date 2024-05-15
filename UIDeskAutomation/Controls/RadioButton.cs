using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using System.Windows.Automation;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a RadioButton UI element.
    /// </summary>
    public class UIDA_RadioButton: ElementBase
    {
        public UIDA_RadioButton(IUIAutomationElement el)
        {
            this.uiElement = el;
        }

        /// <summary>
        /// Gets the selected state of the radio button.
        /// </summary>
        /// <returns>true - is selected, false otherwise</returns>
        public bool IsSelected
        {
			get
			{
				object selectionItemPatternObj = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_SelectionItemPatternId);
				IUIAutomationSelectionItemPattern selectionItemPattern = selectionItemPatternObj as IUIAutomationSelectionItemPattern;

				if (selectionItemPattern == null)
				{
					Engine.TraceInLogFile("RadioButton.GetIsSelected() - SelectionItemPattern not supported");
					throw new Exception("RadioButton.GetIsSelected() - SelectionItemPattern not supported");
				}

				return (selectionItemPattern.CurrentIsSelected != 0);
			}
        }

        /// <summary>
        /// Selects a radio button.
        /// </summary>
        public void Select()
        {
			this.Click();
			//Engine.GetInstance().Sleep(100);
			
			/*if (this.uiElement.CurrentFrameworkId == "WPF")
			{
				// for WPF use mouse click because IUIAutomationSelectionItemPattern.Select() is not calling the radio button's Click event handler
				this.Click();
				return;
			}
			
            object selectionItemPatternObj = this.uiElement.GetCurrentPattern(UIA_PatternIds.UIA_SelectionItemPatternId);
            IUIAutomationSelectionItemPattern selectionItemPattern = selectionItemPatternObj as IUIAutomationSelectionItemPattern;

            if (selectionItemPattern == null)
            {
                Engine.TraceInLogFile("RadioButton.Select() - SelectionItemPattern not supported");
                throw new Exception("RadioButton.Select() - SelectionItemPattern not supported");
            }

            selectionItemPattern.Select(); */
        }
		
		/// <summary>
        /// Gets the text of the radio button
        /// </summary>
		public string Text
		{
			get
			{
				return this.GetText();
			}
		}
		
		private UIA_AutomationPropertyChangedEventHandler UIA_StateChangedEventHandler = null;
		private UIA_AutomationEventHandler UIA_StateChangedEventHandlerWin32 = null;
		
		/// <summary>
        /// Delegate for State Changed event
        /// </summary>
		/// <param name="sender">The radio button that sent the event.</param>
		/// <param name="isChecked">true if the radio button is checked, false if it's unchecked.</param>
		public delegate void StateChanged(UIDA_RadioButton sender, bool isChecked);
		internal StateChanged StateChangedHandler = null;
		
		/// <summary>
        /// Attaches/detaches a handler to checked state changed event
        /// </summary>
		public event StateChanged StateChangedEvent
		{
			add
			{
				try
				{
					if (this.StateChangedHandler == null)
					{
						string cfid = base.uiElement.CurrentFrameworkId;
						
						if (cfid == "Win32")
						{
							this.UIA_StateChangedEventHandlerWin32 = new UIA_AutomationEventHandler(this);
							
							Engine.uiAutomation.AddAutomationEventHandler(UIA_EventIds.UIA_Invoke_InvokedEventId, 
								base.uiElement, TreeScope.TreeScope_Element, null, this.UIA_StateChangedEventHandlerWin32);
						}
						else if (cfid == "WinForm")
						{
							this.UIA_StateChangedEventHandler = new UIA_AutomationPropertyChangedEventHandler(this);
							
							Engine.uiAutomation.AddPropertyChangedEventHandler(base.uiElement, TreeScope.TreeScope_Element, 
								null, this.UIA_StateChangedEventHandler, new int[] { UIA_PropertyIds.UIA_NamePropertyId });
						}
						else
						{
							this.UIA_StateChangedEventHandler = new UIA_AutomationPropertyChangedEventHandler(this);
							
							Engine.uiAutomation.AddPropertyChangedEventHandler(base.uiElement, TreeScope.TreeScope_Element, 
								null, this.UIA_StateChangedEventHandler, new int[] { UIA_PropertyIds.UIA_SelectionItemIsSelectedPropertyId });
						}
					}
					
					this.StateChangedHandler += value;
				}
				catch {}
			}
			remove
			{
				try
				{
					this.StateChangedHandler -= value;
				
					if (this.StateChangedHandler == null)
					{
						if (base.uiElement.CurrentFrameworkId == "Win32")
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
		
		private void RemoveEventHandlerWin32()
		{
			if (this.UIA_StateChangedEventHandlerWin32 == null)
			{
				return;
			}
			
			System.Threading.Tasks.Task.Run(() => 
			{
				try
				{
					Engine.uiAutomation.RemoveAutomationEventHandler(UIA_EventIds.UIA_Invoke_InvokedEventId, 
						base.uiElement, this.UIA_StateChangedEventHandlerWin32);
					UIA_StateChangedEventHandlerWin32 = null;
				}
				catch { }
			}).Wait(5000);
		}
		
		private void RemoveEventHandler()
		{
			if (this.UIA_StateChangedEventHandler == null)
			{
				return;
			}
			
			System.Threading.Tasks.Task.Run(() => 
			{
				try
				{
					Engine.uiAutomation.RemovePropertyChangedEventHandler(base.uiElement, 
						this.UIA_StateChangedEventHandler);
					UIA_StateChangedEventHandler = null;
				}
				catch { }
			}).Wait(5000);
		}
    }
}
