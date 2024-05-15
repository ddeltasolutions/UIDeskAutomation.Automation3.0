using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// This class represents a Group.
    /// </summary>
    public class UIDA_Group : ElementBase
    {
        /// <summary>
        /// Creates a UIDA_Group using an IUIAutomationElement
        /// </summary>
        /// <param name="el">UI Automation Element</param>
        public UIDA_Group(IUIAutomationElement el)
        {
			this.uiElement = el;
		}

        /// <summary>
        /// Expands a group.
        /// </summary>
        public void Expand()
        {
            IUIAutomationExpandCollapsePattern expandCollapsePattern = this.GetExpandCollapsePattern();

            if (expandCollapsePattern == null)
            {
                Engine.TraceInLogFile(
                    "Group.Expand() - ExpandCollapsePattern not supported");

                throw new Exception(
                    "Group.Expand() - ExpandCollapsePattern not supported");
            }

            try
            {
                if (expandCollapsePattern.CurrentExpandCollapseState !=
                    ExpandCollapseState.ExpandCollapseState_Expanded)
                {
                    expandCollapsePattern.Expand();
                }
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("Group.Expand() error: " + ex.Message);
                throw new Exception("Group.Expand() error: " + ex.Message);
            }
        }

        /// <summary>
        /// Collapses a group.
        /// </summary>
        public void Collapse()
        {
            IUIAutomationExpandCollapsePattern expandCollapsePattern =
                this.GetExpandCollapsePattern();

            if (expandCollapsePattern == null)
            {
                Engine.TraceInLogFile(
                    "Group.Collapse() - ExpandCollapsePattern not supported");

                throw new Exception(
                    "Group.Collapse() - ExpandCollapsePattern not supported");
            }

            try
            {
                if (expandCollapsePattern.CurrentExpandCollapseState !=
                    ExpandCollapseState.ExpandCollapseState_Collapsed)
                {
                    expandCollapsePattern.Collapse();
                }
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("Group.Collapse() error: " + ex.Message);
                throw new Exception("Group.Collapse() error: " + ex.Message);
            }
        }

        private IUIAutomationExpandCollapsePattern GetExpandCollapsePattern()
        {
            object expandCollapsePatternObj = uiElement.GetCurrentPattern(UIA_PatternIds.UIA_ExpandCollapsePatternId);
            
            IUIAutomationExpandCollapsePattern expandCollapsePattern =
                    expandCollapsePatternObj as IUIAutomationExpandCollapsePattern;

            return expandCollapsePattern;
        }
    }
}
