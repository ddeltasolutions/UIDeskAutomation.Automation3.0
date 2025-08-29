using System;
using System.Collections.Generic;
using System.Threading;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    partial class ElementBase
    {
		internal List<IUIAutomationElement> FindAll(int type, string name,
            bool searchDescendants, bool bSearchByLabel, bool caseSensitive)
        {
            IUIAutomationCondition typeCondition = Engine.uiAutomation.CreatePropertyCondition(
				UIA_PropertyIds.UIA_ControlTypePropertyId, type);
            TreeScope scope = TreeScope.TreeScope_Children;

            if (searchDescendants)
            {
                scope = TreeScope.TreeScope_Descendants;
            }

            IUIAutomationElementArray collection = this.uiElement.FindAll(scope, typeCondition);

            if (collection == null)
            {
                return null;
            }

            List<IUIAutomationElement> foundElements = Helper.MatchStrings(collection,
                name, bSearchByLabel, caseSensitive);

            return foundElements;
        }
        
        internal List<IUIAutomationElement> FindAllPlusCondition(int type, IUIAutomationCondition cond, 
            string name, bool searchDescendants, bool bSearchByLabel, bool caseSensitive)
        {
            IUIAutomationCondition typeCondition = Engine.uiAutomation.CreatePropertyCondition(
				UIA_PropertyIds.UIA_ControlTypePropertyId, type);
            IUIAutomationCondition orCond = Engine.uiAutomation.CreateOrCondition(typeCondition, cond);

            TreeScope scope = TreeScope.TreeScope_Children;

            if (searchDescendants)
            {
                scope = TreeScope.TreeScope_Descendants;
            }

            IUIAutomationElementArray collection = this.uiElement.FindAll(scope, orCond);

            if (collection == null)
            {
                return null;
            }

            List<IUIAutomationElement> foundElements = Helper.MatchStrings(collection,
                name, bSearchByLabel, caseSensitive);

            return foundElements;
        }

        internal List<IUIAutomationElement> FindAllCustom(int type, string className, 
            string name, bool searchDescendants, bool bSearchByLabel, 
            bool caseSensitive)
        {
            IUIAutomationCondition typeCondition = Engine.uiAutomation.CreatePropertyCondition(
				UIA_PropertyIds.UIA_ControlTypePropertyId, type);

            IUIAutomationCondition classCondition = Engine.uiAutomation.CreatePropertyCondition(
				UIA_PropertyIds.UIA_ClassNamePropertyId, className);

            IUIAutomationCondition condition = Engine.uiAutomation.CreateAndCondition(typeCondition, classCondition);
            TreeScope scope = TreeScope.TreeScope_Children;

            if (searchDescendants)
            {
                scope = TreeScope.TreeScope_Descendants;
            }

            IUIAutomationElementArray collection = this.uiElement.FindAll(scope,
                condition);

            if (collection == null)
            {
                return null;
            }

            List<IUIAutomationElement> foundElements = Helper.MatchStrings(collection,
                name, bSearchByLabel, caseSensitive);

            return foundElements;
        }

        private Errors FindCustomAt(int type, string className,
            string name, int index, bool searchDescendants, bool searchByLabel,
            bool caseSensitive, out IUIAutomationElement returnElement)
        {
            IUIAutomationCondition typeCondition = Engine.uiAutomation.CreatePropertyCondition(
				UIA_PropertyIds.UIA_ControlTypePropertyId, type);

            return FindAtWithConditionAndClassName(typeCondition, name, className, 
                index, searchDescendants, searchByLabel, caseSensitive, out returnElement);
        }

        internal Errors FindAt(int type, string name, int index,
            bool searchDescendants, bool bSearchByLabel, bool caseSensitive,
            out IUIAutomationElement returnElement)
        {
            IUIAutomationCondition typeCondition = Engine.uiAutomation.CreatePropertyCondition(
				UIA_PropertyIds.UIA_ControlTypePropertyId, type);

            return FindAtWithCondition(typeCondition, name, index, searchDescendants,
                bSearchByLabel, caseSensitive, out returnElement);
        }
        
        internal Errors FindAtPlusCondition(int type, IUIAutomationCondition cond, string name, int index,
            bool searchDescendants, bool bSearchByLabel, bool caseSensitive,
            out IUIAutomationElement returnElement)
        {
            IUIAutomationCondition typeCondition = Engine.uiAutomation.CreatePropertyCondition(
				UIA_PropertyIds.UIA_ControlTypePropertyId, type);
            IUIAutomationCondition andCond = Engine.uiAutomation.CreateOrCondition(typeCondition, cond);

            return FindAtWithCondition(andCond, name, index, searchDescendants,
                bSearchByLabel, caseSensitive, out returnElement);
        }

        private Errors FindAtWithCondition(IUIAutomationCondition condition, string name, int index,
            bool searchDescendants, bool bSearchByLabel, bool caseSensitive,
            out IUIAutomationElement returnElement)
        {
            TreeScope scope = TreeScope.TreeScope_Children;

            if (searchDescendants)
            {
                scope = TreeScope.TreeScope_Descendants;
            }

            if (index == 0)
            {
                index = 1;
            }

            int nWaitMs = Engine.GetInstance().Timeout;
            IUIAutomationElementArray collection = null;

            List<IUIAutomationElement> foundElements = null;

            //while (nWaitMs > 0)
			while (true)
            {
                collection = this.uiElement.FindAll(scope, condition);

                if ((collection != null) && (collection.Length >= index))
                {
                    foundElements = Helper.MatchStrings(collection, name,
                        bSearchByLabel, caseSensitive);

                    if ((foundElements != null) && (foundElements.Count >= index))
                    {
                        break;
                    }
                }
				
				/*if (Engine.IsCancelled == true)
				{
					Engine.IsCancelled = false;
					break;
				}*/

                nWaitMs -= ElementBase.waitPeriod;
				if (nWaitMs <= 0)
				{
					break;
				}
                Thread.Sleep(ElementBase.waitPeriod);
            }

            if ((foundElements == null) || (foundElements.Count == 0))
            {
                returnElement = null;
                return Errors.ElementNotFound;
            }

            if (index <= foundElements.Count)
            {
                returnElement = foundElements[index - 1];
                return Errors.None;
            }
            else
            {
                returnElement = null;
                return Errors.IndexTooBig;
            }
        }
        
        private Errors FindAtWithConditionAndClassName(IUIAutomationCondition typeCondition, 
            string name, string className, int index,
            bool searchDescendants, bool bSearchByLabel, bool caseSensitive,
            out IUIAutomationElement returnElement)
        {
            TreeScope scope = TreeScope.TreeScope_Children;
            if (searchDescendants)
            {
                scope = TreeScope.TreeScope_Descendants;
            }
            if (index == 0)
            {
                index = 1;
            }
            
            int nWaitMs = Engine.GetInstance().Timeout;
            IUIAutomationElementArray collection = null;
            List<IUIAutomationElement> foundElements = new List<IUIAutomationElement>();

            //while (nWaitMs > 0)
			while (true)
            {
                collection = this.uiElement.FindAll(scope, typeCondition);

                if ((collection != null) && (collection.Length >= index))
                {
                    List<IUIAutomationElement> elements = Helper.MatchStrings(collection, name,
                        bSearchByLabel, caseSensitive);
                        
                    if (elements != null)
                    {
                        foreach (IUIAutomationElement el in elements)
                        {
                            if (el.CurrentClassName.StartsWith(className))
                            {
                                foundElements.Add(el);
                            }
                        }
                    }

                    if ((foundElements != null) && (foundElements.Count >= index))
                    {
                        break;
                    }
                }
				
				/*if (Engine.IsCancelled == true)
				{
					Engine.IsCancelled = false;
					break;
				}*/
				
                nWaitMs -= ElementBase.waitPeriod;
				if (nWaitMs <= 0)
				{
					break;
				}
                Thread.Sleep(ElementBase.waitPeriod);
            }
            //Engine.TraceInLogFile("elements found: " + foundElements.Count);

            if ((foundElements == null) || (foundElements.Count == 0))
            {
                returnElement = null;
                return Errors.ElementNotFound;
            }

            if (index <= foundElements.Count)
            {
                returnElement = foundElements[index - 1];
                return Errors.None;
            }
            else
            {
                returnElement = null;
                return Errors.IndexTooBig;
            }
        }
		
		internal IUIAutomationElement FindFirst(int type, string name,
            bool searchDescendants, bool bSearchByLabel, bool caseSensitive)
        {
            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(type, name, 0, searchDescendants,
                bSearchByLabel, caseSensitive, out returnElement);

            return returnElement;
        }
        
        internal IUIAutomationElement FindFirstPlusCondition(int type, IUIAutomationCondition cond, 
            string name, bool searchDescendants, bool bSearchByLabel, bool caseSensitive)
        {
            IUIAutomationElement returnElement = null;

            Errors error = this.FindAtPlusCondition(type, cond, name, 0, searchDescendants,
                bSearchByLabel, caseSensitive, out returnElement);

            return returnElement;
        }
	}
}