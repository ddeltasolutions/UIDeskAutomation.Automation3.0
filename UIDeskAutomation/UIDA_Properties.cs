using System;
using System.Collections.Generic;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// UI Automation properties
    /// </summary>
    public enum UIDA_Property
    {
        /// <summary>
        /// Identifies the AcceleratorKey property, which is a string containing the accelerator key (also called shortcut key) combinations for the automation element.
        /// Shortcut key combinations invoke an action. For example, CTRL+O is often used to invoke the Open file common dialog box.
        /// </summary>
        AcceleratorKey,

        /// <summary>
        /// Identifies the AccessKey property, which is a string containing the access key character for the automation element.
        /// An access key (sometimes called a mnemonic) is a character in the text of a menu, menu item, or label of a control such as a button, that activates the associated menu function. For example, to open the File menu, for which the access key is typically F, the user would press ALT+F.
        /// </summary>
        AccessKey,

        /// <summary>
        /// Identifies the AriaRole property, which is a string containing the Accessible Rich Internet Application (ARIA) role information for the automation element.
        /// </summary>
        AriaRole,

        /// <summary>
        /// Identifies the AutomationId property, which is a string containing the UI Automation identifier (ID) for the automation element.
        /// </summary>
        AutomationId,

        /// <summary>
        /// Identifies the ClassName property, which is a string containing the class name for the automation element as assigned by the control developer.
        /// </summary>
        ClassName,

        /// <summary>
        /// Identifies the FrameworkId property, which is a string containing the name of the underlying UI framework that the automation element belongs to.
        /// The FrameworkId enables client applications to process automation elements differently depending on the particular UI framework. Examples of property values include "Win32", "WinForm", and "DirectUI".
        /// </summary>
        FrameworkId,

        /// <summary>
        /// The FullDescription property exposes a localized string which can contain extended description text for an element. FullDescription can contain a more complete description of an element than may be appropriate for the element Name.
        /// </summary>
        FullDescription,

        /// <summary>
        /// Identifies the HelpText property, which is a help text string associated with the automation element.
        /// </summary>
        HelpText,

        /// <summary>
        /// Identifies the ItemStatus property, which is a text string describing the status of an item of the automation element.
        /// ItemStatus enables a client to ascertain whether an element is conveying status about an item as well as what the status is. For example, an item associated with a contact in a messaging application might be "Busy" or "Connected".
        /// </summary>
        ItemStatus,

        /// <summary>
        /// Identifies the ItemType property, which is a text string describing the type of the automation element.
        /// ItemType is used to obtain information about items in a list, tree view, or data grid. For example, an item in a file directory view might be a "Document File" or a "Folder".
        /// </summary>
        ItemType,

        /// <summary>
        /// Identifies the LocalizedLandmarkType, which is a text string describing the type of landmark that the automation element represents.
        /// </summary>
        LocalizedLandmarkType,
        //ProcessId,

        /// <summary>
        /// Identifies the ProviderDescription property, which is a formatted string containing the source information of the UI Automation provider for the automation element, including proxy information.
        /// </summary>
        ProviderDescription
    }

    public partial class ElementBase
    {
        private static Dictionary<UIDA_Property, int> mapPropertyIds = null;

        internal static Dictionary<UIDA_Property, int> MapPropertyIds
        {
            get
            {
                if (mapPropertyIds != null)
                {
                    return mapPropertyIds;
                }

                mapPropertyIds = new Dictionary<UIDA_Property, int>();
                mapPropertyIds.Add(UIDA_Property.AcceleratorKey, UIA_PropertyIds.UIA_AcceleratorKeyPropertyId);
                mapPropertyIds.Add(UIDA_Property.AccessKey, UIA_PropertyIds.UIA_AccessKeyPropertyId);
                mapPropertyIds.Add(UIDA_Property.AriaRole, UIA_PropertyIds.UIA_AriaRolePropertyId);
                mapPropertyIds.Add(UIDA_Property.AutomationId, UIA_PropertyIds.UIA_AutomationIdPropertyId);
                mapPropertyIds.Add(UIDA_Property.ClassName, UIA_PropertyIds.UIA_ClassNamePropertyId);
                mapPropertyIds.Add(UIDA_Property.FrameworkId, UIA_PropertyIds.UIA_FrameworkIdPropertyId);
                mapPropertyIds.Add(UIDA_Property.FullDescription, UIA_PropertyIds.UIA_FullDescriptionPropertyId);
                mapPropertyIds.Add(UIDA_Property.HelpText, UIA_PropertyIds.UIA_HelpTextPropertyId);
                mapPropertyIds.Add(UIDA_Property.ItemStatus, UIA_PropertyIds.UIA_ItemStatusPropertyId);
                mapPropertyIds.Add(UIDA_Property.ItemType, UIA_PropertyIds.UIA_ItemTypePropertyId);
                mapPropertyIds.Add(UIDA_Property.LocalizedLandmarkType, UIA_PropertyIds.UIA_LocalizedLandmarkTypePropertyId);
                //mapPropertyIds.Add(UIDA_Property.ProcessId, UIA_PropertyIds.UIA_ProcessIdPropertyId);
                mapPropertyIds.Add(UIDA_Property.ProviderDescription, UIA_PropertyIds.UIA_ProviderDescriptionPropertyId);
                return mapPropertyIds;
            }
        }

        /// <summary>
        /// Finds the first child given the value of a property.
        /// </summary>
        /// <param name="property">The property</param>
        /// <param name="value">The value of the property</param>
        /// <param name="ignoreCase">true - search case insensitive, false - case sensitive</param>
        /// <param name="matchSubstring">true - the given value is a substring of the actual property value, false - perfect match</param>
        /// <returns>A UIDA_Generic object. You can call As...() functions to cast to whatever element type you want.</returns>
        public UIDA_Generic FindFirstChild(UIDA_Property property, string value, 
            bool ignoreCase = false, bool matchSubstring = false)
        {
            PropertyConditionFlags flags = PropertyConditionFlags.PropertyConditionFlags_None;
            if (ignoreCase == true)
            {
                flags = flags | PropertyConditionFlags.PropertyConditionFlags_IgnoreCase;
            }
            if (matchSubstring == true)
            {
                flags = flags | PropertyConditionFlags.PropertyConditionFlags_MatchSubstring;
            }

            IUIAutomationCondition condition = Engine.uiAutomation.CreatePropertyConditionEx(
                MapPropertyIds[property], value, flags);
            IUIAutomationElement elementFound = uiElement.FindFirst(UIAutomationClient.TreeScope.TreeScope_Children, condition);
            if (elementFound == null) 
            {
                return null;
            }
            return new UIDA_Generic(elementFound);
        }

        /// <summary>
        /// Finds the first descendant given the value of a property.
        /// </summary>
        /// <param name="property">The property</param>
        /// <param name="value">The value of the property</param>
        /// <param name="ignoreCase">true - search case insensitive, false - case sensitive</param>
        /// <param name="matchSubstring">true - the given value is a substring of the actual property value, false - perfect match</param>
        /// <returns>A UIDA_Generic object. You can call As...() functions to cast to whatever element type you want.</returns>
        public UIDA_Generic FindFirstDescendant(UIDA_Property property, string value, 
            bool ignoreCase = false, bool matchSubstring = false)
        {
            PropertyConditionFlags flags = PropertyConditionFlags.PropertyConditionFlags_None;
            if (ignoreCase == true)
            {
                flags = flags | PropertyConditionFlags.PropertyConditionFlags_IgnoreCase;
            }
            if (matchSubstring == true)
            {
                flags = flags | PropertyConditionFlags.PropertyConditionFlags_MatchSubstring;
            }

            IUIAutomationCondition condition = Engine.uiAutomation.CreatePropertyConditionEx(
                MapPropertyIds[property], value, flags);
            IUIAutomationElement elementFound = uiElement.FindFirst(UIAutomationClient.TreeScope.TreeScope_Descendants, condition);
            if (elementFound == null)
            {
                return null;
            }
            return new UIDA_Generic(elementFound);
        }

        /// <summary>
        /// Finds all children given the value of a property.
        /// </summary>
        /// <param name="property">The property</param>
        /// <param name="value">The value of the property</param>
        /// <param name="ignoreCase">true - search case insensitive, false - case sensitive</param>
        /// <param name="matchSubstring">true - the given value is a substring of the actual property value, false - perfect match</param>
        /// <returns>A UIDA_Generic objects array. You can call As...() functions for any object in the array to cast to whatever element type you want.</returns>
        public UIDA_Generic[] FindAllChildren(UIDA_Property property, string value,
            bool ignoreCase = false, bool matchSubstring = false)
        {
            PropertyConditionFlags flags = PropertyConditionFlags.PropertyConditionFlags_None;
            if (ignoreCase == true)
            {
                flags = flags | PropertyConditionFlags.PropertyConditionFlags_IgnoreCase;
            }
            if (matchSubstring == true)
            {
                flags = flags | PropertyConditionFlags.PropertyConditionFlags_MatchSubstring;
            }

            IUIAutomationCondition condition = Engine.uiAutomation.CreatePropertyConditionEx(
                MapPropertyIds[property], value, flags);
            IUIAutomationElementArray elementsArray = uiElement.FindAll(UIAutomationClient.TreeScope.TreeScope_Children, condition);
            if (elementsArray == null)
            {
                return null;
            }

            List<UIDA_Generic> children = new List<UIDA_Generic>();
            int length = elementsArray.Length;
            for (int i = 0; i < length; i++)
            {
                children.Add(new UIDA_Generic(elementsArray.GetElement(i)));
            }
            return children.ToArray();
        }

        /// <summary>
        /// Finds all descendants given the value of a property.
        /// </summary>
        /// <param name="property">The property</param>
        /// <param name="value">The value of the property</param>
        /// <param name="ignoreCase">true - search case insensitive, false - case sensitive</param>
        /// <param name="matchSubstring">true - the given value is a substring of the actual property value, false - perfect match</param>
        /// <returns>A UIDA_Generic objects array. You can call As...() functions for any object in the array to cast to whatever element type you want.</returns>
        public UIDA_Generic[] FindAllDescendants(UIDA_Property property, string value,
            bool ignoreCase = false, bool matchSubstring = false)
        {
            PropertyConditionFlags flags = PropertyConditionFlags.PropertyConditionFlags_None;
            if (ignoreCase == true)
            {
                flags = flags | PropertyConditionFlags.PropertyConditionFlags_IgnoreCase;
            }
            if (matchSubstring == true)
            {
                flags = flags | PropertyConditionFlags.PropertyConditionFlags_MatchSubstring;
            }

            IUIAutomationCondition condition = Engine.uiAutomation.CreatePropertyConditionEx(
                MapPropertyIds[property], value, flags);
            IUIAutomationElementArray elementsArray = uiElement.FindAll(UIAutomationClient.TreeScope.TreeScope_Descendants, condition);
            if (elementsArray == null)
            {
                return null;
            }

            List<UIDA_Generic> children = new List<UIDA_Generic>();
            int length = elementsArray.Length;
            for (int i = 0; i < length; i++)
            {
                children.Add(new UIDA_Generic(elementsArray.GetElement(i)));
            }
            return children.ToArray();
        }
    }
}
