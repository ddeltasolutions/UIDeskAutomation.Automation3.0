using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using UIAutomationClient;

namespace UIDeskAutomationLib
{
    partial class ElementBase
    {
        #region Search functions
		
        /// <summary>
        /// Searches for a title bar among the children of the element or its descendants
        /// </summary>
        /// <param name="name">title bar name, wildcards can be used</param>
        /// <param name="index">title bar index, starts with 1</param>
        /// <param name="searchDescendants">search descendants, default false</param>
        /// <param name="caseSensitive">search name with case sensitive criteria</param>
        /// <returns>UIDA_TitleBar, null if not found</returns>
        public UIDA_TitleBar TitleBarAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("TitleBarAt method - index cannot be less than zero");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("TitleBarAt method - index cannot be less than zero");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_TitleBarControlTypeId, name, index,
                searchDescendants, false, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("TitleBarAt method - titlebar element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("TitleBarAt method - titlebar element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("TitleBarAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("TitleBarAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_TitleBar titleBar = new UIDA_TitleBar(returnElement);
            return titleBar;
        }

        /// <summary>
        /// Searches for a title bar among the children of the element or its descendants
        /// </summary>
        /// <param name="name">title bar name, wildcards can be used</param>
        /// <param name="searchDescendants">search descendants, default false</param>
        /// <param name="caseSensitive">search name with case sensitive criteria</param>
        /// <returns>first UIDA_TitleBar that matches the search criteria, null if not found</returns>
        public UIDA_TitleBar TitleBar(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_TitleBarControlTypeId, 
                name, searchDescendants, false, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("TitleBar method - titlebar element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("TitleBar method - titlebar element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_TitleBar titleBar = new UIDA_TitleBar(returnElement);
            return titleBar;
        }

        /// <summary>
        /// Returns a collection of TitleBars that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of TitleBar elements, use null to return all TitleBars</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>all UIDA_TitleBar elements that match the search criteria</returns>
        public UIDA_TitleBar[] TitleBars(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allTitleBars = FindAll(UIA_ControlTypeIds.UIA_TitleBarControlTypeId,
                name, searchDescendants, false, caseSensitive);

            List<UIDA_TitleBar> titlebars = new List<UIDA_TitleBar>();
            if (allTitleBars != null)
            {
                foreach (IUIAutomationElement crtEl in allTitleBars)
                {
                    titlebars.Add(new UIDA_TitleBar(crtEl));
                }
            }
            return titlebars.ToArray();
        }

        /// <summary>
        /// Searches for a menu bar among the children of the element or its descendants
        /// </summary>
        /// <param name="name">menu bar name, wildcards can be used</param>
        /// <param name="index">menu bar index, starts with 1</param>
        /// <param name="searchDescendants">search descendants, default false</param>
        /// <param name="caseSensitive">search name with case sensitive criteria</param>
        /// <returns>UIDA_MenuBar, null if not found</returns>
        public UIDA_MenuBar MenuBarAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("MenuBarAt method - index cannot be less than zero");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("MenuBarAt method - index cannot be less than zero");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_MenuBarControlTypeId, name, index, searchDescendants,
                false, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("MenuBarAt method - menubar element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("MenuBarAt method - menubar element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("MenuBarAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("MenuBarAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_MenuBar menuBar = new UIDA_MenuBar(returnElement);
            return menuBar;
        }

        /// <summary>
        /// Searches for a menu bar among the children of the element or its descendants
        /// </summary>
        /// <param name="name">menu bar name, wildcards can be used</param>
        /// <param name="searchDescendants">search descendants, default false</param>
        /// <param name="caseSensitive">search name with case sensitive criteria</param>
        /// <returns>first UIDA_MenuBar that matches the search criteria, null if not found</returns>
        public UIDA_MenuBar MenuBar(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_MenuBarControlTypeId, name,
                searchDescendants, false, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("MenuBar method - menubar element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("MenuBar method - menubar element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_MenuBar menuBar = new UIDA_MenuBar(returnElement);
            return menuBar;
        }

        /// <summary>
        /// Returns a collection of MenuBars that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of MenuBar elements, use null to return all MenuBars</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>all UIDA_MenuBar elements that match the search criteria</returns>
        public UIDA_MenuBar[] MenuBars(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allMenuBars = FindAll(UIA_ControlTypeIds.UIA_MenuBarControlTypeId,
                name, searchDescendants, false, caseSensitive);

            List<UIDA_MenuBar> menubars = new List<UIDA_MenuBar>();
            if (allMenuBars != null)
            {
                foreach (IUIAutomationElement crtEl in allMenuBars)
                {
                    menubars.Add(new UIDA_MenuBar(crtEl));
                }
            }
            return menubars.ToArray();
        }

        /// <summary>
        /// Searches for a menu item among the children of the element or its descendants
        /// </summary>
        /// <param name="name">menu item text, wildcards can be used</param>
        /// <param name="index">menu item index, starts with 1</param>
        /// <param name="searchDescendants">search descendants, default false</param>
        /// <param name="caseSensitive">search name with case sensitive criteria</param>
        /// <returns>UIDA_MenuItem, null if not found</returns>
        public UIDA_MenuItem MenuItemAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("MenuItemAt method - index cannot be less than zero");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("MenuItemAt method - index cannot be less than zero");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_MenuItemControlTypeId, name, index, searchDescendants,
                false, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("MenuItemAt method - menuitem element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("MenuItemAt method - menuitem element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("MenuItemAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("MenuItemAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_MenuItem menuItem = new UIDA_MenuItem(returnElement);
            return menuItem;
        }

        /// <summary>
        /// Searches for a menu item among the children of the element or its descendants
        /// </summary>
        /// <param name="name">menu item text, wildcards can be used</param>
        /// <param name="searchDescendants">search descendants, default false</param>
        /// <param name="caseSensitive">search name with case sensitive criteria</param>
        /// <returns>first UIDA_MenuItem that matches the search criteria, null if not found</returns>
        public UIDA_MenuItem MenuItem(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_MenuItemControlTypeId, name,
                searchDescendants, false, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("MenuItem method - menuitem element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("MenuItem method - menuitem element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_MenuItem menuItem = new UIDA_MenuItem(returnElement);
            return menuItem;
        }

        /// <summary>
        /// Returns a collection of MenuItems that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of MenuItem elements, use null to return all MenuItems</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>all UIDA_MenuItem elements that match the search criteria</returns>
        public UIDA_MenuItem[] MenuItems(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allMenuItems = FindAll(UIA_ControlTypeIds.UIA_MenuItemControlTypeId,
                name, searchDescendants, false, caseSensitive);

            List<UIDA_MenuItem> menuitems = new List<UIDA_MenuItem>();
            if (allMenuItems != null)
            {
                foreach (IUIAutomationElement crtEl in allMenuItems)
                {
                    menuitems.Add(new UIDA_MenuItem(crtEl));
                }
            }
            return menuitems.ToArray();
        }

        /// <summary>
        /// Searches for a top-level menu among the children of the element or its descendants
        /// </summary>
        /// <param name="name">menu name, wildcards can be used</param>
        /// <param name="index">menu index, starts with 1</param>
        /// <param name="searchDescendants">search descendants, default false</param>
        /// <param name="caseSensitive">search name with case sensitive criteria</param>
        /// <returns>UIDA_TopLevelMenu, null if not found</returns>
        public UIDA_TopLevelMenu MenuAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("MenuAt method - index cannot be less than zero");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("MenuAt method - index cannot be less than zero");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_MenuControlTypeId, name, index, searchDescendants,
                false, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("MenuAt method - menu element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("MenuAt method - menu element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("MenuAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("MenuAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_TopLevelMenu menu = new UIDA_TopLevelMenu(returnElement);
            return menu;
        }

        /// <summary>
        /// Searches for a top-level menu among the children of the element or its descendants
        /// </summary>
        /// <param name="name">menu text, wildcards can be used</param>
        /// <param name="searchDescendants">search descendants, default false</param>
        /// <param name="caseSensitive">search name with case sensitive criteria</param>
        /// <returns>first UIDA_TopLevelMenu that matches the search criteria, null if not found</returns>
        public UIDA_TopLevelMenu Menu(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_MenuControlTypeId, name,
                searchDescendants, false, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("Menu method - menu element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("Menu method - menu element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_TopLevelMenu menu = new UIDA_TopLevelMenu(returnElement);
            return menu;
        }

        /// <summary>
        /// Returns a collection of TopLevelMenus that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of TopLevelMenu elements, use null to return all Menus</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>all UIDA_TopLevelMenu elements that match the search criteria</returns>
        public UIDA_TopLevelMenu[] Menus(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allMenus = FindAll(UIA_ControlTypeIds.UIA_MenuControlTypeId,
                name, searchDescendants, false, caseSensitive);

            List<UIDA_TopLevelMenu> menus = new List<UIDA_TopLevelMenu>();
            if (allMenus != null)
            {
                foreach (IUIAutomationElement crtEl in allMenus)
                {
                    menus.Add(new UIDA_TopLevelMenu(crtEl));
                }
            }
            return menus.ToArray();
        }

        /// <summary>
        /// Searches for a button among the children of the element or its descendants
        /// </summary>
        /// <param name="name">button name, wildcards can be used</param>
        /// <param name="index">button index, starts with 1</param>
        /// <param name="searchDescendants">search descendants, default false</param>
        /// <param name="caseSensitive">search name with case sensitive criteria</param>
        /// <returns>UIDA_Button, null if not found</returns>
        public UIDA_Button ButtonAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("ButtonAt method - index cannot be less than zero");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ButtonAt method - index cannot be less than zero");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_ButtonControlTypeId, name, index, searchDescendants,
                false, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("ButtonAt method - button element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ButtonAt method - button element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("ButtonAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ButtonAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_Button button = new UIDA_Button(returnElement);
            return button;
        }

        /// <summary>
        /// Searches for a button among the children of the element or its descendants
        /// </summary>
        /// <param name="name">button name, wildcards can be used</param>
        /// <param name="searchDescendants">search descendants, default false</param>
        /// <param name="caseSensitive">search name with case sensitive criteria</param>
        /// <returns>first UIDA_Button that matches the search criteria, null if not found</returns>
        public UIDA_Button Button(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_ButtonControlTypeId, name,
                searchDescendants, false, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("Button method - button element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("Button method - button element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_Button button = new UIDA_Button(returnElement);
            return button;
        }

        /// <summary>
        /// Returns a collection of Buttons that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of Button elements, use null to return all Buttons</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>all UIDA_Button elements that match the search criteria</returns>
        public UIDA_Button[] Buttons(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allButtons = FindAll(UIA_ControlTypeIds.UIA_ButtonControlTypeId,
                name, searchDescendants, false, caseSensitive);

            List<UIDA_Button> buttons = new List<UIDA_Button>();
            if (allButtons != null)
            {
                foreach (IUIAutomationElement crtEl in allButtons)
                {
                    buttons.Add(new UIDA_Button(crtEl));
                }
            }
            return buttons.ToArray();
        }

        /// <summary>
        /// Searches for a calendar among the children of the element or its descendants
        /// </summary>
        /// <param name="name">calendar name, wildcards can be used</param>
        /// <param name="index">calendar index, starts with 1</param>
        /// <param name="searchDescendants">search descendants, default false</param>
        /// <param name="caseSensitive">search name with case sensitive criteria</param>
        /// <returns>UIDA_Calendar, null if not found</returns>
        public UIDA_Calendar CalendarAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("CalendarAt method - index cannot be less than zero");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("CalendarAt method - index cannot be less than zero");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;
            
            IUIAutomationCondition typeCond = Engine.uiAutomation.CreatePropertyCondition(
				UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_PaneControlTypeId);
            IUIAutomationCondition classCond = Engine.uiAutomation.CreatePropertyCondition(
				UIA_PropertyIds.UIA_ClassNamePropertyId, "SysMonthCal32");
            IUIAutomationCondition andCond = Engine.uiAutomation.CreateAndCondition(typeCond, classCond);
            Errors error = this.FindAtPlusCondition(UIA_ControlTypeIds.UIA_CalendarControlTypeId, andCond, name, index, 
                searchDescendants, false, caseSensitive, out returnElement);
            
            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("CalendarAt method - calendar element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("CalendarAt method - calendar element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("CalendarAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("CalendarAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_Calendar calendar = new UIDA_Calendar(returnElement);
            return calendar;
        }

        /// <summary>
        /// Searches for a calendar among the children of the element or its descendants
        /// </summary>
        /// <param name="name">calendar name, wildcards can be used</param>
        /// <param name="searchDescendants">search descendants, default false</param>
        /// <param name="caseSensitive">search name with case sensitive criteria</param>
        /// <returns>first UIDA_Calendar that matches the search criteria, null if not found</returns>
        public UIDA_Calendar Calendar(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationCondition typeCond = Engine.uiAutomation.CreatePropertyCondition(
				UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_PaneControlTypeId);
            IUIAutomationCondition classCond = Engine.uiAutomation.CreatePropertyCondition(
				UIA_PropertyIds.UIA_ClassNamePropertyId, "SysMonthCal32");
            IUIAutomationCondition andCond = Engine.uiAutomation.CreateAndCondition(typeCond, classCond);
            IUIAutomationElement returnElement = this.FindFirstPlusCondition(UIA_ControlTypeIds.UIA_CalendarControlTypeId, 
                andCond, name, searchDescendants, false, caseSensitive);
            
            if (returnElement == null)
            {
                Engine.TraceInLogFile("Calendar method - calendar element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("Calendar method - calendar element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_Calendar calendar = new UIDA_Calendar(returnElement);
            return calendar;
        }

        /// <summary>
        /// Returns a collection of Calendars that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of Calendar elements, use null to return all Calendars</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>all UIDA_Calendar elements that match the search criteria</returns>
        public UIDA_Calendar[] Calendars(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationCondition typeCond = Engine.uiAutomation.CreatePropertyCondition(
				UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_PaneControlTypeId);
            IUIAutomationCondition classCond = Engine.uiAutomation.CreatePropertyCondition(
				UIA_PropertyIds.UIA_ClassNamePropertyId, "SysMonthCal32");
            IUIAutomationCondition andCond = Engine.uiAutomation.CreateAndCondition(typeCond, classCond);
            
            List<IUIAutomationElement> allCalendars = FindAllPlusCondition(UIA_ControlTypeIds.UIA_CalendarControlTypeId,
                andCond, name, searchDescendants, false, caseSensitive);

            List<UIDA_Calendar> calendars = new List<UIDA_Calendar>();
            if (allCalendars != null)
            {
                foreach (IUIAutomationElement crtEl in allCalendars)
                {
                    calendars.Add(new UIDA_Calendar(crtEl));
                }
            }
            return calendars.ToArray();
        }

        /// <summary>
        /// Searches for a checkbox among the children of the element or its descendants
        /// </summary>
        /// <param name="name">checkbox name, wildcards can be used</param>
        /// <param name="index">checkbox index, starts with 1</param>
        /// <param name="searchDescendants">search descendants, default false</param>
        /// <param name="caseSensitive">search name with case sensitive criteria</param>
        /// <returns>UIDA_CheckBox, null if not found</returns>
        public UIDA_CheckBox CheckBoxAt(string name, int index,
            bool searchDescendants = false, bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("CheckboxAt method - index cannot be less than zero");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("CheckboxAt method - index cannot be less than zero");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_CheckBoxControlTypeId, name, index, searchDescendants,
                false, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("CheckboxAt method - checkbox element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("CheckboxAt method - checkbox element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("CheckboxAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("CheckboxAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_CheckBox checkbox = new UIDA_CheckBox(returnElement);
            return checkbox;
        }

        /// <summary>
        /// Searches for a checkbox among the children of the element or its descendants
        /// </summary>
        /// <param name="name">checkbox name, wildcards can be used</param>
        /// <param name="searchDescendants">search descendants, default false</param>
        /// <param name="caseSensitive">search name with case sensitive criteria</param>
        /// <returns>first UIDA_CheckBox that matches the search criteria, null if not found</returns>
        public UIDA_CheckBox CheckBox(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_CheckBoxControlTypeId, name,
                searchDescendants, false, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("Checkbox method - checkbox element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("Checkbox method - checkbox element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_CheckBox checkbox = new UIDA_CheckBox(returnElement);
            return checkbox;
        }

        /// <summary>
        /// Returns a collection of CheckBoxes that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of CheckBox elements, use null to return all CheckBoxes</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>all UIDA_CheckBox elements that match the search criteria</returns>
        public UIDA_CheckBox[] CheckBoxes(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allCheckBoxes = FindAll(UIA_ControlTypeIds.UIA_CheckBoxControlTypeId,
                name, searchDescendants, false, caseSensitive);

            List<UIDA_CheckBox> checkboxes = new List<UIDA_CheckBox>();
            if (allCheckBoxes != null)
            {
                foreach (IUIAutomationElement crtEl in allCheckBoxes)
                {
                    checkboxes.Add(new UIDA_CheckBox(crtEl));
                }
            }
            return checkboxes.ToArray();
        }

        /// <summary>
        /// Searches for a edit control among the children of the element or its descendants
        /// </summary>
        /// <param name="name">edit control name, wildcards can be used</param>
        /// <param name="index">edit control index, starts with 1</param>
        /// <param name="searchDescendants">search descendants, default false</param>
        /// <param name="caseSensitive">search name with case sensitive criteria</param>
        /// <returns>UIDA_Edit, null if not found</returns>
        public UIDA_Edit EditAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("EditAt method - index cannot be less than zero");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("EditAt method - index cannot be less than zero");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_EditControlTypeId, name, index, searchDescendants,
                true, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("EditAt method - edit element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("EditAt method - edit element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("EditAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("EditAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_Edit edit = new UIDA_Edit(returnElement);
            return edit;
        }
        /// <summary>
        /// Searches for a edit control among the children of the element or its descendants
        /// </summary>
        /// <param name="name">edit control name, wildcards can be used</param>
        /// <param name="searchDescendants">search descendants, default false</param>
        /// <param name="caseSensitive">search name with case sensitive criteria</param>
        /// <returns>UIDA_Edit, null if not found</returns>
        public UIDA_Edit Edit(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_EditControlTypeId, name,
                searchDescendants, true, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("Edit method - edit element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("Edit method - edit element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_Edit edit = new UIDA_Edit(returnElement);
            return edit;
        }

        /// <summary>
        /// Returns a collection of Edits that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of Edit elements, use null to return all Edits</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_Edit elements</returns>
        public UIDA_Edit[] Edits(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allEdits = FindAll(UIA_ControlTypeIds.UIA_EditControlTypeId,
                name, searchDescendants, true, caseSensitive);

            List<UIDA_Edit> edits = new List<UIDA_Edit>();
            if (allEdits != null)
            {
                foreach (IUIAutomationElement crtEl in allEdits)
                {
                    edits.Add(new UIDA_Edit(crtEl));
                }
            }
            return edits.ToArray();
        }

        /// <summary>
        /// Searches for a document control among the children of the element or its descendants
        /// </summary>
        /// <param name="name">document control name, wildcards can be used</param>
        /// <param name="index">document control index, starts with 1</param>
        /// <param name="searchDescendants">search descendants, default false</param>
        /// <param name="caseSensitive">search name with case sensitive criteria</param>
        /// <returns>UIDA_Document, null if not found</returns>
        public UIDA_Document DocumentAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("DocumentAt method - index cannot be less than zero");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("DocumentAt method - index cannot be less than zero");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_DocumentControlTypeId, name, index, searchDescendants,
                true, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("DocumentAt method - document element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("DocumentAt method - document element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("DocumentAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("DocumentAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_Document document = new UIDA_Document(returnElement);
            return document;
        }

        /// <summary>
        /// Searches for a document among the children of the element or its descendants
        /// </summary>
        /// <param name="name">document name, wildcards can be used</param>
        /// <param name="searchDescendants">search descendants, default false</param>
        /// <param name="caseSensitive">search name with case sensitive criteria</param>
        /// <returns>first UIDA_Document matching the search criteria, null if not found</returns>
        public UIDA_Document Document(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_DocumentControlTypeId,
                name, searchDescendants, true, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("Document method - document element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("Document method - document element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_Document document = new UIDA_Document(returnElement);
            return document;
        }

        /// <summary>
        /// Returns a collection of Documents that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of Document elements, use null to return all Documents</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>all UIDA_Document elements matching the search criteria</returns>
        public UIDA_Document[] Documents(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allDocuments = FindAll(UIA_ControlTypeIds.UIA_DocumentControlTypeId,
                name, searchDescendants, true, caseSensitive);

            List<UIDA_Document> documents = new List<UIDA_Document>();
            if (allDocuments != null)
            {
                foreach (IUIAutomationElement crtEl in allDocuments)
                {
                    documents.Add(new UIDA_Document(crtEl));
                }
            }
            return documents.ToArray();
        }

        /// <summary>
        /// Returns a pane element at a given index
        /// </summary>
        /// <param name="name">name of pane element, wildcards can be used</param>
        /// <param name="index">index of pane, starts with 1</param>
        /// <param name="searchDescendants">true is search deep through descendants, false otherwise</param>
        /// <param name="caseSensitive">true if search of name is case sensitive, default true</param>
        /// <returns>UIDA_Pane element, null if not found</returns>
        public UIDA_Pane PaneAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("PaneAt method - index cannot be negative");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("PaneAt method - index cannot be negative");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_PaneControlTypeId, name, index, searchDescendants,
                false, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("PaneAt method - pane element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("PaneAt method - pane element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("PaneAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("PaneAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_Pane pane = new UIDA_Pane(returnElement);
            return pane;
        }

        /// <summary>
        /// Searches for a Pane element among the children of the element or its descendants.
        /// </summary>
        /// <param name="name">name of pane, wildcards can be used</param>
        /// <param name="searchDescendants">search deep through descendants</param>
        /// <param name="caseSensitive">specifies if name is searched case sensitive, default true</param>
        /// <returns>UIDA_Pane element, null if not found</returns>
        public UIDA_Pane Pane(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_PaneControlTypeId,
                name, searchDescendants, false, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("Pane method - pane element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("Pane method - pane element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_Pane pane = new UIDA_Pane(returnElement);
            return pane;
        }

        /// <summary>
        /// Returns a collection of Panes that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of Pane elements, use null to return all Panes</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_Pane elements</returns>
        public UIDA_Pane[] Panes(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allPanes = FindAll(UIA_ControlTypeIds.UIA_PaneControlTypeId,
                name, searchDescendants, false, caseSensitive);

            List<UIDA_Pane> panes = new List<UIDA_Pane>();
            if (allPanes != null)
            {
                foreach (IUIAutomationElement crtEl in allPanes)
                {
                    panes.Add(new UIDA_Pane(crtEl));
                }
            }
            return panes.ToArray();
        }

        /// <summary>
        /// Returns a DatePicker element with given name and index
        /// </summary>
        /// <param name="name">name to search, wildcards can be used</param>
        /// <param name="index">index of DatePicker element, starts with 1</param>
        /// <param name="searchDescendants">true if search through all descendants, default false</param>
        /// <param name="caseSensitive">tells is name is searched case sensitive, default true</param>
        /// <returns>UIDA_DatePicker element, null if not found</returns>
        public UIDA_DatePicker DatePickerAt(string name, int index,
            bool searchDescendants = false, bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("DatePickerAt method - index cannot be negative");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("DatePickerAt method - index cannot be negative");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            //// looking for the element
            Errors error = Errors.ElementNotFound;
            string fid = this.uiElement.CurrentFrameworkId;
            if (fid == "WPF")
            {
                error = this.FindCustomAt(UIA_ControlTypeIds.UIA_CustomControlTypeId, "DatePicker", name, index,
                    searchDescendants, true, caseSensitive, out returnElement);
            }
            else if (fid == "Win32")
            {
                error = this.FindCustomAt(UIA_ControlTypeIds.UIA_PaneControlTypeId, "SysDateTimePick32", name, index,
                    searchDescendants, true, caseSensitive, out returnElement);
            }
            else if (fid == "WinForm")
            {
                error = this.FindCustomAt(UIA_ControlTypeIds.UIA_PaneControlTypeId, "WindowsForms10.SysDateTimePick32", name, index,
                    searchDescendants, true, caseSensitive, out returnElement);
            }
            ////////////

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("DatePickerAt method - element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("DatePickerAt method - element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("DatePickerAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("DatePickerAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_DatePicker datePicker = new UIDA_DatePicker(returnElement);
            return datePicker;
        }

        /// <summary>
        /// Searches for a DatePicker element
        /// </summary>
        /// <param name="name">name of DatePicker element, wildcards can be used</param>
        /// <param name="searchDescendants">true if search through descendants, false if search through direct children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_DatePicker, null if not found</returns>
        public UIDA_DatePicker DatePicker(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = null;

            Errors error = Errors.ElementNotFound;
            string fid = this.uiElement.CurrentFrameworkId;
            if (fid == "WPF")
            {
                error = this.FindCustomAt(UIA_ControlTypeIds.UIA_CustomControlTypeId, "DatePicker", name, 0,
                    searchDescendants, true, caseSensitive, out returnElement);
            }
            else if (fid == "Win32")
            {
                error = this.FindCustomAt(UIA_ControlTypeIds.UIA_PaneControlTypeId, "SysDateTimePick32", name, 0,
                    searchDescendants, true, caseSensitive, out returnElement);
            }
            else if (fid == "WinForm")
            {
                error = this.FindCustomAt(UIA_ControlTypeIds.UIA_PaneControlTypeId, "WindowsForms10.SysDateTimePick32", name, 0,
                    searchDescendants, true, caseSensitive, out returnElement);
            }

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("DatePicker method - element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("DatePicker method - element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_DatePicker datePicker = new UIDA_DatePicker(returnElement);
            return datePicker;
        }

        /// <summary>
        /// Returns a collection of DatePickers that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of DatePicker elements, use null to return all DatePickers</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_DatePicker elements</returns>
        public UIDA_DatePicker[] DatePickers(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allDatePickers = null;
            string fid = this.uiElement.CurrentFrameworkId;
            if (fid == "WPF")
            {
                allDatePickers = FindAllCustom(UIA_ControlTypeIds.UIA_CustomControlTypeId, 
                    "DatePicker", name, searchDescendants, true, caseSensitive);
            }
            else if (fid == "Win32")
            {
                allDatePickers = FindAllCustom(UIA_ControlTypeIds.UIA_PaneControlTypeId, 
                    "SysDateTimePick32", name, searchDescendants, true, caseSensitive);
            }
            else if (fid == "WinForm")
            {
                allDatePickers = FindAllCustom(UIA_ControlTypeIds.UIA_PaneControlTypeId, 
                    "WindowsForms10.SysDateTimePick32.app.0.378734a", name, searchDescendants, true, caseSensitive);
            }

            List<UIDA_DatePicker> datepickers = new List<UIDA_DatePicker>();
            if (allDatePickers != null)
            {
                foreach (IUIAutomationElement crtEl in allDatePickers)
                {
                    datepickers.Add(new UIDA_DatePicker(crtEl));
                }
            }
            return datepickers.ToArray();
        }

        /// <summary>
        /// Searches a list in the current element
        /// </summary>
        /// <param name="name">name of list, wildcards can be used</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is case sensitive, default true</param>
        /// <returns>UIDA_List element, null if not found</returns>
        public UIDA_List List(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_ListControlTypeId,
                name, searchDescendants, true, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("List method - list element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("List method - list element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_List list = new UIDA_List(returnElement);
            return list;
        }

        /// <summary>
        /// Searches for a list with a given name at a given index.
        /// </summary>
        /// <param name="name">name, if it is null then it will be ignored. Wildcards can be used.</param>
        /// <param name="index">index, starts with 1</param>
        /// <param name="searchDescendants">specifies if search is done through descendants or through children, default is false which means search is done through children</param>
        /// <param name="caseSensitive">specifies is name search is done case sensitive or not, default true</param>
        /// <returns>UIDA_List element, null if not found</returns>
        public UIDA_List ListAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("ListAt method - index cannot be negative");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ListAt method - index cannot be negative");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_ListControlTypeId, name, index, searchDescendants,
                true, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("ListAt method - list element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ListAt method - list element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("ListAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ListAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_List list = new UIDA_List(returnElement);
            return list;
        }

        /// <summary>
        /// Returns a collection of Lists that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of List elements, use null to return all Lists.</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_List elements</returns>
        public UIDA_List[] Lists(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allLists = FindAll(UIA_ControlTypeIds.UIA_ListControlTypeId,
                name, searchDescendants, true, caseSensitive);

            List<UIDA_List> lists = new List<UIDA_List>();
            if (allLists != null)
            {
                foreach (IUIAutomationElement crtEl in allLists)
                {
                    lists.Add(new UIDA_List(crtEl));
                }
            }
            return lists.ToArray();
        }

        /// <summary>
        /// Searches a ComboBox in the current element
        /// </summary>
        /// <param name="name">name of ComboBox, wildcards can be used</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is case sensitive, default true</param>
        /// <returns>UIDA_ComboBox element, null if not found</returns>
        public UIDA_ComboBox ComboBox(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_ComboBoxControlTypeId,
                name, searchDescendants, true, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("ComboBox method - ComboBox element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ComboBox method - ComboBox element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_ComboBox comboBox = new UIDA_ComboBox(returnElement);
            return comboBox;
        }

        /// <summary>
        /// Searches for a ComboBox with a specified name at a specified index.
        /// </summary>
        /// <param name="name">name or label of ComboBox, wildcards can be used</param>
        /// <param name="index">index of ComboBox, starts with 1</param>
        /// <param name="searchDescendants">true if search through descendants, false if search only through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_ComboBox element, null if not found</returns>
        public UIDA_ComboBox ComboBoxAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("ComboBoxAt method - index cannot be negative");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ComboBoxAt method - index cannot be negative");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_ComboBoxControlTypeId, name, index, searchDescendants,
                true, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("ComboBoxAt method - ComboBox element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ComboBoxAt method - ComboBox element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("ComboBoxAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ComboBoxAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_ComboBox combo = new UIDA_ComboBox(returnElement);
            return combo;
        }

        /// <summary>
        /// Returns a collection of ComboBoxes that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of ComboBox elements, use null to return all ComboBoxes</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_ComboBox elements</returns>
        public UIDA_ComboBox[] ComboBoxes(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allCombos = FindAll(UIA_ControlTypeIds.UIA_ComboBoxControlTypeId,
                name, searchDescendants, true, caseSensitive);

            List<UIDA_ComboBox> combos = new List<UIDA_ComboBox>();
            if (allCombos != null)
            {
                foreach (IUIAutomationElement crtEl in allCombos)
                {
                    combos.Add(new UIDA_ComboBox(crtEl));
                }
            }
            return combos.ToArray();
        }

        /// <summary>
        /// Searches for a RadioButton with a specified name at a specified index.
        /// </summary>
        /// <param name="name">name of RadioButton, wildcards can be used</param>
        /// <param name="index">index of RadioButton, starts with 1</param>
        /// <param name="searchDescendants">true if search through descendants, false if search only through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_RadioButton element, null if not found</returns>
        public UIDA_RadioButton RadioButtonAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("RadioButtonAt method - index cannot be negative");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("RadioButtonAt method - index cannot be negative");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_RadioButtonControlTypeId, name, index, searchDescendants,
                true, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("RadioButtonAt method - RadioButton element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("RadioButtonAt method - RadioButton element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("RadioButtonAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("RadioButtonAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_RadioButton radio = new UIDA_RadioButton(returnElement);
            return radio;
        }

        /// <summary>
        /// Searches a RadioButton in the current element
        /// </summary>
        /// <param name="name">name of RadioButton, wildcards can be used</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is case sensitive, default true</param>
        /// <returns>UIDA_RadioButton element, null if not found</returns>
        public UIDA_RadioButton RadioButton(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_RadioButtonControlTypeId,
                name, searchDescendants, true, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("RadioButton method - RadioButton element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("RadioButton method - RadioButton element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_RadioButton radio = new UIDA_RadioButton(returnElement);
            return radio;
        }

        /// <summary>
        /// Returns a collection of RadioButtons that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of RadioButton elements, use null to return all RadioButtons</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_RadioButton elements</returns>
        public UIDA_RadioButton[] RadioButtons(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allRadios = FindAll(UIA_ControlTypeIds.UIA_RadioButtonControlTypeId,
                name, searchDescendants, true, caseSensitive);

            List<UIDA_RadioButton> radios = new List<UIDA_RadioButton>();
            if (allRadios != null)
            {
                foreach (IUIAutomationElement crtEl in allRadios)
                {
                    radios.Add(new UIDA_RadioButton(crtEl));
                }
            }
            return radios.ToArray();
        }

        /// <summary>
        /// Searches a Label in the current element
        /// </summary>
        /// <param name="name">text of Label, wildcards can be used</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is case sensitive, default true</param>
        /// <returns>UIDA_Label element, null if not found</returns>
        public UIDA_Label Label(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_TextControlTypeId,
                name, searchDescendants, false, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("Label method - Label element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("Label method - Label element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_Label label = new UIDA_Label(returnElement);
            return label;
        }

        /// <summary>
        /// Searches for a Static Text (Label) with a specified text at a specified index.
        /// </summary>
        /// <param name="name">text of Label, wildcards can be used</param>
        /// <param name="index">index of Label, starts with 1</param>
        /// <param name="searchDescendants">true if search through descendants, false if search only through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_Label element, null if not found</returns>
        public UIDA_Label LabelAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("LabelAt method - index cannot be negative");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("LabelAt method - index cannot be negative");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_TextControlTypeId, name, index, searchDescendants,
                false, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("LabelAt method - Label element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("LabelAt method - Label element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("LabelAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("LabelAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_Label label = new UIDA_Label(returnElement);
            return label;
        }

        /// <summary>
        /// Returns a collection of Labels that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of Label elements, use null to return all Labels</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_Label elements</returns>
        public UIDA_Label[] Labels(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allLabels = FindAll(UIA_ControlTypeIds.UIA_TextControlTypeId,
                name, searchDescendants, false, caseSensitive);

            List<UIDA_Label> labels = new List<UIDA_Label>();
            if (allLabels != null)
            {
                foreach (IUIAutomationElement crtEl in allLabels)
                {
                    labels.Add(new UIDA_Label(crtEl));
                }
            }
            return labels.ToArray();
        }

        /// <summary>
        /// Searches a Tree control in the current element
        /// </summary>
        /// <param name="name">name or label of tree control, wildcards can be used</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is case sensitive, default true</param>
        /// <returns>UIDA_Tree element, null if not found</returns>
        public UIDA_Tree Tree(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_TreeControlTypeId,
                name, searchDescendants, true, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("Tree method - Tree element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("Tree method - Tree element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_Tree tree = new UIDA_Tree(returnElement);
            return tree;
        }

        /// <summary>
        /// Searches for a Tree control with a specified name or label at a specified index.
        /// </summary>
        /// <param name="name">name or label of Tree control, wildcards can be used</param>
        /// <param name="index">index of Tree control, starts with 1</param>
        /// <param name="searchDescendants">true if search through descendants, false if search only through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_Tree element, null if not found</returns>
        public UIDA_Tree TreeAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("TreeAt method - index cannot be negative");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("TreeAt method - index cannot be negative");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_TreeControlTypeId, name, index, searchDescendants,
                true, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("TreeAt method - Tree element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("TreeAt method - Tree element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("TreeAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("TreeAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_Tree tree = new UIDA_Tree(returnElement);
            return tree;
        }

        /// <summary>
        /// Returns a collection of Trees that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of Tree elements, use null to return all Trees</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_Tree elements</returns>
        public UIDA_Tree[] Trees(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allTrees = FindAll(UIA_ControlTypeIds.UIA_TreeControlTypeId,
                name, searchDescendants, true, caseSensitive);

            List<UIDA_Tree> trees = new List<UIDA_Tree>();
            if (allTrees != null)
            {
                foreach (IUIAutomationElement crtEl in allTrees)
                {
                    trees.Add(new UIDA_Tree(crtEl));
                }
            }
            return trees.ToArray();
        }

        /// <summary>
        /// Searches for a spinner among the children of the element or its descendants
        /// </summary>
        /// <param name="name">spinner name or label, wildcards can be used</param>
        /// <param name="index">spinner index, starts with 1</param>
        /// <param name="searchDescendants">search descendants, default false</param>
        /// <param name="caseSensitive">search name with case sensitive criteria</param>
        /// <returns>UIDA_Spinner, null if not found</returns>
        public UIDA_Spinner SpinnerAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("SpinnerAt method - index cannot be less than zero");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("SpinnerAt method - index cannot be less than zero");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_SpinnerControlTypeId, name, index, searchDescendants,
                true, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("SpinnerAt method - button element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("SpinnerAt method - button element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("SpinnerAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("SpinnerAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_Spinner spinner = new UIDA_Spinner(returnElement);
            return spinner;
        }

        /// <summary>
        /// Searches a Spinner control in the current element
        /// </summary>
        /// <param name="name">name or label of spinner control, wildcards can be used</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is case sensitive, default true</param>
        /// <returns>UIDA_Spinner element, null if not found</returns>
        public UIDA_Spinner Spinner(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_SpinnerControlTypeId,
                name, searchDescendants, true, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("Spinner method - Spinner element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("Spinner method - Spinner element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_Spinner spinner = new UIDA_Spinner(returnElement);
            return spinner;
        }

        /// <summary>
        /// Returns a collection of Spinners that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of Spinner elements, use null to return all Spinners</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_Spinner elements</returns>
        public UIDA_Spinner[] Spinners(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allSpinners = FindAll(UIA_ControlTypeIds.UIA_SpinnerControlTypeId,
                name, searchDescendants, true, caseSensitive);

            List<UIDA_Spinner> spinners = new List<UIDA_Spinner>();
            if (allSpinners != null)
            {
                foreach (IUIAutomationElement crtEl in allSpinners)
                {
                    spinners.Add(new UIDA_Spinner(crtEl));
                }
            }
            return spinners.ToArray();
        }

        /// <summary>
        /// Searches for a slider among the children of the element or its descendants
        /// </summary>
        /// <param name="name">slider name or label, wildcards can be used</param>
        /// <param name="index">slider index, starts with 1</param>
        /// <param name="searchDescendants">search descendants, default false</param>
        /// <param name="caseSensitive">search name with case sensitive criteria</param>
        /// <returns>UIDA_Slider, null if not found</returns>
        public UIDA_Slider SliderAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("SliderAt method - index cannot be less than zero");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("SliderAt method - index cannot be less than zero");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_SliderControlTypeId, name, index, searchDescendants,
                true, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("SliderAt method - button element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("SliderAt method - button element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("SliderAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("SliderAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_Slider slider = new UIDA_Slider(returnElement);
            return slider;
        }

        /// <summary>
        /// Searches a Slider control in the current element
        /// </summary>
        /// <param name="name">name or label of slider control, wildcards can be used</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is case sensitive, default true</param>
        /// <returns>UIDA_Slider element, null if not found</returns>
        public UIDA_Slider Slider(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_SliderControlTypeId,
                name, searchDescendants, true, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("Slider method - Slider element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("Slider method - Slider element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_Slider slider = new UIDA_Slider(returnElement);
            return slider;
        }

        /// <summary>
        /// Returns a collection of Sliders that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of Slider elements, use null to return all Sliders</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_Slider elements</returns>
        public UIDA_Slider[] Sliders(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allSliders = FindAll(UIA_ControlTypeIds.UIA_SliderControlTypeId,
                name, searchDescendants, true, caseSensitive);

            List<UIDA_Slider> sliders = new List<UIDA_Slider>();
            if (allSliders != null)
            {
                foreach (IUIAutomationElement crtEl in allSliders)
                {
                    sliders.Add(new UIDA_Slider(crtEl));
                }
            }
            return sliders.ToArray();
        }

        /// <summary>
        /// Searches for a progressbar among the children of the element or its descendants
        /// </summary>
        /// <param name="name">progressbar name or label, wildcards can be used</param>
        /// <param name="index">progressbar index, starts with 1</param>
        /// <param name="searchDescendants">search descendants, default false</param>
        /// <param name="caseSensitive">search name with case sensitive criteria</param>
        /// <returns>UIDA_ProgressBar, null if not found</returns>
        public UIDA_ProgressBar ProgressBarAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("ProgressBarAt method - index cannot be less than zero");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ProgressBarAt method - index cannot be less than zero");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_ProgressBarControlTypeId, name, index, searchDescendants,
                true, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("ProgressBarAt method - button element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ProgressBarAt method - button element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("ProgressBarAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ProgressBarAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_ProgressBar progressbar = new UIDA_ProgressBar(returnElement);
            return progressbar;
        }

        /// <summary>
        /// Searches a ProgressBar control in the current element
        /// </summary>
        /// <param name="name">name or label of progressbar control, wildcards can be used</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is case sensitive, default true</param>
        /// <returns>UIDA_ProgressBar element, null if not found</returns>
        public UIDA_ProgressBar ProgressBar(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_ProgressBarControlTypeId,
                name, searchDescendants, true, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("ProgressBar method - ProgressBar element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ProgressBar method - ProgressBar element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_ProgressBar progressbar = new UIDA_ProgressBar(returnElement);
            return progressbar;
        }

        /// <summary>
        /// Returns a collection of ProgressBars that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of ProgressBar elements, use null to return all ProgressBars</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_ProgressBar elements</returns>
        public UIDA_ProgressBar[] ProgressBars(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allProgressBars = FindAll(UIA_ControlTypeIds.UIA_ProgressBarControlTypeId,
                name, searchDescendants, true, caseSensitive);

            List<UIDA_ProgressBar> progressBars = new List<UIDA_ProgressBar>();
            if (allProgressBars != null)
            {
                foreach (IUIAutomationElement crtEl in allProgressBars)
                {
                    progressBars.Add(new UIDA_ProgressBar(crtEl));
                }
            }
            return progressBars.ToArray();
        }

        /// <summary>
        /// Searches a HyperLink control in the current element
        /// </summary>
        /// <param name="name">text of hyperlink control, wildcards can be used</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is case sensitive, default true</param>
        /// <returns>UIDA_HyperLink element, null if not found</returns>
        public UIDA_HyperLink Hyperlink(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_HyperlinkControlTypeId,
                name, searchDescendants, false, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("HyperLink method - HyperLink element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("HyperLink method - HyperLink element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_HyperLink hyperlink = new UIDA_HyperLink(returnElement);
            return hyperlink;
        }

        /// <summary>
        /// Searches for a hyperlink among the children of the element or its descendants
        /// </summary>
        /// <param name="name">text of hyperlink, wildcards can be used</param>
        /// <param name="index">hyperlink index, starts with 1</param>
        /// <param name="searchDescendants">search descendants, default false</param>
        /// <param name="caseSensitive">search name with case sensitive criteria</param>
        /// <returns>UIDA_HyperLink, null if not found</returns>
        public UIDA_HyperLink HyperlinkAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("HyperLinkAt method - index cannot be less than zero");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("HyperLinkAt method - index cannot be less than zero");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_HyperlinkControlTypeId, name, index, searchDescendants,
                false, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("HyperLinkAt method - hyperlink element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("HyperLinkAt method - hyperlink element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("HyperLinkAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("HyperLinkAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_HyperLink hyperlink = new UIDA_HyperLink(returnElement);
            return hyperlink;
        }

        /// <summary>
        /// Returns a collection of HyperLinks that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of HyperLink elements, use null to return all HyperLinks</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_HyperLink elements</returns>
        public UIDA_HyperLink[] HyperLinks(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allHyperLinks = FindAll(UIA_ControlTypeIds.UIA_HyperlinkControlTypeId,
                name, searchDescendants, false, caseSensitive);

            List<UIDA_HyperLink> hyperlinks = new List<UIDA_HyperLink>();
            if (allHyperLinks != null)
            {
                foreach (IUIAutomationElement crtEl in allHyperLinks)
                {
                    hyperlinks.Add(new UIDA_HyperLink(crtEl));
                }
            }
            return hyperlinks.ToArray();
        }

        /// <summary>
        /// Searches a Tab control in the current element
        /// </summary>
        /// <param name="name">name of Tab control</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is case sensitive, default true</param>
        /// <returns>UIDA_TabCtrl element, null if not found</returns>
        public UIDA_TabCtrl TabCtrl(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_TabControlTypeId,
                name, searchDescendants, false, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("TabCtrl method - Tab element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("TabCtrl method - Tab element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_TabCtrl tabCtrl = new UIDA_TabCtrl(returnElement);
            return tabCtrl;
        }

        /// <summary>
        /// Searches for a tab control among the children of the element or its descendants
        /// </summary>
        /// <param name="name">name of tab control</param>
        /// <param name="index">tab control index, starts with 1</param>
        /// <param name="searchDescendants">search descendants, default false</param>
        /// <param name="caseSensitive">search name with case sensitive criteria</param>
        /// <returns>UIDA_TabCtrl, null if not found</returns>
        public UIDA_TabCtrl TabCtrlAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("TabCtrlAt method - index cannot be less than zero");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("TabCtrlAt method - index cannot be less than zero");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_TabControlTypeId, name, index, searchDescendants,
                false, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("TabCtrlAt method - tab element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("TabCtrlAt method - tab element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("TabCtrlAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("TabCtrlAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_TabCtrl tabCtrl = new UIDA_TabCtrl(returnElement);
            return tabCtrl;
        }

        /// <summary>
        /// Returns a collection of TabCtrls that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of TabCtrl elements, use null to return all TabCtrls</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_TabCtrl elements</returns>
        public UIDA_TabCtrl[] TabCtrls(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allTabCtrls = FindAll(UIA_ControlTypeIds.UIA_TabControlTypeId,
                name, searchDescendants, false, caseSensitive);

            List<UIDA_TabCtrl> tabctrls = new List<UIDA_TabCtrl>();
            if (allTabCtrls != null)
            {
                foreach (IUIAutomationElement crtEl in allTabCtrls)
                {
                    tabctrls.Add(new UIDA_TabCtrl(crtEl));
                }
            }
            return tabctrls.ToArray();
        }
		
		/// <summary>
        /// Searches a TabItem page in the current element
        /// </summary>
        /// <param name="name">name of tab item</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is case sensitive, default true</param>
        /// <returns>UIDA_TabItem element, null if not found</returns>
        public UIDA_TabItem TabItem(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_TabItemControlTypeId,
                name, searchDescendants, false, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("TabItem method - TabItem element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("TabItem method - TabItem element not found");
                }
                else
                {
                    return null;
                }
            }

			UIDA_TabItem tabItem = null;
			if (this.uiElement.CurrentControlType == UIA_ControlTypeIds.UIA_TabControlTypeId)
			{
				tabItem = new UIDA_TabItem(returnElement, this as UIDA_TabCtrl);
			}
			else
			{
				tabItem = new UIDA_TabItem(returnElement);
			}
            return tabItem;
        }

        /// <summary>
        /// Searches for a tab item among the children or descendants of the current element
        /// or its descendants.
        /// </summary>
        /// <param name="name">name of tab item</param>
        /// <param name="index">tab item index, starts with 1</param>
        /// <param name="searchDescendants">search descendants, default false</param>
        /// <param name="caseSensitive">search name with case sensitive criteria</param>
        /// <returns>UIDA_TabItem, null if not found</returns>
        public UIDA_TabItem TabItemAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("TabItemAt method - index cannot be less than zero");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("TabItemAt method - index cannot be less than zero");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_TabItemControlTypeId, name, index, searchDescendants,
                false, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("TabItemAt method - button element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("TabItemAt method - button element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("TabItemAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("TabItemAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

			UIDA_TabItem tabItem = null;
			if (this.uiElement.CurrentControlType == UIA_ControlTypeIds.UIA_TabControlTypeId)
			{
				tabItem = new UIDA_TabItem(returnElement, this as UIDA_TabCtrl);
			}
			else
			{
				tabItem = new UIDA_TabItem(returnElement);
			}
            return tabItem;
        }

        /// <summary>
        /// Returns a collection of TabItems that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of TabItem elements, use null to return all TabItems</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_TabItem elements</returns>
        public UIDA_TabItem[] TabItems(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allTabItems = FindAll(UIA_ControlTypeIds.UIA_TabItemControlTypeId,
                name, searchDescendants, false, caseSensitive);

			bool isTabCtrl = (this.uiElement.CurrentControlType == UIA_ControlTypeIds.UIA_TabControlTypeId);
            List<UIDA_TabItem> tabitems = new List<UIDA_TabItem>();
            if (allTabItems != null)
            {
				if (isTabCtrl)
				{
					foreach (IUIAutomationElement crtEl in allTabItems)
					{
						tabitems.Add(new UIDA_TabItem(crtEl, this as UIDA_TabCtrl));
					}
				}
				else
				{
					foreach (IUIAutomationElement crtEl in allTabItems)
					{
						tabitems.Add(new UIDA_TabItem(crtEl));
					}
				}
            }
            return tabitems.ToArray();
        }

        /// <summary>
        /// Searches a Image control in the current element
        /// </summary>
        /// <param name="name">name or label of Image control, wildcards can be used</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is case sensitive, default true</param>
        /// <returns>UIDA_Image element, null if not found</returns>
        public UIDA_Image Image(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_ImageControlTypeId,
                name, searchDescendants, true, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("Image method - Image element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("Image method - Image element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_Image image = new UIDA_Image(returnElement);
            return image;
        }

        /// <summary>
        /// Searches for a image control among the children of the element or its descendants
        /// </summary>
        /// <param name="name">name or label of image control, wildcards can be used</param>
        /// <param name="index">image control index, starts with 1</param>
        /// <param name="searchDescendants">search descendants, default false</param>
        /// <param name="caseSensitive">search name with case sensitive criteria</param>
        /// <returns>UIDA_Image, null if not found</returns>
        public UIDA_Image ImageAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("ImageAt method - index cannot be less than zero");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ImageAt method - index cannot be less than zero");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_ImageControlTypeId, name, index, searchDescendants,
                true, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("ImageAt method - image element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ImageAt method - image element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("ImageAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ImageAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_Image image = new UIDA_Image(returnElement);
            return image;
        }

        /// <summary>
        /// Returns a collection of Images that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of Image elements, use null to return all Images</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_Image elements</returns>
        public UIDA_Image[] Images(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allImages = FindAll(UIA_ControlTypeIds.UIA_ImageControlTypeId,
                name, searchDescendants, true, caseSensitive);

            List<UIDA_Image> images = new List<UIDA_Image>();
            if (allImages != null)
            {
                foreach (IUIAutomationElement crtEl in allImages)
                {
                    images.Add(new UIDA_Image(crtEl));
                }
            }
            return images.ToArray();
        }

        /// <summary>
        /// Searches a ScrollBar control in the current element.
        /// </summary>
        /// <param name="name">name of ScrollBar control, wildcards can be used</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is case sensitive, default true</param>
        /// <returns>UIDA_ScrollBar element, null if not found</returns>
        public UIDA_ScrollBar ScrollBar(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_ScrollBarControlTypeId,
                name, searchDescendants, false, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("ScrollBar method - ScrollBar element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ScrollBar method - ScrollBar element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_ScrollBar scrollbar = new UIDA_ScrollBar(returnElement);
            return scrollbar;
        }

        /// <summary>
        /// Searches for a scrollbar control among the children of the element or its descendants
        /// </summary>
        /// <param name="name">name of scrollbar control, wildcards can be used</param>
        /// <param name="index">scrollbar control index, starts with 1</param>
        /// <param name="searchDescendants">search descendants, default false</param>
        /// <param name="caseSensitive">search name with case sensitive criteria</param>
        /// <returns>UIDA_ScrollBar, null if not found</returns>
        public UIDA_ScrollBar ScrollBarAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("ScrollBarAt method - index cannot be less than zero");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ScrollBarAt method - index cannot be less than zero");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_ScrollBarControlTypeId, name, index, searchDescendants,
                false, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("ScrollBarAt method - scrollbar element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ScrollBarAt method - scrollbar element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("ScrollBarAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ScrollBarAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_ScrollBar scrollbar = new UIDA_ScrollBar(returnElement);
            return scrollbar;
        }

        /// <summary>
        /// Returns a collection of ScrollBars that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of ScrollBar elements, use null to return all ScrollBars</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_ScrollBar elements</returns>
        public UIDA_ScrollBar[] ScrollBars(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allScrollBars = FindAll(UIA_ControlTypeIds.UIA_ScrollBarControlTypeId,
                name, searchDescendants, false, caseSensitive);

            List<UIDA_ScrollBar> scrollbars = new List<UIDA_ScrollBar>();
            if (allScrollBars != null)
            {
                foreach (IUIAutomationElement crtEl in allScrollBars)
                {
                    scrollbars.Add(new UIDA_ScrollBar(crtEl));
                }
            }
            return scrollbars.ToArray();
        }

        /// <summary>
        /// Searches a Table control in the current element.
        /// </summary>
        /// <param name="name">name or label of Table control, wildcards can be used</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is case sensitive, default true</param>
        /// <returns>UIDA_Table element, null if not found</returns>
        public UIDA_Table Table(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_TableControlTypeId,
                name, searchDescendants, true, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("Table method - Table element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("Table method - Table element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_Table table = new UIDA_Table(returnElement);
            return table;
        }

        /// <summary>
        /// Searches for a table control among the children of the element or its descendants
        /// </summary>
        /// <param name="name">name or label of table control, wildcards can be used</param>
        /// <param name="index">table control index, starts with 1</param>
        /// <param name="searchDescendants">search descendants, default false</param>
        /// <param name="caseSensitive">search name with case sensitive criteria</param>
        /// <returns>UIDA_Table element, null if not found</returns>
        public UIDA_Table TableAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("TableAt method - index cannot be less than zero");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("TableAt method - index cannot be less than zero");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_TableControlTypeId, name, index, searchDescendants,
                true, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("TableAt method - table element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("TableAt method - table element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("TableAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("TableAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_Table table = new UIDA_Table(returnElement);
            return table;
        }
		
		/// <summary>
        /// Returns a collection of Tables that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of Table elements</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_Table collection</returns>
        public UIDA_Table[] Tables(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allTables = FindAll(UIA_ControlTypeIds.UIA_TableControlTypeId,
                name, searchDescendants, false, caseSensitive);

			List<UIDA_Table> tables = new List<UIDA_Table>();
            if (allTables != null)
            {
                foreach (IUIAutomationElement crtEl in allTables)
                {
                    tables.Add(new UIDA_Table(crtEl));
                }
            }
            return tables.ToArray();
        }

        /// <summary>
        /// Searches for a datagrid control among the children of the element or its descendants
        /// </summary>
        /// <param name="name">name or label of datagrid control, wildcards can be used</param>
        /// <param name="index">datagrid control index, starts with 1</param>
        /// <param name="searchDescendants">search descendants, default false</param>
        /// <param name="caseSensitive">search name with case sensitive criteria</param>
        /// <returns>UIDA_DataGrid element, null if not found</returns>
        public UIDA_DataGrid DataGridAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("DataGridAt method - index cannot be less than zero");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("DataGridAt method - index cannot be less than zero");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_DataGridControlTypeId, name, index, searchDescendants,
                true, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("DataGridAt method - datagrid element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("DataGridAt method - datagrid element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("DataGridAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("DataGridAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_DataGrid dataGrid = new UIDA_DataGrid(returnElement);
            return dataGrid;
        }

        /// <summary>
        /// Searches a DataGrid control in the current element.
        /// </summary>
        /// <param name="name">name or label of DataGrid control, wildcards can be used</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is case sensitive, default true</param>
        /// <returns>UIDA_DataGrid element, null if not found</returns>
        public UIDA_DataGrid DataGrid(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_DataGridControlTypeId,
                name, searchDescendants, true, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("DataGrid method - DataGrid element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("DataGrid method - DataGrid element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_DataGrid dataGrid = new UIDA_DataGrid(returnElement);
            return dataGrid;
        }
		
		/// <summary>
        /// Returns a collection of DataGrids that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of DataGrid elements</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_DataGrid collection</returns>
        public UIDA_DataGrid[] DataGrids(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allDataGrids = FindAll(UIA_ControlTypeIds.UIA_DataGridControlTypeId,
                name, searchDescendants, false, caseSensitive);

			List<UIDA_DataGrid> dataGrids = new List<UIDA_DataGrid>();
            if (allDataGrids != null)
            {
                foreach (IUIAutomationElement crtEl in allDataGrids)
                {
                    dataGrids.Add(new UIDA_DataGrid(crtEl));
                }
            }
            return dataGrids.ToArray();
        }

        /// <summary>
        /// Searches for a Custom control with a specified text at a specified index.
        /// </summary>
        /// <param name="name">text of Custom control, wildcards can be used</param>
        /// <param name="index">index of Custom control, starts with 1</param>
        /// <param name="searchDescendants">true if search through descendants, false if search only through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_Custom element, null if not found</returns>
        public UIDA_Custom CustomAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("CustomAt method - index cannot be negative");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("CustomAt method - index cannot be negative");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_CustomControlTypeId, name, index, searchDescendants,
                false, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("CustomAt method - Custom element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("CustomAt method - Custom element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("CustomAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("CustomAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_Custom custom = new UIDA_Custom(returnElement);
            return custom;
        }

        /// <summary>
        /// Searches a Custom control in the current element.
        /// </summary>
        /// <param name="name">text of Custom control, wildcards can be used</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is case sensitive, default true</param>
        /// <returns>UIDA_Custom element, null if not found</returns>
        public UIDA_Custom Custom(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_CustomControlTypeId,
                name, searchDescendants, false, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("Custom method - Custom element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("Custom method - Custom element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_Custom custom = new UIDA_Custom(returnElement);
            return custom;
        }

        /// <summary>
        /// Returns a collection of Customs that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of Custom elements, use null to return all Custom elements</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_Custom elements</returns>
        public UIDA_Custom[] Customs(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allCustoms = FindAll(UIA_ControlTypeIds.UIA_CustomControlTypeId,
                name, searchDescendants, false, caseSensitive);

            List<UIDA_Custom> customs = new List<UIDA_Custom>();
            if (allCustoms != null)
            {
                foreach (IUIAutomationElement crtEl in allCustoms)
                {
                    customs.Add(new UIDA_Custom(crtEl));
                }
            }
            return customs.ToArray();
        }

        /// <summary>
        /// Searches for a Separator control with a specified text at a specified index.
        /// </summary>
        /// <param name="name">text of Separator control, wildcards can be used</param>
        /// <param name="index">index of Separator control, starts with 1</param>
        /// <param name="searchDescendants">true if search through descendants, false if search only through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_Separator element, null if not found</returns>
        public UIDA_Separator SeparatorAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("SeparatorAt method - index cannot be negative");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("SeparatorAt method - index cannot be negative");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_SeparatorControlTypeId, name, index, searchDescendants,
                false, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("SeparatorAt method - Separator element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("SeparatorAt method - Separator element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("SeparatorAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("SeparatorAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_Separator separator = new UIDA_Separator(returnElement);
            return separator;
        }

        /// <summary>
        /// Searches a Separator control in the current element.
        /// </summary>
        /// <param name="name">text of Separator control, wildcards can be used</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is case sensitive, default true</param>
        /// <returns>UIDA_Separator element, null if not found</returns>
        public UIDA_Separator Separator(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_SeparatorControlTypeId,
                name, searchDescendants, false, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("Separator method - Separator element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("Separator method - Separator element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_Separator separator = new UIDA_Separator(returnElement);
            return separator;
        }

        /// <summary>
        /// Returns a collection of Separators that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of Separator elements, use null to return all Separators</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UID_Separator elements</returns>
        public UIDA_Separator[] Separators(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allSeparators = FindAll(UIA_ControlTypeIds.UIA_SeparatorControlTypeId,
                name, searchDescendants, false, caseSensitive);

            List<UIDA_Separator> separators = new List<UIDA_Separator>();
            if (allSeparators != null)
            {
                foreach (IUIAutomationElement crtEl in allSeparators)
                {
                    separators.Add(new UIDA_Separator(crtEl));
                }
            }
            return separators.ToArray();
        }

        /// <summary>
        /// Searches for a SplitButton control with a specified text at a specified index.
        /// </summary>
        /// <param name="name">text of SplitButton control, wildcards can be used</param>
        /// <param name="index">index of SplitButton control, starts with 1</param>
        /// <param name="searchDescendants">true if search through descendants, false if search only through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_SplitButton element, null if not found</returns>
        public UIDA_SplitButton SplitButtonAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("SplitButtonAt method - index cannot be negative");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("SplitButtonAt method - index cannot be negative");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_SplitButtonControlTypeId, name, index, searchDescendants,
                false, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("SplitButtonAt method - SplitButton element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("SplitButtonAt method - SplitButton element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("SplitButtonAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("SplitButtonAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_SplitButton splitButton = new UIDA_SplitButton(returnElement);
            return splitButton;
        }

        /// <summary>
        /// Searches a SplitButton control in the current element.
        /// </summary>
        /// <param name="name">text of SplitButton control, wildcards can be used</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is case sensitive, default true</param>
        /// <returns>UIDA_SplitButton element, null if not found</returns>
        public UIDA_SplitButton SplitButton(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_SplitButtonControlTypeId,
                name, searchDescendants, false, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("SplitButton method - SplitButton element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("SplitButton method - SplitButton element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_SplitButton splitButton = new UIDA_SplitButton(returnElement);
            return splitButton;
        }

        /// <summary>
        /// Returns a collection of SplitButtons that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of SplitButton elements, use null to return all SplitButtons</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_SplitButton elements</returns>
        public UIDA_SplitButton[] SplitButtons(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allSplitButtons = FindAll(UIA_ControlTypeIds.UIA_SplitButtonControlTypeId,
                name, searchDescendants, false, caseSensitive);

            List<UIDA_SplitButton> splitButtons = new List<UIDA_SplitButton>();
            if (allSplitButtons != null)
            {
                foreach (IUIAutomationElement crtEl in allSplitButtons)
                {
                    splitButtons.Add(new UIDA_SplitButton(crtEl));
                }
            }
            return splitButtons.ToArray();
        }

        /// <summary>
        /// Searches for a StatusBar control with a specified text at a specified index.
        /// </summary>
        /// <param name="name">text of StatusBar control, wildcards can be used</param>
        /// <param name="index">index of StatusBar control, starts with 1</param>
        /// <param name="searchDescendants">true if search through descendants, false if search only through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_StatusBar element, null if not found</returns>
        public UIDA_StatusBar StatusBarAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("StatusBarAt method - index cannot be negative");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("StatusBarAt method - index cannot be negative");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_StatusBarControlTypeId, name, index, searchDescendants,
                false, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("StatusBarAt method - StatusBar element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("StatusBarAt method - StatusBar element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("StatusBarAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("StatusBarAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_StatusBar statusBar = new UIDA_StatusBar(returnElement);
            return statusBar;
        }

        /// <summary>
        /// Searches a StatusBar control in the current element.
        /// </summary>
        /// <param name="name">text of StatusBar control, wildcards can be used</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is case sensitive, default true</param>
        /// <returns>UIDA_StatusBar element, null if not found</returns>
        public UIDA_StatusBar StatusBar(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_StatusBarControlTypeId,
                name, searchDescendants, false, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("StatusBar method - StatusBar element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("StatusBar method - StatusBar element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_StatusBar statusBar = new UIDA_StatusBar(returnElement);
            return statusBar;
        }
		
		/// <summary>
        /// Returns a collection of StatusBars that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of StatusBar elements</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_StatusBar collection</returns>
        public UIDA_StatusBar[] StatusBars(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allStatusBars = FindAll(UIA_ControlTypeIds.UIA_StatusBarControlTypeId,
                name, searchDescendants, false, caseSensitive);

			List<UIDA_StatusBar> statusBars = new List<UIDA_StatusBar>();
            if (allStatusBars != null)
            {
                foreach (IUIAutomationElement crtEl in allStatusBars)
                {
                    statusBars.Add(new UIDA_StatusBar(crtEl));
                }
            }
            return statusBars.ToArray();
        }

        /// <summary>
        /// Searches for a Thumb control with a specified text at a specified index.
        /// </summary>
        /// <param name="name">text or name of Thumb control, wildcards can be used</param>
        /// <param name="index">index of Thumb control, starts with 1</param>
        /// <param name="searchDescendants">true if search through descendants, false if search only through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_Thumb element, null if not found</returns>
        public UIDA_Thumb ThumbAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("ThumbAt method - index cannot be negative");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ThumbAt method - index cannot be negative");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_ThumbControlTypeId, name, index, searchDescendants,
                false, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("ThumbAt method - Thumb element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ThumbAt method - Thumb element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("ThumbAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ThumbAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_Thumb thumb = new UIDA_Thumb(returnElement);
            return thumb;
        }

        /// <summary>
        /// Searches a Thumb control in the current element.
        /// </summary>
        /// <param name="name">text or name of Thumb control, wildcards can be used</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is case sensitive, default true</param>
        /// <returns>UIDA_Thumb element, null if not found</returns>
        public UIDA_Thumb Thumb(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_ThumbControlTypeId,
                name, searchDescendants, false, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("Thumb method - Thumb element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("Thumb method - Thumb element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_Thumb thumb = new UIDA_Thumb(returnElement);
            return thumb;
        }
		
		/// <summary>
        /// Returns a collection of Thumbs that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of Thumb elements</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_Thumb collection</returns>
        public UIDA_Thumb[] Thumbs(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allThumbs = FindAll(UIA_ControlTypeIds.UIA_ThumbControlTypeId,
                name, searchDescendants, false, caseSensitive);

			List<UIDA_Thumb> thumbs = new List<UIDA_Thumb>();
            if (allThumbs != null)
            {
                foreach (IUIAutomationElement crtEl in allThumbs)
                {
                    thumbs.Add(new UIDA_Thumb(crtEl));
                }
            }
            return thumbs.ToArray();
        }

        /// <summary>
        /// Searches for a ToolBar control with a specified text at a specified index.
        /// </summary>
        /// <param name="name">text or name of ToolBar control, wildcards can be used</param>
        /// <param name="index">index of ToolBar control, starts with 1</param>
        /// <param name="searchDescendants">true if search through descendants, false if search only through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_ToolBar element, null if not found</returns>
        public UIDA_ToolBar ToolBarAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("ToolBarAt method - index cannot be negative");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ToolBarAt method - index cannot be negative");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_ToolBarControlTypeId, name, index, searchDescendants,
                false, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("Toolbar method - Toolbar element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("Toolbar method - Toolbar element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("Toolbar method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("Toolbar method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_ToolBar toolbar = new UIDA_ToolBar(returnElement);
            return toolbar;
        }

        /// <summary>
        /// Searches a ToolBar control in the current element.
        /// </summary>
        /// <param name="name">text or name of ToolBar control, wildcards can be used</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is case sensitive, default true</param>
        /// <returns>UIDA_ToolBar element, null if not found</returns>
        public UIDA_ToolBar ToolBar(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_ToolBarControlTypeId,
                name, searchDescendants, false, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("ToolBar method - ToolBar element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ToolBar method - ToolBar element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_ToolBar toolbar = new UIDA_ToolBar(returnElement);
            return toolbar;
        }
		
		/// <summary>
        /// Returns a collection of ToolBars that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of ToolBar elements</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_ToolBar collection</returns>
        public UIDA_ToolBar[] ToolBars(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allToolbars = FindAll(UIA_ControlTypeIds.UIA_ToolBarControlTypeId,
                name, searchDescendants, false, caseSensitive);

			List<UIDA_ToolBar> toolbars = new List<UIDA_ToolBar>();
            if (allToolbars != null)
            {
                foreach (IUIAutomationElement crtEl in allToolbars)
                {
                    toolbars.Add(new UIDA_ToolBar(crtEl));
                }
            }
            return toolbars.ToArray();
        }

        /// <summary>
        /// Searches for a ToolTip control with a specified text at a specified index.
        /// </summary>
        /// <param name="name">text of ToolTip control, wildcards can be used</param>
        /// <param name="index">index of ToolTip control, starts with 1</param>
        /// <param name="searchDescendants">true if search through descendants, false if search only through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_ToolTip element, null if not found</returns>
        public UIDA_ToolTip ToolTipAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("ToolTipAt method - index cannot be negative");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ToolTipAt method - index cannot be negative");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_ToolTipControlTypeId, name, index, searchDescendants,
                false, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("ToolTip method - ToolTip element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ToolTip method - ToolTip element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("ToolTip method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ToolTip method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_ToolTip tooltip = new UIDA_ToolTip(returnElement);
            return tooltip;
        }

        /// <summary>
        /// Searches a ToolTip control in the current element.
        /// </summary>
        /// <param name="name">text of ToolTip control, wildcards can be used</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is case sensitive, default true</param>
        /// <returns>UIDA_ToolTip element, null if not found</returns>
        public UIDA_ToolTip ToolTip(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_ToolTipControlTypeId,
                name, searchDescendants, false, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("ToolTip method - ToolTip element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ToolTip method - ToolTip element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_ToolTip tooltip = new UIDA_ToolTip(returnElement);
            return tooltip;
        }
		
		/// <summary>
        /// Returns a collection of ToolTips that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of ToolTip elements</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_ToolTip collection</returns>
        public UIDA_ToolTip[] ToolTips(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allTooltips = FindAll(UIA_ControlTypeIds.UIA_ToolTipControlTypeId,
                name, searchDescendants, false, caseSensitive);

			List<UIDA_ToolTip> tooltips = new List<UIDA_ToolTip>();
            if (allTooltips != null)
            {
                foreach (IUIAutomationElement crtEl in allTooltips)
                {
                    tooltips.Add(new UIDA_ToolTip(crtEl));
                }
            }
            return tooltips.ToArray();
        }
		
		/// <summary>
        /// Searches a Group control in the current element.
        /// </summary>
        /// <param name="name">text of Group control, wildcards can be used</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is case sensitive, default true</param>
        /// <returns>UIDA_Group element, null if not found</returns>
        public UIDA_Group Group(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_GroupControlTypeId,
                name, searchDescendants, false, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("Group method - Group element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("Group method - Group element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_Group group = new UIDA_Group(returnElement);
            return group;
        }
		
		/// <summary>
        /// Searches for a Group with a specified text at a specified index.
        /// </summary>
        /// <param name="name">text of Group control, wildcards can be used</param>
        /// <param name="index">index of Group control, starts with 1</param>
        /// <param name="searchDescendants">true if search through descendants, false if search only through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_Group element, null if not found</returns>
        public UIDA_Group GroupAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("GroupAt method - index cannot be negative");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("GroupAt method - index cannot be negative");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_GroupControlTypeId, name, index, searchDescendants,
                false, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("GroupAt method - Group element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("GroupAt method - Group element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("GroupAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("GroupAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_Group group = new UIDA_Group(returnElement);
            return group;
        }
		
		/// <summary>
        /// Returns a collection of Groups that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of Group elements</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_Group collection</returns>
        public UIDA_Group[] Groups(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allGroups = FindAll(UIA_ControlTypeIds.UIA_GroupControlTypeId,
                name, searchDescendants, false, caseSensitive);

			List<UIDA_Group> groups = new List<UIDA_Group>();
            if (allGroups != null)
            {
                foreach (IUIAutomationElement crtEl in allGroups)
                {
                    groups.Add(new UIDA_Group(crtEl));
                }
            }
            return groups.ToArray();
        }
		
		/// <summary>
        /// Searches a header in the current element
        /// </summary>
        /// <param name="name">text of header</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is case sensitive, default true</param>
        /// <returns>UIDA_Header element, null if not found</returns>
        public UIDA_Header Header(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_HeaderControlTypeId,
                name, searchDescendants, false, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("Header method - Header element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("Header method - Header element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_Header header = new UIDA_Header(returnElement);
            return header;
        }
		
		/// <summary>
        /// Searches for a Header with a specified text at a specified index.
        /// </summary>
        /// <param name="name">text of Header</param>
        /// <param name="index">index of Header</param>
        /// <param name="searchDescendants">true if search through descendants, false if search only through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_Header element, null if not found</returns>
        public UIDA_Header HeaderAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("HeaderAt method - index cannot be negative");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("HeaderAt method - index cannot be negative");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_HeaderControlTypeId, name, index, searchDescendants,
                false, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("HeaderAt method - Header element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("HeaderAt method - Header element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("HeaderAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("HeaderAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_Header header = new UIDA_Header(returnElement);
            return header;
        }
		
		/// <summary>
        /// Returns a collection of Headers that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of Header elements</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_Header collection</returns>
        public UIDA_Header[] Headers(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allHeaders = FindAll(UIA_ControlTypeIds.UIA_HeaderControlTypeId,
                name, searchDescendants, false, caseSensitive);

			List<UIDA_Header> headers = new List<UIDA_Header>();
			foreach (IUIAutomationElement element in allHeaders)
			{
				UIDA_Header header = new UIDA_Header(element);
				headers.Add(header);
			}
			
			return headers.ToArray();
        }
		
		/// <summary>
        /// Searches a header item in the current element
        /// </summary>
        /// <param name="name">text of header item</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is case sensitive, default true</param>
        /// <returns>UIDA_HeaderItem element, null if not found</returns>
        public UIDA_HeaderItem HeaderItem(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_HeaderItemControlTypeId,
                name, searchDescendants, false, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("HeaderItem method - HeaderItem element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("HeaderItem method - HeaderItem element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_HeaderItem headerItem = new UIDA_HeaderItem(returnElement);
            return headerItem;
        }

        /// <summary>
        /// Searches for a HeaderItem with a specified text at a specified index.
        /// </summary>
        /// <param name="name">text of HeaderItem</param>
        /// <param name="index">index of HeaderItem</param>
        /// <param name="searchDescendants">true if search through descendants, false if search only through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_HeaderItem element, null if not found</returns>
        public UIDA_HeaderItem HeaderItemAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("HeaderItemAt method - index cannot be negative");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("HeaderItemAt method - index cannot be negative");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_HeaderItemControlTypeId, name, index, searchDescendants,
                false, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("HeaderItemAt method - HeaderItem element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("HeaderItemAt method - HeaderItem element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("HeaderItemAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("HeaderItemAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_HeaderItem headerItem = new UIDA_HeaderItem(returnElement);
            return headerItem;
        }
		
		/// <summary>
        /// Returns a collection of HeaderItems that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of HeaderItem elements</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_HeaderItem collection</returns>
        public UIDA_HeaderItem[] HeaderItems(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allHeaderItems = FindAll(UIA_ControlTypeIds.UIA_HeaderItemControlTypeId,
                name, searchDescendants, false, caseSensitive);

			List<UIDA_HeaderItem> headerItems = new List<UIDA_HeaderItem>();
			foreach (IUIAutomationElement element in allHeaderItems)
			{
				UIDA_HeaderItem headerItem = new UIDA_HeaderItem(element);
				headerItems.Add(headerItem);
			}
			
			return headerItems.ToArray();
        }
		
		/// <summary>
        /// Searches a data item in the current element
        /// </summary>
        /// <param name="name">text of data item</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is case sensitive, default true</param>
        /// <returns>UIDA_DataItem element, null if not found</returns>
        public UIDA_DataItem DataItem(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_DataItemControlTypeId,
                name, searchDescendants, false, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("DataItem method - DataItem element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("DataItem method - DataItem element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_DataItem dataItem = new UIDA_DataItem(returnElement);
            return dataItem;
        }

        /// <summary>
        /// Searches for a DataItem with a specified text at a specified index.
        /// </summary>
        /// <param name="name">text of DataItem</param>
        /// <param name="index">index of DataItem</param>
        /// <param name="searchDescendants">true if search through descendants, false if search only through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_DataItem element, null if not found</returns>
        public UIDA_DataItem DataItemAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("DataItemAt method - index cannot be negative");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("DataItemAt method - index cannot be negative");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_DataItemControlTypeId, name, index, searchDescendants,
                false, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("DataItemAt method - DataItem element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("DataItemAt method - DataItem element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("DataItemAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("DataItemAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_DataItem dataItem = new UIDA_DataItem(returnElement);
            return dataItem;
        }
        
        /// <summary>
        /// Returns a collection of DataItem that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of DataItem elements, use null to return all DataItems</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_DataItem elements</returns>
        public UIDA_DataItem[] DataItems(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allDataItems = FindAll(UIA_ControlTypeIds.UIA_DataItemControlTypeId,
                name, searchDescendants, false, caseSensitive);

            List<UIDA_DataItem> dataitems = new List<UIDA_DataItem>();
            if (allDataItems != null)
            {
                foreach (IUIAutomationElement crtEl in allDataItems)
                {
                    dataitems.Add(new UIDA_DataItem(crtEl));
                }
            }
            return dataitems.ToArray();
        }

        /// <summary>
        /// Searches a Window control in the current element.
        /// </summary>
        /// <param name="name">text of Window control, wildcards can be used</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is case sensitive, default true</param>
        /// <returns>UIDA_Window element, null if not found</returns>
        public UIDA_Window Window(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_WindowControlTypeId,
                name, searchDescendants, false, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("Window method - Window element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("Window method - Window element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_Window window = new UIDA_Window(returnElement);
            return window;
        }

        /// <summary>
        /// Searches for a Window with a specified text at a specified index.
        /// </summary>
        /// <param name="name">text of Window control, wildcards can be used</param>
        /// <param name="index">index of Window control, starts with 1</param>
        /// <param name="searchDescendants">true if search through descendants, false if search only through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_Window element, null if not found</returns>
        public UIDA_Window WindowAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("WindowAt method - index cannot be negative");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("WindowAt method - index cannot be negative");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_WindowControlTypeId, name, index, searchDescendants,
                false, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("WindowAt method - Window element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("WindowAt method - Window element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("WindowAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("WindowAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_Window window = new UIDA_Window(returnElement);
            return window;
        }

        /// <summary>
        /// Returns a collection of Windows that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of Window elements</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_Window elements</returns>
        public UIDA_Window[] Windows(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allWindows = FindAll(UIA_ControlTypeIds.UIA_WindowControlTypeId,
                name, searchDescendants, false, caseSensitive);

        	List<UIDA_Window> windows = new List<UIDA_Window>();
        	
        	foreach (IUIAutomationElement el in allWindows)
        	{
        		UIDA_Window window = new UIDA_Window(el);
        		windows.Add(window);
        	}
        	
            return windows.ToArray();
        }

        /// <summary>
        /// Searches for a Tree Item in the current tree item scope
        /// </summary>
        /// <param name="text">text of tree item, wildcards can be used</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is case sensitive, default true</param>
        /// <returns>UIDA_TreeItem element, null if not found</returns>
        public UIDA_TreeItem TreeItem(string text = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_TreeItemControlTypeId,
                text, searchDescendants, false, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("TreeItem method - TreeItem element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("TreeItem method - TreeItem element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_TreeItem treeItem = new UIDA_TreeItem(returnElement);
            return treeItem;
        }
		
		/// <summary>
        /// Searches for a TreeView item with the specified text at a specified index.
        /// </summary>
        /// <param name="text">text of TreeView item, wildcards can be used</param>
        /// <param name="index">index of TreeView item, starts with 1</param>
        /// <param name="searchDescendants">true if search through descendants, false if search only through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_TreeItem element, null if not found</returns>
        public UIDA_TreeItem TreeItemAt(string text, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("TreeItemAt method - index cannot be negative");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("TreeItemAt method - index cannot be negative");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_TreeItemControlTypeId, text, index, searchDescendants,
                false, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("TreeItemAt method - TreeItem element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("TreeItemAt method - TreeItem element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("TreeItemAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("TreeItemAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_TreeItem treeItem = new UIDA_TreeItem(returnElement);
            return treeItem;
        }
		
		/// <summary>
        /// Returns a collection of TreeItems that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of TreeItem elements</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_TreeItem collection</returns>
        public UIDA_TreeItem[] TreeItems(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allTreeItems = FindAll(UIA_ControlTypeIds.UIA_TreeItemControlTypeId,
                name, searchDescendants, false, caseSensitive);

			List<UIDA_TreeItem> treeItems = new List<UIDA_TreeItem>();
            if (allTreeItems != null)
            {
                foreach (IUIAutomationElement crtEl in allTreeItems)
                {
                    treeItems.Add(new UIDA_TreeItem(crtEl));
                }
            }
            return treeItems.ToArray();
        }
		
		/// <summary>
        /// Searches a list item with a specified name within the current list element
        /// </summary>
        /// <param name="name">text of list item, wildcards accepted</param>
        /// <param name="searchDescendants">true is search through all descendants, false if search only through children collection, default false</param>
        /// <param name="caseSensitive">true is name search is done case sensitive, default true</param>
        /// <returns>UIDA_ListItem element, null if not found</returns>
        public UIDA_ListItem ListItem(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            IUIAutomationElement returnElement = this.FindFirst(UIA_ControlTypeIds.UIA_ListItemControlTypeId, name,
                searchDescendants, false, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("ListItem method - list item element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ListItem method - list item element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_ListItem listItem = new UIDA_ListItem(returnElement);
            return listItem;
        }

        /// <summary>
        /// Searches for a list item with specified name at specified index.
        /// </summary>
        /// <param name="name">text of list item, wildcards accepted</param>
        /// <param name="index">index, starts with 1</param>
        /// <param name="searchDescendants">true is search through descendants, false if search through children, default false</param>
        /// <param name="caseSensitive">true is name search is done case sensitive, false otherwise, default true</param>
        /// <returns>UIDA_ListItem element, null if not found</returns>
        public UIDA_ListItem ListItemAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("ListItemAt method - index cannot be negative");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ListItemAt method - index cannot be negative");
                }
                else
                {
                    return null;
                }
            }

            IUIAutomationElement returnElement = null;

            Errors error = this.FindAt(UIA_ControlTypeIds.UIA_ListItemControlTypeId, name, index, searchDescendants,
                false, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("ListItemAt method - list item element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ListItemAt method - list item element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("ListItemAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("ListItemAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_ListItem listItem = new UIDA_ListItem(returnElement);
            return listItem;
        }

        /// <summary>
        /// Returns a collection of ListItems that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of ListItem elements, use null to return all ListItems</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_ListItem elements</returns>
        public UIDA_ListItem[] ListItems(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<IUIAutomationElement> allListItems = FindAll(UIA_ControlTypeIds.UIA_ListItemControlTypeId,
                name, searchDescendants, false, caseSensitive);

            List<UIDA_ListItem> listitems = new List<UIDA_ListItem>();
            if (allListItems != null)
            {
                foreach (IUIAutomationElement crtEl in allListItems)
                {
                    listitems.Add(new UIDA_ListItem(crtEl));
                }
            }
            return listitems.ToArray();
        }

        #endregion
    }
}
