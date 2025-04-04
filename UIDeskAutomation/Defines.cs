
namespace UIDeskAutomationLib
{
	internal abstract class UIA_ControlTypeIds
	{
		internal const int UIA_ButtonControlTypeId = 50000;
		internal const int UIA_CalendarControlTypeId = 50001;
		internal const int UIA_CheckBoxControlTypeId = 50002;
		internal const int UIA_ComboBoxControlTypeId = 50003;
		internal const int UIA_EditControlTypeId = 50004;
		internal const int UIA_HyperlinkControlTypeId = 50005;
		internal const int UIA_ImageControlTypeId = 50006;
		internal const int UIA_ListItemControlTypeId = 50007;
		internal const int UIA_ListControlTypeId = 50008;
		internal const int UIA_MenuControlTypeId = 50009;
		internal const int UIA_MenuBarControlTypeId = 50010;
		internal const int UIA_MenuItemControlTypeId = 50011;
		internal const int UIA_ProgressBarControlTypeId = 50012;
		internal const int UIA_RadioButtonControlTypeId = 50013;
		internal const int UIA_ScrollBarControlTypeId = 50014;
		internal const int UIA_SliderControlTypeId = 50015;
		internal const int UIA_SpinnerControlTypeId = 50016;
		internal const int UIA_StatusBarControlTypeId = 50017;
		internal const int UIA_TabControlTypeId = 50018;
		internal const int UIA_TabItemControlTypeId = 50019;
		internal const int UIA_TextControlTypeId = 50020;
		internal const int UIA_ToolBarControlTypeId = 50021;
		internal const int UIA_ToolTipControlTypeId = 50022;
		internal const int UIA_TreeControlTypeId = 50023;
		internal const int UIA_TreeItemControlTypeId = 50024;
		internal const int UIA_CustomControlTypeId = 50025;
		internal const int UIA_GroupControlTypeId = 50026;
		internal const int UIA_ThumbControlTypeId = 50027;
		internal const int UIA_DataGridControlTypeId = 50028;
		internal const int UIA_DataItemControlTypeId = 50029;
		internal const int UIA_DocumentControlTypeId = 50030;
		internal const int UIA_SplitButtonControlTypeId = 50031;
		internal const int UIA_WindowControlTypeId = 50032;
		internal const int UIA_PaneControlTypeId = 50033;
		internal const int UIA_HeaderControlTypeId = 50034;
		internal const int UIA_HeaderItemControlTypeId = 50035;
		internal const int UIA_TableControlTypeId = 50036;
		internal const int UIA_TitleBarControlTypeId = 50037;
		internal const int UIA_SeparatorControlTypeId = 50038;
	}
	
	internal abstract class UIA_PatternIds
	{
		internal const int UIA_InvokePatternId = 10000;
		internal const int UIA_SelectionPatternId = 10001;
		internal const int UIA_ValuePatternId = 10002;
		internal const int UIA_RangeValuePatternId = 10003;
		internal const int UIA_ScrollPatternId = 10004;
		internal const int UIA_ExpandCollapsePatternId = 10005;
		internal const int UIA_GridPatternId = 10006;
		internal const int UIA_GridItemPatternId = 10007;
		internal const int UIA_MultipleViewPatternId = 10008;
		internal const int UIA_WindowPatternId = 10009;
		internal const int UIA_SelectionItemPatternId = 10010;
		internal const int UIA_DockPatternId = 10011;
		internal const int UIA_TablePatternId = 10012;
		internal const int UIA_TableItemPatternId = 10013;
		internal const int UIA_TextPatternId = 10014;
		internal const int UIA_TogglePatternId = 10015;
		internal const int UIA_TransformPatternId = 10016;
		internal const int UIA_ScrollItemPatternId = 10017;
		internal const int UIA_LegacyIAccessiblePatternId = 10018;
		internal const int UIA_ItemContainerPatternId = 10019;
		internal const int UIA_VirtualizedItemPatternId = 10020;
		internal const int UIA_SynchronizedInputPatternId = 10021;
	}
	
	internal abstract class UIA_PropertyIds
	{
		internal const int UIA_ControlTypePropertyId = 30003;
		internal const int UIA_NamePropertyId = 30005;
		internal const int UIA_ClassNamePropertyId = 30012;
		internal const int UIA_ToggleToggleStatePropertyId = 30086;
		internal const int UIA_SelectionItemIsSelectedPropertyId = 30079;
		internal const int UIA_ValueValuePropertyId = 30045;
		internal const int UIA_RangeValueValuePropertyId = 30047;
		internal const int UIA_ExpandCollapseExpandCollapseStatePropertyId = 30070;
		internal const int UIA_WindowWindowVisualStatePropertyId = 30075;
		internal const int UIA_BoundingRectanglePropertyId = 30001;
		internal const int UIA_AccessKeyPropertyId = 30007;
		internal const int UIA_AcceleratorKeyPropertyId = 30006;
        internal const int UIA_AriaRolePropertyId = 30101;
        internal const int UIA_AutomationIdPropertyId = 30011;
        internal const int UIA_FrameworkIdPropertyId = 30024;
        internal const int UIA_FullDescriptionPropertyId = 30159;
        internal const int UIA_HelpTextPropertyId = 30013;
        internal const int UIA_ItemStatusPropertyId = 30026;
        internal const int UIA_ItemTypePropertyId = 300021;
        internal const int UIA_LocalizedLandmarkTypePropertyId = 30158;
        internal const int UIA_ProcessIdPropertyId = 30002;
        internal const int UIA_ProviderDescriptionPropertyId = 30107;
    }
	
	internal abstract class UIA_EventIds
	{
		internal const int UIA_Invoke_InvokedEventId = 20009;
		internal const int UIA_Window_WindowClosedEventId = 20017;
		internal const int UIA_Text_TextChangedEventId = 20015;
		internal const int UIA_Text_TextSelectionChangedEventId = 20014;
		internal const int UIA_SelectionItem_ElementSelectedEventId = 20012;
		internal const int UIA_SelectionItem_ElementAddedToSelectionEventId = 20010;
		internal const int UIA_SelectionItem_ElementRemovedFromSelectionEventId = 20011;
		internal const int UIA_MenuOpenedEventId = 20003;
	}
}