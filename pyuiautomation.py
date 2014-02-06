from xml.dom import minidom

import comtypes
import comtypes.client

UIAutomationClient = comtypes.client.GetModule('UIAutomationCore.dll')

# noinspection PyProtectedMember
_IUIAutomation = comtypes.CoCreateInstance(UIAutomationClient.CUIAutomation._reg_clsid_,
                                           interface=UIAutomationClient.IUIAutomation,
                                           clsctx=comtypes.CLSCTX_INPROC_SERVER)

_control_type = {
    UIAutomationClient.UIA_ButtonControlTypeId: 'UIA_ButtonControlTypeId',
    UIAutomationClient.UIA_CalendarControlTypeId: 'UIA_CalendarControlTypeId',
    UIAutomationClient.UIA_CheckBoxControlTypeId: 'UIA_CheckBoxControlTypeId',
    UIAutomationClient.UIA_ComboBoxControlTypeId: 'UIA_ComboBoxControlTypeId',
    UIAutomationClient.UIA_CustomControlTypeId: 'UIA_CustomControlTypeId',
    UIAutomationClient.UIA_DataGridControlTypeId: 'UIA_DataGridControlTypeId',
    UIAutomationClient.UIA_DataItemControlTypeId: 'UIA_DataItemControlTypeId',
    UIAutomationClient.UIA_DocumentControlTypeId: 'UIA_DocumentControlTypeId',
    UIAutomationClient.UIA_EditControlTypeId: 'UIA_EditControlTypeId',
    UIAutomationClient.UIA_GroupControlTypeId: 'UIA_GroupControlTypeId',
    UIAutomationClient.UIA_HeaderControlTypeId: 'UIA_HeaderControlTypeId',
    UIAutomationClient.UIA_HeaderItemControlTypeId: 'UIA_HeaderItemControlTypeId',
    UIAutomationClient.UIA_HyperlinkControlTypeId: 'UIA_HyperlinkControlTypeId',
    UIAutomationClient.UIA_ImageControlTypeId: 'UIA_ImageControlTypeId',
    UIAutomationClient.UIA_ListControlTypeId: 'UIA_ListControlTypeId',
    UIAutomationClient.UIA_ListItemControlTypeId: 'UIA_ListItemControlTypeId',
    UIAutomationClient.UIA_MenuBarControlTypeId: 'UIA_MenuBarControlTypeId',
    UIAutomationClient.UIA_MenuControlTypeId: 'UIA_MenuControlTypeId',
    UIAutomationClient.UIA_MenuItemControlTypeId: 'UIA_MenuItemControlTypeId',
    UIAutomationClient.UIA_PaneControlTypeId: 'UIA_PaneControlTypeId',
    UIAutomationClient.UIA_ProgressBarControlTypeId: 'UIA_ProgressBarControlTypeId',
    UIAutomationClient.UIA_RadioButtonControlTypeId: 'UIA_RadioButtonControlTypeId',
    UIAutomationClient.UIA_ScrollBarControlTypeId: 'UIA_ScrollBarControlTypeId',
    UIAutomationClient.UIA_SeparatorControlTypeId: 'UIA_SeparatorControlTypeId',
    UIAutomationClient.UIA_SliderControlTypeId: 'UIA_SliderControlTypeId',
    UIAutomationClient.UIA_SpinnerControlTypeId: 'UIA_SpinnerControlTypeId',
    UIAutomationClient.UIA_SplitButtonControlTypeId: 'UIA_SplitButtonControlTypeId',
    UIAutomationClient.UIA_StatusBarControlTypeId: 'UIA_StatusBarControlTypeId',
    UIAutomationClient.UIA_TabControlTypeId: 'UIA_TabControlTypeId',
    UIAutomationClient.UIA_TabItemControlTypeId: 'UIA_TabItemControlTypeId',
    UIAutomationClient.UIA_TableControlTypeId: 'UIA_TableControlTypeId',
    UIAutomationClient.UIA_TextControlTypeId: 'UIA_TextControlTypeId',
    UIAutomationClient.UIA_ThumbControlTypeId: 'UIA_ThumbControlTypeId',
    UIAutomationClient.UIA_TitleBarControlTypeId: 'UIA_TitleBarControlTypeId',
    UIAutomationClient.UIA_ToolBarControlTypeId: 'UIA_ToolBarControlTypeId',
    UIAutomationClient.UIA_ToolTipControlTypeId: 'UIA_ToolTipControlTypeId',
    UIAutomationClient.UIA_TreeControlTypeId: 'UIA_TreeControlTypeId',
    UIAutomationClient.UIA_TreeItemControlTypeId: 'UIA_TreeItemControlTypeId',
    UIAutomationClient.UIA_WindowControlTypeId: 'UIA_WindowControlTypeId'
}

_tree_scope = {
    'ancestors': UIAutomationClient.TreeScope_Ancestors,
    'children': UIAutomationClient.TreeScope_Children,
    'descendants': UIAutomationClient.TreeScope_Descendants,
    'element': UIAutomationClient.TreeScope_Element,
    'parent': UIAutomationClient.TreeScope_Parent,
    'subtree': UIAutomationClient.TreeScope_Subtree
}


class _UIAutomationElement(object):
    def __init__(self, IUIAutomationElement):
        self.IUIAutomationElement = IUIAutomationElement

    @property
    def CurrentAutomationId(self):
        """Retrieves the UI Automation identifier of the element

        :rtype : unicode
        """
        return self.IUIAutomationElement.CurrentAutomationId
    AutomationId = CurrentAutomationId

    @property
    def CurrentBoundingRectangle(self):
        """Retrieves the coordinates of the rectangle that completely encloses the element.

        Returns tuple (left, top, right, bottom)

        :rtype : tuple
        """
        rect = self.IUIAutomationElement.CurrentBoundingRectangle
        return rect.left, rect.top, rect.right, rect.bottom
    BoundingRectangle = CurrentBoundingRectangle

    @property
    def CurrentClassName(self):
        """Retrieves the class name of the element

        :rtype : unicode
        """
        return self.IUIAutomationElement.CurrentClassName
    ClassName = CurrentClassName

    @property
    def CurrentControlType(self):
        """Retrieves the control type of the element

        :rtype : int
        """
        return self.IUIAutomationElement.CurrentControlType
    ControlType = CurrentControlType

    @property
    def CurrentControlTypeName(self):
        """Retrieves the name of the control type of the element.

        Ref: http://msdn.microsoft.com/en-us/library/windows/desktop/ee671198(v=vs.85).aspx

        :rtype : str
        """
        return _control_type[self.CurrentControlType]
    ControlTypeName = CurrentControlTypeName

    @property
    def CurrentIsEnabled(self):
        """Indicates whether the element is enabled

        :rtype : bool
        """
        return bool(self.IUIAutomationElement.CurrentIsEnabled)
    IsEnabled = CurrentIsEnabled

    @property
    def CurrentName(self):
        """Retrieves the name of the element

        :rtype : unicode
        """
        return self.IUIAutomationElement.CurrentName
    Name = CurrentName

    def GetClickablePoint(self):
        """Retrieves a point on the element that can be clicked

        Returns tuple (x, y) if a clickable point was retrieved, or None otherwise

        :rtype : tuple
        """
        point = self.IUIAutomationElement.GetClickablePoint()
        if point[1]:
            return point[0].x, point[0].y
        else:
            return None

    def _build_condition(self, Name, ControlType, AutomationId):
        condition = _IUIAutomation.CreateTrueCondition()

        if Name is not None:
            name_condition = _IUIAutomation.CreatePropertyCondition(UIAutomationClient.UIA_NamePropertyId, Name)
            condition = _IUIAutomation.CreateAndCondition(condition, name_condition)

        if ControlType is not None:
            control_type_condition = _IUIAutomation.CreatePropertyCondition(
                UIAutomationClient.UIA_ControlTypePropertyId, ControlType)
            condition = _IUIAutomation.CreateAndCondition(condition, control_type_condition)

        if AutomationId is not None:
            automation_id_condition = _IUIAutomation.CreatePropertyCondition(
                UIAutomationClient.UIA_AutomationIdPropertyId, AutomationId)
            condition = _IUIAutomation.CreateAndCondition(condition, automation_id_condition)

        return condition

    def findfirst(self, tree_scope, Name=None, ControlType=None, AutomationId=None):
        """Retrieves the first child or descendant element that matches specified conditions

        Returns None if there is no element that matches specified conditions
        If Name is None, element with any name will match
        If ControlType is None, element with any control type will match

        :param tree_scope: Should be one of 'element', 'children', 'descendants', 'parent', 'ancestors', 'subtree'.
            Ref: http://msdn.microsoft.com/en-us/library/windows/desktop/ee671699(v=vs.85).aspx
        :type tree_scope: str
        :param Name: Name of the element.
        :type Name: str
        :param ControlType: Control type of the element (one of UIAutomationClient.UIA_*ControlTypeId).
            Ref: http://msdn.microsoft.com/en-us/library/windows/desktop/ee671198(v=vs.85).aspx
        :type ControlType: int
        :param AutomationId: UI Automation identifier (ID) for the automation element.
            Ref: http://msdn.microsoft.com/en-us/library/windows/desktop/ee695997(v=vs.85).aspx
        :type AutomationId: str

        :rtype : _UIAutomationElement
        """
        tree_scope = _tree_scope[tree_scope]
        condition = self._build_condition(Name, ControlType, AutomationId)
        element = self.IUIAutomationElement.FindFirst(tree_scope, condition)
        return _UIAutomationElement(element) if element else None

    def findall(self, tree_scope, Name=None, ControlType=None, AutomationId=None):
        """Returns list of UI Automation elements that satisfy specified conditions

        Returns empty list if there are no elements that matches specified conditions
        If Name is None, elements with any name will match
        If ControlType is None, elements with any control type will match

        :param tree_scope: Should be one of 'element', 'children', 'descendants', 'parent', 'ancestors', 'subtree'.
            Ref: http://msdn.microsoft.com/en-us/library/windows/desktop/ee671699(v=vs.85).aspx
        :type tree_scope: str
        :param Name: Name of the element.
        :type Name: str
        :param ControlType: Control type of the element (one of UIAutomationClient.UIA_*ControlTypeId).
            Ref: http://msdn.microsoft.com/en-us/library/windows/desktop/ee671198(v=vs.85).aspx
        :type ControlType: int
        :param AutomationId: UI Automation identifier (ID) for the automation element.
            Ref: http://msdn.microsoft.com/en-us/library/windows/desktop/ee695997(v=vs.85).aspx
        :type AutomationId: str

        :rtype : list
        """
        tree_scope = _tree_scope[tree_scope]
        condition = self._build_condition(Name, ControlType, AutomationId)

        IUIAutomationElementArray = self.IUIAutomationElement.FindAll(tree_scope, condition)
        return [_UIAutomationElement(IUIAutomationElementArray.GetElement(i)) for i in
                xrange(IUIAutomationElementArray.Length)]

    def Invoke(self):
        IUnknown = self.IUIAutomationElement.GetCurrentPattern(UIAutomationClient.UIA_InvokePatternId)
        IUIAutomationInvokePattern = IUnknown.QueryInterface(UIAutomationClient.IUIAutomationInvokePattern)
        IUIAutomationInvokePattern.Invoke()

    def __str__(self):
        return '<%s (Name: %s, Class: %s, AutomationId: %s>' % (
            self.CurrentControlTypeName, self.CurrentName, self.CurrentClassName, self.CurrentAutomationId)

    def toxml(self):
        xml = minidom.Document()
        queue = [(self, xml)]
        while queue:
            element, xml_node = queue.pop(0)
            xml_element = minidom.Element(element.CurrentControlTypeName)
            xml_element.attributes['Name'] = str(element.CurrentName)
            xml_element.attributes['AutomationId'] = str(element.CurrentAutomationId)
            xml_element.attributes['ClassName'] = str(element.CurrentClassName)
            xml_element.ownerDocument = xml
            xml_node.appendChild(xml_element)
            for child in element.findall('children'):
                queue.append((child, xml_element))
        return xml.toprettyxml()


def ElementFromHandle(handle):
    return _UIAutomationElement(_IUIAutomation.ElementFromHandle(handle))


def GetRootElement():
    return _UIAutomationElement(_IUIAutomation.GetRootElement())
