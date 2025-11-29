#include <windows.h>
#include <comdef.h> // For _bstr_t
#include <oleacc.h>
#include <numeric>
#include <string>
#include <list>

#include "AccessibleInfo.h"
#include "BoundingBox.h"

AccessibleInfo::AccessibleInfo(IAccessible* pAccessible)
{
    if (!pAccessible) return;

    VARIANT varChild{};
    varChild.vt = VT_I4;
    varChild.lVal = CHILDID_SELF;

    BSTR bstrName = nullptr;
    if (SUCCEEDED(pAccessible->get_accName(varChild, &bstrName)))
    {
        name = AccessibleInfo::SerializeAccessibleString(bstrName);
        SysFreeString(bstrName);
    }
    if (name.size() == 0) {
        name = "<unknown>";
    }

	BSTR bstrValue = nullptr;
    if (SUCCEEDED(pAccessible->get_accValue(varChild, &bstrValue)))
    {
        val = AccessibleInfo::SerializeAccessibleString(bstrValue);
        SysFreeString(bstrValue);
	}
    if (val.size() == 0) {
		val = "<unknown>";
    }

    VARIANT varRole{};
    if (SUCCEEDED(pAccessible->get_accRole(varChild, &varRole)))
    {
        role = AccessibleInfo::SerializeRole(varRole);
    }

    VARIANT varState{};
    if (SUCCEEDED(pAccessible->get_accState(varChild, &varState)))
    {
        state = AccessibleInfo::SerializeState(varState);
    }

    box = BoundingBox(pAccessible);
}

std::string AccessibleInfo::ToString() const
{
    return role + ": " + name + " - " + val + " | " + "State: " + state + " | " + "Bounding Box: " + box.ToString();
}

std::string AccessibleInfo::SerializeAccessibleString(BSTR bstr)
{
    if (!bstr) return "";

    _bstr_t result(bstr, false);
    return std::string(result);
}

std::string AccessibleInfo::SerializeRole(VARIANT varRole)
{
    if (varRole.vt == VT_I4)
    {
        switch (varRole.lVal)
        {
        case ROLE_SYSTEM_TITLEBAR: return "Title bar";
        case ROLE_SYSTEM_MENUBAR: return "Menu bar";
        case ROLE_SYSTEM_SCROLLBAR: return "Scrollbar";
        case ROLE_SYSTEM_GRIP: return "Grip";
        case ROLE_SYSTEM_SOUND: return "Sound";
        case ROLE_SYSTEM_CURSOR: return "Cursor";
        case ROLE_SYSTEM_CARET: return "Caret";
        case ROLE_SYSTEM_ALERT: return "Alert";
        case ROLE_SYSTEM_WINDOW: return "Window";
        case ROLE_SYSTEM_CLIENT: return "Client";
        case ROLE_SYSTEM_MENUPOPUP: return "Menu popup";
        case ROLE_SYSTEM_MENUITEM: return "Menu item";
        case ROLE_SYSTEM_TOOLTIP: return "Tooltip";
        case ROLE_SYSTEM_APPLICATION: return "Application";
        case ROLE_SYSTEM_DOCUMENT: return "Document";
        case ROLE_SYSTEM_PANE: return "Pane";
        case ROLE_SYSTEM_CHART: return "Chart";
        case ROLE_SYSTEM_DIALOG: return "Dialog";
        case ROLE_SYSTEM_BORDER: return "Border";
        case ROLE_SYSTEM_GROUPING: return "Grouping";
        case ROLE_SYSTEM_SEPARATOR: return "Separator";
        case ROLE_SYSTEM_TOOLBAR: return "Toolbar";
        case ROLE_SYSTEM_STATUSBAR: return "Status bar";
        case ROLE_SYSTEM_TABLE: return "Table";
        case ROLE_SYSTEM_COLUMNHEADER: return "Column header";
        case ROLE_SYSTEM_ROWHEADER: return "Row header";
        case ROLE_SYSTEM_COLUMN: return "Column";
        case ROLE_SYSTEM_ROW: return "Row";
        case ROLE_SYSTEM_CELL: return "Cell";
        case ROLE_SYSTEM_LINK: return "Link";
        case ROLE_SYSTEM_HELPBALLOON: return "Help balloon";
        case ROLE_SYSTEM_CHARACTER: return "Character";
        case ROLE_SYSTEM_LIST: return "List";
        case ROLE_SYSTEM_LISTITEM: return "List item";
        case ROLE_SYSTEM_OUTLINE: return "Outline";
        case ROLE_SYSTEM_OUTLINEITEM: return "Outline item";
        case ROLE_SYSTEM_PAGETAB: return "Page tab";
        case ROLE_SYSTEM_PROPERTYPAGE: return "Property page";
        case ROLE_SYSTEM_INDICATOR: return "Indicator";
        case ROLE_SYSTEM_GRAPHIC: return "Graphic";
        case ROLE_SYSTEM_STATICTEXT: return "Static text";
        case ROLE_SYSTEM_TEXT: return "Text";
        case ROLE_SYSTEM_PUSHBUTTON: return "Push button";
        case ROLE_SYSTEM_CHECKBUTTON: return "Check button";
        case ROLE_SYSTEM_RADIOBUTTON: return "Radio button";
        case ROLE_SYSTEM_COMBOBOX: return "Combobox";
        case ROLE_SYSTEM_DROPLIST: return "Drop list";
        case ROLE_SYSTEM_PROGRESSBAR: return "Progress bar";
        case ROLE_SYSTEM_DIAL: return "Dial";
        case ROLE_SYSTEM_HOTKEYFIELD: return "Hotkey field";
        case ROLE_SYSTEM_SLIDER: return "Slider";
        case ROLE_SYSTEM_SPINBUTTON: return "Spin button";
        case ROLE_SYSTEM_DIAGRAM: return "Diagram";
        case ROLE_SYSTEM_ANIMATION: return "Animation";
        case ROLE_SYSTEM_EQUATION: return "Equation";
        case ROLE_SYSTEM_BUTTONDROPDOWN: return "Button dropdown";
        case ROLE_SYSTEM_BUTTONMENU: return "Button menu";
        case ROLE_SYSTEM_BUTTONDROPDOWNGRID: return "Button dropdown grid";
        case ROLE_SYSTEM_WHITESPACE: return "Whitespace";
        case ROLE_SYSTEM_PAGETABLIST: return "Page tab list";
        case ROLE_SYSTEM_CLOCK: return "Clock";
        case ROLE_SYSTEM_SPLITBUTTON: return "Split button";
        case ROLE_SYSTEM_IPADDRESS: return "IP address";
        case ROLE_SYSTEM_OUTLINEBUTTON: return "Outline button";
        default: return "Unknown Role";
        }
    }
    return "Unknown Role Type";
}

std::string AccessibleInfo::SerializeState(VARIANT varState)
{
    std::list<std::string> states;
    if (varState.vt == VT_I4)
    {
        LONG state = varState.lVal;
        if (state & STATE_SYSTEM_UNAVAILABLE) states.push_back("Unavailable");
        if (state & STATE_SYSTEM_SELECTED) states.push_back("Selected");
        if (state & STATE_SYSTEM_FOCUSED) states.push_back("Focused");
        if (state & STATE_SYSTEM_PRESSED) states.push_back("Pressed");
        if (state & STATE_SYSTEM_CHECKED) states.push_back("Checked");
        if (state & STATE_SYSTEM_MIXED) states.push_back("Mixed");
        if (state & STATE_SYSTEM_READONLY) states.push_back("Readonly");
        if (state & STATE_SYSTEM_HOTTRACKED) states.push_back("Hot tracked");
        if (state & STATE_SYSTEM_DEFAULT) states.push_back("Default");
        if (state & STATE_SYSTEM_EXPANDED) states.push_back("Expanded");
        if (state & STATE_SYSTEM_COLLAPSED) states.push_back("Collapsed");
        if (state & STATE_SYSTEM_BUSY) states.push_back("Busy");
        if (state & STATE_SYSTEM_FLOATING) states.push_back("Floating");
        if (state & STATE_SYSTEM_MARQUEED) states.push_back("Marqueed");
        if (state & STATE_SYSTEM_ANIMATED) states.push_back("Animated");
        if (state & STATE_SYSTEM_INVISIBLE) states.push_back("Invisible");
        if (state & STATE_SYSTEM_OFFSCREEN) states.push_back("Offscreen");
        if (state & STATE_SYSTEM_SIZEABLE) states.push_back("Sizeable");
        if (state & STATE_SYSTEM_MOVEABLE) states.push_back("Moveable");
        if (state & STATE_SYSTEM_SELFVOICING) states.push_back("Self voicing");
        if (state & STATE_SYSTEM_FOCUSABLE) states.push_back("Focusable");
        if (state & STATE_SYSTEM_SELECTABLE) states.push_back("Selectable");
        if (state & STATE_SYSTEM_LINKED) states.push_back("Linked");
        if (state & STATE_SYSTEM_TRAVERSED) states.push_back("Traversed");
        if (state & STATE_SYSTEM_MULTISELECTABLE) states.push_back("Multi selectable");
        if (state & STATE_SYSTEM_EXTSELECTABLE) states.push_back("Ext selectable");
        if (state & STATE_SYSTEM_ALERT_LOW) states.push_back("Alert Low");
        if (state & STATE_SYSTEM_ALERT_MEDIUM) states.push_back("Alert Medium");
        if (state & STATE_SYSTEM_ALERT_HIGH) states.push_back("Alert High");
        if (state & STATE_SYSTEM_PROTECTED) states.push_back("Protected");
    }
    return states.empty() ? "None" : std::accumulate(std::next(states.begin()), states.end(), states.front(),
        [](const std::string& a, const std::string& b) { return a + ", " + b; });
}
