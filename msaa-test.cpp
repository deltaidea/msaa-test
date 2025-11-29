#include <windows.h>
#include <comdef.h> // For _bstr_t
#include <iostream>
#include <list>
#include <oleacc.h>
#include <numeric>
#include <string>

#include "BoundingBox.h"
#include "AccessibleInfo.h"

#pragma comment(lib, "oleacc.lib")

// Function to get bounding box of an accessible object
static BoundingBox GetAccessibleBoundingBox(IAccessible* pAccessible)
{
	BoundingBox box = {};

    if (!pAccessible) return {};

    VARIANT varChild{};
    varChild.vt = VT_I4;
    varChild.lVal = CHILDID_SELF;
    LONG left, top, width, height;

    try {
        if (pAccessible->accLocation(&left, &top, &width, &height, varChild) == S_OK)
		    return BoundingBox{ left, top, width, height };
    }
    catch (...) {}

    return {};
}

static std::list<IAccessible*> GetAccessibleChildren(IAccessible* pAccessible)
{
    std::list<IAccessible*> children;
    if (!pAccessible) return children;
    long childCount = 0;

    if (pAccessible->get_accChildCount(&childCount) != S_OK || childCount <= 0)
        return children;

    VARIANT* varChildren = new VARIANT[childCount];
    long obtained = 0;
    if (AccessibleChildren(pAccessible, 0, childCount, varChildren, &obtained) == S_OK)
    {
        for (long i = 0; i < obtained; ++i)
        {
            if (varChildren[i].vt == VT_DISPATCH && varChildren[i].pdispVal)
            {
                IAccessible* pChildAccessible = nullptr;
                if (varChildren[i].pdispVal->QueryInterface(IID_IAccessible, (void**)&pChildAccessible) == S_OK)
                {
                    children.push_back(pChildAccessible);
                }
            }
        }
    }
    for (long i = 0; i < obtained; ++i)
    {
        VariantClear(&varChildren[i]);
    }
    delete[] varChildren;
    return children;
}

static AccessibleInfo BuildAccessibleInfo(IAccessible* pAccessible)
{
    AccessibleInfo info;
    if (!pAccessible) return info;

    VARIANT varChild{};
    varChild.vt = VT_I4;
    varChild.lVal = CHILDID_SELF;
    BSTR bstrName = nullptr;
    if (pAccessible->get_accName(varChild, &bstrName) == S_OK)
    {
        info.name = AccessibleInfo::SerializeAccessibleString(bstrName);
        SysFreeString(bstrName);
    }
    BSTR bstrValue = nullptr;
    if (pAccessible->get_accValue(varChild, &bstrValue) == S_OK)
    {
        info.val = AccessibleInfo::SerializeAccessibleString(bstrValue);
        SysFreeString(bstrValue);
    }
    BSTR bstrDescription = nullptr;
    if (pAccessible->get_accDescription(varChild, &bstrDescription) == S_OK)
    {
        info.description = AccessibleInfo::SerializeAccessibleString(bstrDescription);
        SysFreeString(bstrDescription);
    }
    VARIANT varRole{};
    if (pAccessible->get_accRole(varChild, &varRole) == S_OK)
    {
        info.role = AccessibleInfo::SerializeRole(varRole);
    }
    VARIANT varState{};
    if (pAccessible->get_accState(varChild, &varState) == S_OK)
    {
        info.state = AccessibleInfo::SerializeState(varState);
    }
    info.box = GetAccessibleBoundingBox(pAccessible);
    return info;
}

static void PrintAccessibleInfo(const AccessibleInfo& info, int indent = 0)
{
	std::cout << std::string(indent, ' ') << info.ToString() << std::endl;
}

int main()
{
    HRESULT hr = CoInitialize(nullptr);
    if (FAILED(hr))
    {
        std::cerr << "CoInitialize failed: 0x" << std::hex << hr << std::endl;
        return 1;
    }

    while (true)
    {
        Sleep(1000);

        POINT cursorPoint;
		if (!GetCursorPos(&cursorPoint)) continue;

        HWND hwnd = WindowFromPoint(cursorPoint);
        IAccessible* pAccessible = nullptr;

        if (FAILED(AccessibleObjectFromWindow(hwnd, OBJID_CLIENT, IID_IAccessible, (void**)&pAccessible))) continue;
		if (!pAccessible) continue;

        AccessibleInfo info = BuildAccessibleInfo(pAccessible);
        std::cout << "Top level: " << info.ToString() << std::endl;

		IAccessible* mostSpecificAccessible = pAccessible;

        while (true)
        {
            bool foundMoreSpecific = false;
            
            std::list<IAccessible*> children = GetAccessibleChildren(mostSpecificAccessible);

            for (IAccessible* child : children)
            {
                AccessibleInfo childInfo = BuildAccessibleInfo(child);
                if (cursorPoint.x >= childInfo.box.left &&
                    cursorPoint.x <= childInfo.box.left + childInfo.box.width &&
                    cursorPoint.y >= childInfo.box.top &&
                    cursorPoint.y <= childInfo.box.top + childInfo.box.height)
                {
					mostSpecificAccessible->Release();
                    mostSpecificAccessible = child;
                    std::cout << "Going deeper: " << childInfo.ToString() << std::endl;
                    foundMoreSpecific = true;
                }
            }

            for (IAccessible* child : children)
            {
                try {
				    if (child != mostSpecificAccessible) child->Release();
                } catch (...) {}
			}

            if (!foundMoreSpecific) break;
		}

        std::cout << "Accessible Object under cursor: " << BuildAccessibleInfo(mostSpecificAccessible).ToString() << std::endl;
		std::cout << "----------------------------------------" << std::endl;

        mostSpecificAccessible->Release();
    }

    CoUninitialize();
}
