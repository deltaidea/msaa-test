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

static std::list<IAccessible*> GetAccessibleChildren(IAccessible* pAccessible)
{
    std::list<IAccessible*> children;
    if (!pAccessible) return children;
    long childCount = 0;

    if (FAILED(pAccessible->get_accChildCount(&childCount)) || childCount <= 0)
        return children;

    VARIANT* varChildren = new VARIANT[childCount];
    long obtained = 0;
    if (SUCCEEDED(AccessibleChildren(pAccessible, 0, childCount, varChildren, &obtained)))
    {
        for (long i = 0; i < min(childCount, obtained); i += 1)
        {
            if (varChildren[i].vt == VT_DISPATCH && varChildren[i].pdispVal != nullptr)
            {
                IAccessible* pChildAccessible = nullptr;
                if (SUCCEEDED(varChildren[i].pdispVal->QueryInterface(IID_IAccessible, (void**)&pChildAccessible)))
                {
                    children.push_back(pChildAccessible);
                }
            }
        }
    }
    for (long i = 0; i < min(childCount, obtained); i += 1)
    {
        VariantClear(&varChildren[i]);
    }
    delete[] varChildren;
    return children;
}

static void PrintAccessibleInfo(const AccessibleInfo& info, int indent = 0)
{
	std::cout << std::string(indent, ' ') << info.ToString() << std::endl;
}

static std::string GetWindowTitle(HWND hwnd)
{
    const int titleSize = 512;
    char titleBuffer[titleSize];
    if (GetWindowTextA(hwnd, titleBuffer, titleSize) > 0)
    {
		return std::string(titleBuffer);
    }
    return "<unknown>";
}

void InspectAccessibleTree()
{
    POINT cursorPoint;
    if (!GetCursorPos(&cursorPoint)) return;

    HWND hwnd = WindowFromPoint(cursorPoint);
    IAccessible* pAccessible = nullptr;
	std::cout << "Inspecting window: " << GetWindowTitle(hwnd) << std::endl;

    if (FAILED(AccessibleObjectFromWindow(hwnd, OBJID_CLIENT, IID_IAccessible, (void**)&pAccessible))) return;
    if (!pAccessible) return;

    AccessibleInfo info = AccessibleInfo(pAccessible);
    std::cout << "Initial object: " << info.ToString() << std::endl;

    IAccessible* mostSpecificAccessible = pAccessible;

    while (true)
    {
        bool foundMoreSpecific = false;

        std::list<IAccessible*> children = GetAccessibleChildren(mostSpecificAccessible);

        for (IAccessible* child : children)
        {
            AccessibleInfo childInfo = AccessibleInfo(child);
			if (!childInfo.isOffscreen() &&
                cursorPoint.x >= childInfo.box.left &&
                cursorPoint.x <= childInfo.box.left + childInfo.box.width &&
                cursorPoint.y >= childInfo.box.top &&
                cursorPoint.y <= childInfo.box.top + childInfo.box.height)
            {
                mostSpecificAccessible = child;
                std::cout << "Found child: " << childInfo.ToString() << std::endl;
                foundMoreSpecific = true;
            }
        }

        if (!foundMoreSpecific) break;
    }
}

int main()
{
    while (true)
    {
        std::cout << "With CoInitialize:" << std::endl;
        HRESULT _ = CoInitialize(nullptr);
        InspectAccessibleTree();

        std::cout << std::endl << "Without CoInitialize:" << std::endl;
        CoUninitialize();
        InspectAccessibleTree();
		std::cout << "----------------------------------------" << std::endl;

        Sleep(1000);
    }
}
