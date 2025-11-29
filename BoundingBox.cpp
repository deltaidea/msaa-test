#include "BoundingBox.h"

std::string BoundingBox::ToString() const
{
    return "(" + std::to_string(left) + "," + std::to_string(top) + ", size " +
        std::to_string(width) + "," + std::to_string(height) + ")";
}

BoundingBox::BoundingBox(IAccessible* pAccessible)
{
	left = 0;
	top = 0;
	width = 0;
	height = 0;

    if (!pAccessible) return;

    VARIANT varChild{};
    varChild.vt = VT_I4;
    varChild.lVal = CHILDID_SELF;
    LONG _left, _top, _width, _height;

    try {
        if (SUCCEEDED(pAccessible->accLocation(&_left, &_top, &_width, &_height, varChild)))
        {
            left = _left;
            top = _top;
            width = _width;
            height = _height;
        }
    }
    catch (...) {}
}
