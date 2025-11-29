#pragma once

#include <windows.h>
#include <string>

class BoundingBox
{
public:
    LONG left;
    LONG top;
    LONG width;
    LONG height;

    BoundingBox() : left(0), top(0), width(0), height(0) {}
	BoundingBox(LONG l, LONG t, LONG w, LONG h) : left(l), top(t), width(w), height(h) {}

    std::string ToString() const;
};
