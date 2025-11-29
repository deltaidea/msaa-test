#include "BoundingBox.h"

std::string BoundingBox::ToString() const
{
    return "(" + std::to_string(left) + "," + std::to_string(top) + "," +
        std::to_string(width) + "," + std::to_string(height) + ")";
}
