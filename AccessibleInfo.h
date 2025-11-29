#pragma once
#include "BoundingBox.h"

class AccessibleInfo
{
public:
    std::string name;
    std::string val;
    std::string description;
    std::string role;
    std::string state;
    BoundingBox box;

    AccessibleInfo() : name(""), val(""), description(""), role(""), state(""), box() {};
    std::string ToString() const;

    static std::string SerializeAccessibleString(BSTR bstr);
    static std::string SerializeRole(VARIANT varRole);
    static std::string SerializeState(VARIANT varState);
};
