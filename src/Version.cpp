// Version.cpp
#include "Version.h"
#include <sstream>
#include <vector>

Version::Version(int major, int minor, int build, int revision)
    : major_(major), minor_(minor), build_(build), revision_(revision) {}

Version::Version(const std::string& versionStr) {
    parseVersionString(versionStr);
}

bool Version::operator==(const Version& other) const {
    return std::tie(major_, minor_, build_, revision_) ==
        std::tie(other.major_, other.minor_, other.build_, other.revision_);
}

bool Version::operator!=(const Version& other) const {
    return !(*this == other);
}

bool Version::operator<(const Version& other) const {
    return std::tie(major_, minor_, build_, revision_) <
        std::tie(other.major_, other.minor_, other.build_, other.revision_);
}

bool Version::operator>(const Version& other) const {
    return other < *this;
}

bool Version::operator<=(const Version& other) const {
    return !(other < *this);
}

bool Version::operator>=(const Version& other) const {
    return !(*this < other);
}

std::ostream& operator<<(std::ostream& os, const Version& v) {
    os << v.major_ << "." << v.minor_;
    if (v.build_ != -1) {
        os << "." << v.build_;
        if (v.revision_ != -1) {
            os << "." << v.revision_;
        }
    }
    return os;
}

void Version::parseVersionString(const std::string& versionStr) {
    std::istringstream versionStream(versionStr);
    std::string segment;
    std::vector<int> segments;

    while (std::getline(versionStream, segment, '.')) {
        segments.push_back(std::stoi(segment));
    }

    if (segments.size() > 0) major_ = segments[0];
    if (segments.size() > 1) minor_ = segments[1];
    if (segments.size() > 2) build_ = segments[2];
    if (segments.size() > 3) revision_ = segments[3];
}
