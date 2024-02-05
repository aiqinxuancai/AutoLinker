#pragma once
// Version.h
#ifndef VERSION_H
#define VERSION_H

#include <iostream>
#include <string>

class Version {
public:
    // 默认构造函数
    Version() = default;
    // 从整数构造
    Version(int major, int minor, int build = -1, int revision = -1);
    // 从字符串构造
    Version(const std::string& versionStr);

    // 比较操作符
    bool operator==(const Version& other) const;
    bool operator!=(const Version& other) const;
    bool operator<(const Version& other) const;
    bool operator>(const Version& other) const;
    bool operator<=(const Version& other) const;
    bool operator>=(const Version& other) const;

    // 输出版本号
    friend std::ostream& operator<<(std::ostream& os, const Version& v);

private:
    int major_ = 0;
    int minor_ = 0;
    int build_ = -1;
    int revision_ = -1;

    void parseVersionString(const std::string& versionStr);
};

#endif // VERSION_H