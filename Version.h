#pragma once
// Version.h
#ifndef VERSION_H
#define VERSION_H

#include <iostream>
#include <string>

class Version {
public:
    // Ĭ�Ϲ��캯��
    Version() = default;
    // ����������
    Version(int major, int minor, int build = -1, int revision = -1);
    // ���ַ�������
    Version(const std::string& versionStr);

    // �Ƚϲ�����
    bool operator==(const Version& other) const;
    bool operator!=(const Version& other) const;
    bool operator<(const Version& other) const;
    bool operator>(const Version& other) const;
    bool operator<=(const Version& other) const;
    bool operator>=(const Version& other) const;

    // ����汾��
    friend std::ostream& operator<<(std::ostream& os, const Version& v);

private:
    int major_ = 0;
    int minor_ = 0;
    int build_ = -1;
    int revision_ = -1;

    void parseVersionString(const std::string& versionStr);
};

#endif // VERSION_H