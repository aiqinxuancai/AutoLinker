#include "PathHelper.h"


std::string GetBasePath() {
    char buffer[MAX_PATH];
    GetModuleFileName(NULL, buffer, MAX_PATH);
    std::string::size_type pos = std::string(buffer).find_last_of("\\/");
    return std::string(buffer).substr(0, pos);
}

std::string ExtractBetweenDashes(const std::string& text) {
    std::string delimiter = " - ";

    // �ҵ���һ��" - "��λ��
    size_t start = text.find(delimiter);
    if (start == std::string::npos) {
        // û���ҵ�" - "�����ؿ��ַ���
        return "";
    }
    start += delimiter.length(); // ����" - "����ʼ�ڵ�һ��" - "֮����ַ�

    // ��startλ�ÿ�ʼ���ҵ���һ��" - "��λ��
    size_t end = text.find(delimiter, start);
    if (end == std::string::npos) {
        // û���ҵ��ڶ���" - "�����ؿ��ַ���
        return "";
    }

    // ȡ������" - "֮������ַ���
    return text.substr(start, end - start);
}