#include <iostream>
#include <string>
#include <vector>
#include <fstream>
#include <sstream>
#include <string>
#include <vector>
#include <filesystem> 


std::string ReplaceSubstring(std::string source, const std::string& toFind, const std::string& toReplace) {
    size_t pos = 0;
    while ((pos = source.find(toFind, pos)) != std::string::npos) {
        source.replace(pos, toFind.length(), toReplace);
        pos += toReplace.length();
    }
    return source;
}


std::vector<std::string> ReadFileAndSplitLines(const std::string& filePath) {
    // ����ļ��Ƿ����
    if (!std::filesystem::exists(filePath)) {
        return {}; // ����ļ������ڣ����ؿյ�vector
    }

    std::ifstream file(filePath);
    if (!file.is_open()) {
        return {}; // ����ļ��޷��򿪣����ؿյ�vector
    }

    std::stringstream buffer;
    buffer << file.rdbuf(); // ��ȡ�ļ����ַ�����

    std::string content = buffer.str();

    // �滻���� \r\n Ϊ \n
    size_t pos = 0;
    while ((pos = content.find("\r\n", pos)) != std::string::npos) {
        content.replace(pos, 2, "\n");
        pos += 1;
    }

    // ���зָ�
    std::vector<std::string> lines;
    std::istringstream stream(content);
    std::string line;
    while (std::getline(stream, line)) {
        lines.push_back(line);
    }

    return lines;
}

/// <summary>
/// ���ݷ��ŷָ�Ϊ���룬���û���ţ���Ϊһ��ԭ�ı�
/// </summary>
/// <param name="input"></param>
/// <param name="delimiter"></param>
/// <returns></returns>
std::vector<std::string> SplitStringTwo(const std::string& input, char delimiter) {
    std::vector<std::string> result;
    size_t pos = input.find(delimiter);

    if (pos != std::string::npos) {
        // ����ҵ��˷ָ���
        std::string first = input.substr(0, pos);
        std::string second = input.substr(pos + 1);

        if (!second.empty()) {
            // ����ڶ����ַǿգ����������
            result.push_back(first);
            result.push_back(second);
            return result;
        }
    }

    // �������������ԭ�ַ���
    result.push_back(input);
    return result;
}