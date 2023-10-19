#include <iostream>
#include <string>



std::string replaceSubstring(std::string source, const std::string& toFind, const std::string& toReplace) {
    size_t pos = 0;
    while ((pos = source.find(toFind, pos)) != std::string::npos) {
        source.replace(pos, toFind.length(), toReplace);
        pos += toReplace.length();
    }
    return source;
}