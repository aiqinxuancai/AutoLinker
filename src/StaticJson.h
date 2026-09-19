#pragma once

#include "..\thirdparty\json.hpp"

// 固定 JSON 文本在首次使用时解析并缓存；返回只读引用，赋给 json 变量时仍复制独立数据。
// 每个调用位置拥有独立的函数局部静态值，初始化由 C++ 保证线程安全，仅在执行到调用点时解析。
// 仅用于编译期固定、合法的 JSON 文本；中文使用 JSON Unicode 转义，避免受执行字符集影响。
#define AUTOLINKER_STATIC_JSON(text) \
    ([]() -> const nlohmann::json& { \
        static const nlohmann::json value = nlohmann::json::parse(text); \
        return value; \
    }())
