#pragma once

// AutoLinkerTestApi.h
// 提供可脱离易语言 IDE 运行的轻量测试导出接口。

#ifdef ELIB_FNE_EXPORTS
#define AUTOLINKER_TEST_API __declspec(dllexport)
#else
#define AUTOLINKER_TEST_API __declspec(dllimport)
#endif

constexpr int AUTOLINKER_TEST_STRING_INVALID_ARGUMENT = -1;
constexpr int AUTOLINKER_TEST_STRING_BUFFER_TOO_SMALL = -2;

extern "C" {

// 比较两个版本号字符串，成功时通过 outResult 返回 -1/0/1。
AUTOLINKER_TEST_API bool AutoLinkerTest_CompareVersion(const char* left, const char* right, int* outResult);

// 提取 /out:"..." 中的输出文件名，成功时返回写入长度，不含结尾空字符。
AUTOLINKER_TEST_API int AutoLinkerTest_GetLinkerOutFileName(const char* commandLine, char* buffer, int bufferSize);

// 提取链接命令中的 krnln 静态库路径，成功时返回写入长度，不含结尾空字符。
AUTOLINKER_TEST_API int AutoLinkerTest_GetLinkerKrnlnFileName(const char* commandLine, char* buffer, int bufferSize);

// 提取两个 " - " 之间的文本，成功时返回写入长度，不含结尾空字符。
AUTOLINKER_TEST_API int AutoLinkerTest_ExtractBetweenDashes(const char* text, char* buffer, int bufferSize);

// 返回当前 AutoLinker 版本文本，成功时返回写入长度，不含结尾空字符。
AUTOLINKER_TEST_API int AutoLinkerTest_GetVersionText(char* buffer, int bufferSize);

// 对指定 DeepSeek 模型执行连通性、普通对话、thinking、多轮工具调用集成测试，并输出 JSON 结果。
AUTOLINKER_TEST_API int AutoLinkerTest_RunDeepSeekModelIntegrationTest(
	const char* apiKey,
	const char* model,
	const char* baseUrl,
	char* buffer,
	int bufferSize);

// 对指定 DeepSeek 模型仅执行基础连通性测试，并输出 JSON 结果。
AUTOLINKER_TEST_API int AutoLinkerTest_RunDeepSeekConnectionOnly(
	const char* apiKey,
	const char* model,
	const char* baseUrl,
	char* buffer,
	int bufferSize);

// 对指定 OpenAI Chat Completions 模型执行连通性、普通对话、多轮工具调用集成测试，并输出 JSON 结果。
AUTOLINKER_TEST_API int AutoLinkerTest_RunOpenAIChatIntegrationTest(
	const char* apiKey,
	const char* model,
	const char* baseUrl,
	char* buffer,
	int bufferSize);

// 对指定 OpenAI Responses 模型执行连通性、普通对话、多轮工具调用集成测试，并输出 JSON 结果。
AUTOLINKER_TEST_API int AutoLinkerTest_RunOpenAIResponsesIntegrationTest(
	const char* apiKey,
	const char* model,
	const char* baseUrl,
	char* buffer,
	int bufferSize);

// 对指定 Gemini 模型执行连通性、普通对话、多轮工具调用集成测试，并输出 JSON 结果。
AUTOLINKER_TEST_API int AutoLinkerTest_RunGeminiIntegrationTest(
	const char* apiKey,
	const char* model,
	const char* baseUrl,
	char* buffer,
	int bufferSize);

// 对指定 Claude 模型执行连通性、普通对话、多轮工具调用集成测试，并输出 JSON 结果。
AUTOLINKER_TEST_API int AutoLinkerTest_RunClaudeIntegrationTest(
	const char* apiKey,
	const char* model,
	const char* baseUrl,
	char* buffer,
	int bufferSize);

}
