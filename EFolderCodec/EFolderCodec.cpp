#include <cstdlib>
#include <winsock2.h>
#include <ws2tcpip.h>
#include <Windows.h>

#include <algorithm>
#include <chrono>
#include <filesystem>
#include <fstream>
#include <iostream>
#include <map>
#include <optional>
#include <set>
#include <string>
#include <thread>
#include <vector>

#include "..\thirdparty\json.hpp"
#include "..\src\EFolderCodec.h"
#include "..\src\e2txt.h"

#pragma comment(lib, "Ws2_32.lib")

namespace {

using json = nlohmann::json;

constexpr const char* kDefaultEidePath = "C:\\Users\\aiqin\\OneDrive\\e5.6\\e5.95.exe";
constexpr const char* kMcpHost = "127.0.0.1";
constexpr int kMcpBasePort = 19207;
constexpr int kMcpMaxPortAttempts = 16;
constexpr int kDefaultOpenTimeoutSeconds = 30;

struct IdeInstanceInfo {
	int port = 0;
	std::uint32_t processId = 0;
	std::string endpoint;
	std::string sourceFilePath;
	std::string projectType;
	std::string processPath;
	json structured;
};

int PrintStringResult(const char* label, int result, const char* text)
{
	if (result >= 0) {
		std::cout << label << ": " << text << std::endl;
		return EXIT_SUCCESS;
	}

	if (text != nullptr && text[0] != '\0') {
		std::cerr << label << " failed: " << text << std::endl;
	}
	else if (result == -2) {
		std::cerr << label << " failed: buffer too small" << std::endl;
	}
	else {
		std::cerr << label << " failed: invalid argument" << std::endl;
	}
	return EXIT_FAILURE;
}

std::wstring Utf8ToWide(const std::string& text)
{
	if (text.empty()) {
		return std::wstring();
	}

	const int wideLen = MultiByteToWideChar(
		CP_UTF8,
		MB_ERR_INVALID_CHARS,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0);
	if (wideLen <= 0) {
		return std::wstring();
	}

	std::wstring wide(static_cast<size_t>(wideLen), L'\0');
	if (MultiByteToWideChar(
			CP_UTF8,
			MB_ERR_INVALID_CHARS,
			text.data(),
			static_cast<int>(text.size()),
			wide.data(),
			wideLen) <= 0) {
		return std::wstring();
	}
	return wide;
}

std::string WideToUtf8(const std::wstring& text)
{
	if (text.empty()) {
		return std::string();
	}

	const int utf8Len = WideCharToMultiByte(
		CP_UTF8,
		0,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0,
		nullptr,
		nullptr);
	if (utf8Len <= 0) {
		return std::string();
	}

	std::string utf8(static_cast<size_t>(utf8Len), '\0');
	if (WideCharToMultiByte(
			CP_UTF8,
			0,
			text.data(),
			static_cast<int>(text.size()),
			utf8.data(),
			utf8Len,
			nullptr,
			nullptr) <= 0) {
		return std::string();
	}
	return utf8;
}

std::string PathToUtf8(const std::filesystem::path& path)
{
	return WideToUtf8(path.wstring());
}

std::string ToLowerAscii(std::string text)
{
	std::transform(text.begin(), text.end(), text.begin(), [](unsigned char ch) {
		return static_cast<char>(std::tolower(ch));
	});
	return text;
}

std::string NormalizeUtf8PathForCompare(const std::filesystem::path& path)
{
	if (path.empty()) {
		return std::string();
	}

	std::error_code ec;
	std::filesystem::path normalized = std::filesystem::absolute(path, ec);
	if (ec) {
		normalized = path;
	}
	normalized = normalized.lexically_normal();

	std::string text = PathToUtf8(normalized);
	std::replace(text.begin(), text.end(), '/', '\\');
	return ToLowerAscii(text);
}

std::string NormalizeUtf8PathTextForCompare(const std::string& utf8PathText)
{
	if (utf8PathText.empty()) {
		return std::string();
	}

	const std::wstring wide = Utf8ToWide(utf8PathText);
	if (wide.empty()) {
		std::string text = utf8PathText;
		std::replace(text.begin(), text.end(), '/', '\\');
		return ToLowerAscii(text);
	}
	return NormalizeUtf8PathForCompare(std::filesystem::path(wide));
}

bool DoUnpack(const std::string& inputPath, const std::string& outputDir, std::string& outSummary, std::string& outError)
{
	e2txt::Generator generator;
	e2txt::ProjectBundle bundle;
	if (!generator.GenerateBundle(inputPath, bundle, &outError)) {
		return false;
	}

	e2txt::BundleDirectoryCodec codec;
	if (!codec.WriteBundle(bundle, outputDir, &outError)) {
		return false;
	}

	outSummary =
		"source_files=" + std::to_string(bundle.sourceFiles.size()) +
		", form_files=" + std::to_string(bundle.formFiles.size()) +
		", resources=" + std::to_string(bundle.resources.size()) +
		", output=" + outputDir;
	return true;
}

bool DoPack(const std::string& inputDir, const std::string& outputPath, std::string& outSummary, std::string& outError)
{
	e2txt::BundleDirectoryCodec codec;
	e2txt::ProjectBundle bundle;
	if (!codec.ReadBundle(inputDir, bundle, &outError)) {
		return false;
	}

	e2txt::Restorer restorer;
	if (!restorer.RestoreBundleToFile(bundle, outputPath, &outSummary, &outError)) {
		return false;
	}
	return true;
}

int RunUnpack(const char* inputPath, const char* outputDir)
{
	std::string summary;
	std::string error;
	if (!DoUnpack(inputPath, outputDir, summary, error)) {
		return PrintStringResult("unpack", -1, error.c_str());
	}
	return PrintStringResult("unpack", 0, summary.c_str());
}

int RunPack(const char* inputDir, const char* outputPath)
{
	std::string summary;
	std::string error;
	if (!DoPack(inputDir, outputPath, summary, error)) {
		return PrintStringResult("pack", -1, error.c_str());
	}
	return PrintStringResult("pack", 0, summary.c_str());
}

int RunRoundTrip(const char* inputPath, const char* workDir, const char* outputPath)
{
	const std::filesystem::path root(workDir);
	const std::filesystem::path unpackDir = root / "unpacked";
	std::error_code ec;
	std::filesystem::remove_all(unpackDir, ec);
	std::filesystem::create_directories(unpackDir, ec);

	std::string summary;
	std::string error;
	if (!DoUnpack(inputPath, unpackDir.string(), summary, error)) {
		return PrintStringResult("roundtrip", -1, error.c_str());
	}
	if (!DoPack(unpackDir.string(), outputPath, summary, error)) {
		return PrintStringResult("roundtrip", -1, error.c_str());
	}
	return PrintStringResult("roundtrip", 0, summary.c_str());
}

bool ReadFileBytes(const std::filesystem::path& path, std::vector<std::uint8_t>& outBytes, std::string& outError)
{
	outBytes.clear();

	std::ifstream in(path, std::ios::binary);
	if (!in.is_open()) {
		outError = "open_failed: " + PathToUtf8(path);
		return false;
	}

	in.seekg(0, std::ios::end);
	const std::streamoff size = in.tellg();
	if (size < 0) {
		outError = "tellg_failed: " + PathToUtf8(path);
		return false;
	}
	in.seekg(0, std::ios::beg);

	outBytes.resize(static_cast<size_t>(size));
	if (size > 0) {
		in.read(reinterpret_cast<char*>(outBytes.data()), size);
		if (!in.good() && static_cast<size_t>(in.gcount()) != outBytes.size()) {
			outError = "read_failed: " + PathToUtf8(path);
			return false;
		}
	}

	return true;
}

std::string StripUtf8Bom(const std::string& text)
{
	if (text.size() >= 3 &&
		static_cast<unsigned char>(text[0]) == 0xEF &&
		static_cast<unsigned char>(text[1]) == 0xBB &&
		static_cast<unsigned char>(text[2]) == 0xBF) {
		return text.substr(3);
	}
	return text;
}

std::string NormalizeTextForCompare(const std::string& text)
{
	const std::string withoutBom = StripUtf8Bom(text);
	std::string normalized;
	normalized.reserve(withoutBom.size());

	size_t lineStart = 0;
	while (lineStart <= withoutBom.size()) {
		size_t lineEnd = withoutBom.find_first_of("\r\n", lineStart);
		if (lineEnd == std::string::npos) {
			lineEnd = withoutBom.size();
		}

		size_t contentStart = lineStart;
		while (contentStart < lineEnd &&
			(withoutBom[contentStart] == ' ' || withoutBom[contentStart] == '\t')) {
			++contentStart;
		}

		size_t contentEnd = lineEnd;
		while (contentEnd > contentStart &&
			(withoutBom[contentEnd - 1] == ' ' || withoutBom[contentEnd - 1] == '\t')) {
			--contentEnd;
		}

		if (contentEnd > contentStart) {
			normalized.append(withoutBom, contentStart, contentEnd - contentStart);
			normalized.push_back('\n');
		}

		if (lineEnd == withoutBom.size()) {
			break;
		}
		lineStart = lineEnd + 1;
		if (lineStart < withoutBom.size() &&
			withoutBom[lineEnd] == '\r' &&
			withoutBom[lineStart] == '\n') {
			++lineStart;
		}
	}

	while (!normalized.empty() && normalized.back() == '\n') {
		normalized.pop_back();
	}
	return normalized;
}

bool CompareNormalizedTextFile(
	const std::filesystem::path& leftPath,
	const std::filesystem::path& rightPath,
	std::string& outSummary)
{
	std::vector<std::uint8_t> leftBytes;
	std::vector<std::uint8_t> rightBytes;
	std::string error;
	if (!ReadFileBytes(leftPath, leftBytes, error)) {
		outSummary = error;
		return false;
	}
	if (!ReadFileBytes(rightPath, rightBytes, error)) {
		outSummary = error;
		return false;
	}

	const std::string leftText = NormalizeTextForCompare(std::string(leftBytes.begin(), leftBytes.end()));
	const std::string rightText = NormalizeTextForCompare(std::string(rightBytes.begin(), rightBytes.end()));
	if (leftText == rightText) {
		return true;
	}

	outSummary = "text_mismatch: " + PathToUtf8(leftPath);
	return false;
}

void NormalizeJsonForCompare(json& value)
{
	if (!value.is_object()) {
		return;
	}

	value.erase("sourcePath");
	value.erase("nativeBundleDigest");
	value.erase("projectNameStored");

	auto it = value.find("rootChildKeys");
	if (it != value.end() && it->is_array()) {
		std::vector<std::string> keys;
		for (const auto& item : *it) {
			if (item.is_string()) {
				keys.push_back(item.get<std::string>());
			}
		}
		std::sort(keys.begin(), keys.end());
		*it = json::array();
		for (const auto& key : keys) {
			it->push_back(key);
		}
	}
}

bool ShouldIgnorePathForRoundTripCompare(const std::string& relativePath)
{
	return relativePath.starts_with("src/.native_");
}

bool CompareJsonFile(
	const std::filesystem::path& leftPath,
	const std::filesystem::path& rightPath,
	std::string& outSummary)
{
	std::vector<std::uint8_t> leftBytes;
	std::vector<std::uint8_t> rightBytes;
	std::string error;
	if (!ReadFileBytes(leftPath, leftBytes, error)) {
		outSummary = error;
		return false;
	}
	if (!ReadFileBytes(rightPath, rightBytes, error)) {
		outSummary = error;
		return false;
	}

	try {
		auto leftJson = json::parse(StripUtf8Bom(std::string(leftBytes.begin(), leftBytes.end())));
		auto rightJson = json::parse(StripUtf8Bom(std::string(rightBytes.begin(), rightBytes.end())));
		NormalizeJsonForCompare(leftJson);
		NormalizeJsonForCompare(rightJson);
		if (leftJson == rightJson) {
			return true;
		}

		outSummary = "json_mismatch: " + PathToUtf8(leftPath);
		return false;
	}
	catch (const std::exception& ex) {
		outSummary = std::string("json_parse_failed: ") + ex.what();
		return false;
	}
}

bool BuildFileMap(
	const std::filesystem::path& root,
	std::map<std::string, std::filesystem::path>& outFiles,
	std::string& outError)
{
	outFiles.clear();

	std::error_code ec;
	if (!std::filesystem::exists(root, ec)) {
		outError = "path_not_found: " + PathToUtf8(root);
		return false;
	}

	for (std::filesystem::recursive_directory_iterator it(root, ec), end; it != end; it.increment(ec)) {
		if (ec) {
			outError = "enumerate_failed: " + PathToUtf8(root);
			return false;
		}
		if (!it->is_regular_file()) {
			continue;
		}

		const std::filesystem::path relative = std::filesystem::relative(it->path(), root, ec);
		if (ec) {
			outError = "relative_path_failed: " + PathToUtf8(it->path());
			return false;
		}
		outFiles.emplace(relative.generic_string(), it->path());
	}

	return true;
}

bool CompareDirectoryTrees(
	const std::filesystem::path& leftRoot,
	const std::filesystem::path& rightRoot,
	std::string& outSummary)
{
	outSummary.clear();

	std::map<std::string, std::filesystem::path> leftFiles;
	std::map<std::string, std::filesystem::path> rightFiles;
	std::string error;
	if (!BuildFileMap(leftRoot, leftFiles, error)) {
		outSummary = error;
		return false;
	}
	if (!BuildFileMap(rightRoot, rightFiles, error)) {
		outSummary = error;
		return false;
	}

	size_t comparedCount = 0;
	for (const auto& [relativePath, leftPath] : leftFiles) {
		if (ShouldIgnorePathForRoundTripCompare(relativePath)) {
			continue;
		}

		const auto rightIt = rightFiles.find(relativePath);
		if (rightIt == rightFiles.end()) {
			outSummary = "missing_in_roundtrip: " + relativePath;
			return false;
		}

		const std::filesystem::path extension = leftPath.extension();
		if (extension == std::filesystem::path(L".json")) {
			if (!CompareJsonFile(leftPath, rightIt->second, outSummary)) {
				return false;
			}
			++comparedCount;
			continue;
		}
		if (extension == std::filesystem::path(L".txt") ||
			extension == std::filesystem::path(L".xml")) {
			if (!CompareNormalizedTextFile(leftPath, rightIt->second, outSummary)) {
				return false;
			}
			++comparedCount;
			continue;
		}

		std::vector<std::uint8_t> leftBytes;
		std::vector<std::uint8_t> rightBytes;
		if (!ReadFileBytes(leftPath, leftBytes, error)) {
			outSummary = error;
			return false;
		}
		if (!ReadFileBytes(rightIt->second, rightBytes, error)) {
			outSummary = error;
			return false;
		}
		if (leftBytes != rightBytes) {
			outSummary =
				"content_mismatch: " + relativePath +
				", left_bytes=" + std::to_string(leftBytes.size()) +
				", right_bytes=" + std::to_string(rightBytes.size());
			return false;
		}
		++comparedCount;
	}

	for (const auto& [relativePath, rightPath] : rightFiles) {
		(void)rightPath;
		if (ShouldIgnorePathForRoundTripCompare(relativePath)) {
			continue;
		}
		if (!leftFiles.contains(relativePath)) {
			outSummary = "extra_in_roundtrip: " + relativePath;
			return false;
		}
	}

	outSummary =
		"compared_files=" + std::to_string(comparedCount) +
		", left=" + PathToUtf8(leftRoot) +
		", right=" + PathToUtf8(rightRoot);
	return true;
}

int RunVerifyRoundTrip(const char* inputPath, const char* workDir, const char* outputPath)
{
	const std::filesystem::path root(workDir);
	const std::filesystem::path originalDir = root / "original_unpacked";
	const std::filesystem::path roundtripDir = root / "roundtrip_unpacked";
	std::error_code ec;
	std::filesystem::remove_all(root, ec);
	std::filesystem::create_directories(root, ec);

	std::string summary;
	std::string error;
	if (!DoUnpack(inputPath, originalDir.string(), summary, error)) {
		return PrintStringResult("verify-roundtrip", -1, error.c_str());
	}
	if (!DoPack(originalDir.string(), outputPath, summary, error)) {
		return PrintStringResult("verify-roundtrip", -1, error.c_str());
	}
	if (!DoUnpack(outputPath, roundtripDir.string(), summary, error)) {
		return PrintStringResult("verify-roundtrip", -1, error.c_str());
	}

	std::string compareSummary;
	if (!CompareDirectoryTrees(originalDir, roundtripDir, compareSummary)) {
		return PrintStringResult("verify-roundtrip", -1, compareSummary.c_str());
	}

	return PrintStringResult("verify-roundtrip", 0, compareSummary.c_str());
}

class WsaSession {
public:
	WsaSession()
	{
		m_ok = WSAStartup(MAKEWORD(2, 2), &m_data) == 0;
	}

	~WsaSession()
	{
		if (m_ok) {
			WSACleanup();
		}
	}

	bool IsOk() const
	{
		return m_ok;
	}

private:
	WSADATA m_data = {};
	bool m_ok = false;
};

bool TryHttpPostJson(const int port, const std::string& requestBody, std::string& outBody, std::string& outError)
{
	outBody.clear();
	outError.clear();

	WsaSession wsa;
	if (!wsa.IsOk()) {
		outError = "wsa_startup_failed";
		return false;
	}

	SOCKET sock = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP);
	if (sock == INVALID_SOCKET) {
		outError = "socket_create_failed";
		return false;
	}

	const DWORD sendTimeoutMs = 5000;
	const DWORD recvTimeoutMs = 120000;
	setsockopt(sock, SOL_SOCKET, SO_RCVTIMEO, reinterpret_cast<const char*>(&recvTimeoutMs), sizeof(recvTimeoutMs));
	setsockopt(sock, SOL_SOCKET, SO_SNDTIMEO, reinterpret_cast<const char*>(&sendTimeoutMs), sizeof(sendTimeoutMs));

	sockaddr_in addr = {};
	addr.sin_family = AF_INET;
	addr.sin_port = htons(static_cast<u_short>(port));
	if (inet_pton(AF_INET, kMcpHost, &addr.sin_addr) != 1) {
		closesocket(sock);
		outError = "inet_pton_failed";
		return false;
	}

	if (connect(sock, reinterpret_cast<const sockaddr*>(&addr), sizeof(addr)) == SOCKET_ERROR) {
		closesocket(sock);
		outError = "connect_failed";
		return false;
	}

	const std::string request =
		"POST /mcp HTTP/1.1\r\n"
		"Host: 127.0.0.1:" + std::to_string(port) + "\r\n"
		"Content-Type: application/json; charset=utf-8\r\n"
		"Connection: close\r\n"
		"Content-Length: " + std::to_string(requestBody.size()) + "\r\n\r\n" +
		requestBody;

	size_t sentTotal = 0;
	while (sentTotal < request.size()) {
		const int sent = send(
			sock,
			request.data() + sentTotal,
			static_cast<int>(request.size() - sentTotal),
			0);
		if (sent <= 0) {
			closesocket(sock);
			outError = "send_failed";
			return false;
		}
		sentTotal += static_cast<size_t>(sent);
	}

	std::string response;
	size_t headerEnd = std::string::npos;
	std::optional<size_t> expectedBodySize;
	const auto tryParseContentLength = [&response, &headerEnd, &expectedBodySize]() {
		if (headerEnd == std::string::npos || expectedBodySize.has_value()) {
			return;
		}
		const std::string headers = response.substr(0, headerEnd);
		const std::string contentLengthHeader = "\r\nContent-Length:";
		size_t contentLengthPos = headers.find(contentLengthHeader);
		if (contentLengthPos == std::string::npos) {
			if (headers.rfind("Content-Length:", 0) == 0) {
				contentLengthPos = 0;
			}
			else {
				return;
			}
		}
		if (contentLengthPos != 0) {
			contentLengthPos += 2;
		}
		contentLengthPos += sizeof("Content-Length:") - 1;
		while (contentLengthPos < headers.size() &&
			(headers[contentLengthPos] == ' ' || headers[contentLengthPos] == '\t')) {
			++contentLengthPos;
		}
		size_t contentLengthEnd = contentLengthPos;
		while (contentLengthEnd < headers.size() &&
			headers[contentLengthEnd] >= '0' &&
			headers[contentLengthEnd] <= '9') {
			++contentLengthEnd;
		}
		if (contentLengthEnd == contentLengthPos) {
			return;
		}
		try {
			expectedBodySize = static_cast<size_t>(std::stoull(headers.substr(contentLengthPos, contentLengthEnd - contentLengthPos)));
		}
		catch (...) {
			expectedBodySize.reset();
		}
	};
	char buffer[4096] = {};
	while (true) {
		const int received = recv(sock, buffer, static_cast<int>(sizeof(buffer)), 0);
		if (received == 0) {
			break;
		}
		if (received < 0) {
			const int recvError = WSAGetLastError();
			if (!response.empty()) {
				if (headerEnd == std::string::npos) {
					headerEnd = response.find("\r\n\r\n");
				}
				tryParseContentLength();
				if (headerEnd != std::string::npos) {
					const size_t bodySize = response.size() - (headerEnd + 4);
					if (!expectedBodySize.has_value() || bodySize >= *expectedBodySize) {
						break;
					}
				}
			}
			closesocket(sock);
			outError = "recv_failed:" + std::to_string(recvError);
			return false;
		}
		response.append(buffer, static_cast<size_t>(received));
		if (headerEnd == std::string::npos) {
			headerEnd = response.find("\r\n\r\n");
		}
		tryParseContentLength();
		if (headerEnd != std::string::npos &&
			expectedBodySize.has_value() &&
			response.size() >= headerEnd + 4 + *expectedBodySize) {
			break;
		}
	}

	closesocket(sock);

	if (headerEnd == std::string::npos) {
		headerEnd = response.find("\r\n\r\n");
	}
	if (headerEnd == std::string::npos) {
		outError = "http_header_missing";
		return false;
	}

	if (expectedBodySize.has_value()) {
		const size_t bodySize = response.size() - (headerEnd + 4);
		if (bodySize < *expectedBodySize) {
			outError = "http_body_truncated";
			return false;
		}
	}

	outBody = response.substr(headerEnd + 4);
	return true;
}

std::string BuildJsonRpcErrorText(const json& response)
{
	if (!response.is_object()) {
		return "invalid_json_rpc_response";
	}
	if (response.contains("error") && response["error"].is_object()) {
		const auto& error = response["error"];
		if (error.contains("message") && error["message"].is_string()) {
			return error["message"].get<std::string>();
		}
	}
	return "json_rpc_error";
}

bool TryCallMcpTool(
	const int port,
	const std::string& toolName,
	const json& arguments,
	json& outStructured,
	std::string& outError)
{
	outStructured = json::object();
	outError.clear();

	const json request = {
		{"jsonrpc", "2.0"},
		{"id", 1},
		{"method", "tools/call"},
		{"params", {
			{"name", toolName},
			{"arguments", arguments}
		}}
	};

	std::string responseBody;
	if (!TryHttpPostJson(port, request.dump(), responseBody, outError)) {
		return false;
	}

	json response;
	try {
		response = json::parse(responseBody);
	}
	catch (const std::exception& ex) {
		outError = std::string("json_parse_failed: ") + ex.what();
		return false;
	}

	if (!response.is_object()) {
		outError = "json_rpc_response_invalid";
		return false;
	}
	if (response.contains("error")) {
		outError = BuildJsonRpcErrorText(response);
		return false;
	}
	if (!response.contains("result") || !response["result"].is_object()) {
		outError = "json_rpc_result_missing";
		return false;
	}

	const auto& result = response["result"];
	if (result.contains("structuredContent") && result["structuredContent"].is_object()) {
		outStructured = result["structuredContent"];
	}

	const bool isError = result.value("isError", false);
	if (!isError) {
		return true;
	}

	if (outStructured.contains("error") && outStructured["error"].is_string()) {
		outError = outStructured["error"].get<std::string>();
	}
	else if (result.contains("content") && result["content"].is_array() && !result["content"].empty()) {
		const auto& first = result["content"][0];
		if (first.is_object() && first.contains("text") && first["text"].is_string()) {
			outError = first["text"].get<std::string>();
		}
		else {
			outError = "tool_call_failed";
		}
	}
	else {
		outError = "tool_call_failed";
	}
	return false;
}

bool TryQueryInstanceInfo(const int port, IdeInstanceInfo& outInfo, std::string& outError)
{
	outInfo = {};

	json structured;
	if (!TryCallMcpTool(port, "get_current_eide_info", json::object(), structured, outError)) {
		return false;
	}
	if (!structured.value("ok", false)) {
		outError = structured.value("error", std::string("get_current_eide_info_not_ok"));
		return false;
	}

	outInfo.port = port;
	outInfo.processId = structured.value("process_id", 0U);
	outInfo.endpoint = structured.value("mcp_endpoint", std::string());
	outInfo.sourceFilePath = structured.value("source_file_path", std::string());
	outInfo.projectType = structured.value("project_type", std::string());
	outInfo.processPath = structured.value("process_path", std::string());
	outInfo.structured = std::move(structured);
	return true;
}

std::vector<IdeInstanceInfo> ScanEideInstances()
{
	std::vector<IdeInstanceInfo> instances;
	for (int port = kMcpBasePort; port < kMcpBasePort + kMcpMaxPortAttempts; ++port) {
		IdeInstanceInfo info;
		std::string error;
		if (TryQueryInstanceInfo(port, info, error)) {
			instances.push_back(std::move(info));
		}
	}
	return instances;
}

json BuildInstanceListJson(const std::vector<IdeInstanceInfo>& instances)
{
	json array = json::array();
	for (const auto& instance : instances) {
		array.push_back({
			{"port", instance.port},
			{"process_id", instance.processId},
			{"endpoint", instance.endpoint},
			{"source_file_path", instance.sourceFilePath},
			{"project_type", instance.projectType},
			{"process_path", instance.processPath}
		});
	}
	return array;
}

std::wstring QuoteCommandLineArg(const std::wstring& arg)
{
	if (arg.empty()) {
		return L"\"\"";
	}

	const bool needsQuote = arg.find_first_of(L" \t\"") != std::wstring::npos;
	if (!needsQuote) {
		return arg;
	}

	std::wstring quoted = L"\"";
	size_t backslashes = 0;
	for (const wchar_t ch : arg) {
		if (ch == L'\\') {
			++backslashes;
			continue;
		}
		if (ch == L'"') {
			quoted.append(backslashes * 2 + 1, L'\\');
			quoted.push_back(L'"');
			backslashes = 0;
			continue;
		}
		quoted.append(backslashes, L'\\');
		backslashes = 0;
		quoted.push_back(ch);
	}
	quoted.append(backslashes * 2, L'\\');
	quoted.push_back(L'"');
	return quoted;
}

bool LaunchEideProcess(
	const std::filesystem::path& eidePath,
	const std::filesystem::path& sourcePath,
	std::uint32_t& outProcessId,
	std::string& outError)
{
	outProcessId = 0;
	outError.clear();

	std::error_code ec;
	if (!std::filesystem::exists(eidePath, ec) || ec) {
		outError = "eide_not_found: " + PathToUtf8(eidePath);
		return false;
	}
	if (!std::filesystem::exists(sourcePath, ec) || ec) {
		outError = "roundtrip_e_not_found: " + PathToUtf8(sourcePath);
		return false;
	}

	const std::wstring commandLineWide =
		QuoteCommandLineArg(eidePath.wstring()) + L" " + QuoteCommandLineArg(sourcePath.wstring());
	std::vector<wchar_t> commandLine(commandLineWide.begin(), commandLineWide.end());
	commandLine.push_back(L'\0');

	STARTUPINFOW startupInfo = {};
	startupInfo.cb = sizeof(startupInfo);
	PROCESS_INFORMATION processInfo = {};

	const std::wstring workingDir = eidePath.parent_path().wstring();
	if (!CreateProcessW(
			nullptr,
			commandLine.data(),
			nullptr,
			nullptr,
			FALSE,
			0,
			nullptr,
			workingDir.empty() ? nullptr : workingDir.c_str(),
			&startupInfo,
			&processInfo)) {
		outError = "create_process_failed: " + std::to_string(GetLastError());
		return false;
	}

	outProcessId = processInfo.dwProcessId;
	CloseHandle(processInfo.hThread);
	CloseHandle(processInfo.hProcess);
	return true;
}

bool WaitForRoundtripInstance(
	const std::string& normalizedRoundtripPath,
	const std::set<std::uint32_t>& ignoredProcessIds,
	const std::uint32_t preferredProcessId,
	const int timeoutSeconds,
	IdeInstanceInfo& outInfo,
	std::vector<IdeInstanceInfo>& outLastInstances,
	std::string& outError)
{
	outInfo = {};
	outLastInstances.clear();
	outError.clear();

	const auto deadline = std::chrono::steady_clock::now() + std::chrono::seconds(timeoutSeconds);
	while (std::chrono::steady_clock::now() < deadline) {
		outLastInstances = ScanEideInstances();
		for (const auto& instance : outLastInstances) {
			if (NormalizeUtf8PathTextForCompare(instance.sourceFilePath) != normalizedRoundtripPath) {
				continue;
			}
			if (preferredProcessId != 0 && instance.processId == preferredProcessId) {
				outInfo = instance;
				return true;
			}
			if (!ignoredProcessIds.contains(instance.processId)) {
				outInfo = instance;
				return true;
			}
		}
		std::this_thread::sleep_for(std::chrono::milliseconds(500));
	}

	outError = "roundtrip_e_not_opened_in_e595";
	return false;
}

bool TryCompileEcom(
	const IdeInstanceInfo& instance,
	const std::filesystem::path& compileOutputPath,
	json& outCompileResult,
	std::string& outError)
{
	outCompileResult = json::object();
	outError.clear();

	const json arguments = {
		{"target", "ecom"},
		{"output_path", PathToUtf8(compileOutputPath)},
		{"static_compile", false}
	};
	if (!TryCallMcpTool(instance.port, "compile_with_output_path", arguments, outCompileResult, outError)) {
		return false;
	}
	if (!outCompileResult.value("ok", false)) {
		outError = outCompileResult.value("error", std::string("compile_with_output_path_not_ok"));
		return false;
	}
	return true;
}

int RunCompileEcomFile(
	const char* inputPath,
	const char* compileOutputPath,
	const char* eidePath,
	const int openTimeoutSeconds)
{
	std::error_code ec;
	const std::filesystem::path inputFile = std::filesystem::absolute(std::filesystem::path(inputPath), ec);
	ec.clear();
	const std::filesystem::path compileOutputFile = std::filesystem::absolute(std::filesystem::path(compileOutputPath), ec);
	ec.clear();
	const std::filesystem::path eideExePath = std::filesystem::absolute(
		std::filesystem::path(eidePath == nullptr || eidePath[0] == '\0' ? kDefaultEidePath : eidePath), ec);
	ec.clear();

	json summary = {
		{"ok", false},
		{"input_e_path", PathToUtf8(inputFile)},
		{"compile_output_path", PathToUtf8(compileOutputFile)},
		{"eide_path", PathToUtf8(eideExePath)}
	};

	std::filesystem::create_directories(compileOutputFile.parent_path(), ec);
	std::filesystem::remove(compileOutputFile, ec);

	const std::string normalizedInputPath = NormalizeUtf8PathForCompare(inputFile);
	const std::vector<IdeInstanceInfo> preLaunchInstances = ScanEideInstances();
	std::set<std::uint32_t> preExistingInputProcessIds;
	for (const auto& instance : preLaunchInstances) {
		if (NormalizeUtf8PathTextForCompare(instance.sourceFilePath) == normalizedInputPath) {
			preExistingInputProcessIds.insert(instance.processId);
		}
	}
	summary["prelaunch_instances"] = BuildInstanceListJson(preLaunchInstances);

	std::string error;
	std::uint32_t launchedProcessId = 0;
	if (!LaunchEideProcess(eideExePath, inputFile, launchedProcessId, error)) {
		summary["error"] = "launch_failed";
		summary["detail"] = error;
		return PrintStringResult("compile-ecom", -1, summary.dump().c_str());
	}
	summary["launched_process_id"] = launchedProcessId;

	IdeInstanceInfo openedInstance;
	std::vector<IdeInstanceInfo> lastInstances;
	if (!WaitForRoundtripInstance(
			normalizedInputPath,
			preExistingInputProcessIds,
			launchedProcessId,
			openTimeoutSeconds,
			openedInstance,
			lastInstances,
			error)) {
		summary["error"] = error;
		summary["last_scanned_instances"] = BuildInstanceListJson(lastInstances);
		return PrintStringResult("compile-ecom", -1, summary.dump().c_str());
	}

	summary["mcp_endpoint"] = openedInstance.endpoint;
	summary["ide_process_id"] = openedInstance.processId;
	summary["opened_source_file_path"] = openedInstance.sourceFilePath;
	summary["project_type"] = openedInstance.projectType;

	if (NormalizeUtf8PathTextForCompare(openedInstance.sourceFilePath) != normalizedInputPath) {
		summary["error"] = "opened_source_path_mismatch";
		return PrintStringResult("compile-ecom", -1, summary.dump().c_str());
	}
	if (openedInstance.projectType != "ecom") {
		summary["error"] = "project_type_invalid";
		return PrintStringResult("compile-ecom", -1, summary.dump().c_str());
	}

	json compileResult;
	if (!TryCompileEcom(openedInstance, compileOutputFile, compileResult, error)) {
		summary["error"] = "compile_failed";
		summary["detail"] = error;
		summary["compile_result"] = compileResult;
		return PrintStringResult("compile-ecom", -1, summary.dump().c_str());
	}

	summary["compile_result"] = compileResult;
	summary["compile_ok"] = compileResult.value("ok", false);
	summary["output_window_text"] = compileResult.value("output_window_text", std::string());
	summary["caret_page_name"] = compileResult.value("caret_page_name", std::string());
	summary["caret_row"] = compileResult.value("caret_row", -1);
	summary["caret_line_text"] = compileResult.value("caret_line_text", std::string());

	const std::string normalizedCompileOutputPath = NormalizeUtf8PathForCompare(compileOutputFile);
	const std::string normalizedReportedCompilePath =
		NormalizeUtf8PathTextForCompare(compileResult.value("output_path", std::string()));
	if (normalizedReportedCompilePath != normalizedCompileOutputPath) {
		summary["error"] = "compile_output_path_mismatch";
		return PrintStringResult("compile-ecom", -1, summary.dump().c_str());
	}
	if (!compileResult.value("output_file_exists", false) ||
		!compileResult.value("output_file_modified_after_compile", false)) {
		summary["error"] = "compile_output_not_updated";
		return PrintStringResult("compile-ecom", -1, summary.dump().c_str());
	}

	summary["ok"] = true;
	return PrintStringResult("compile-ecom", 0, summary.dump().c_str());
}

int RunVerifyRoundTripCompile(
	const char* inputPath,
	const char* workDir,
	const char* outputPath,
	const char* compileOutputPath,
	const char* eidePath,
	const int openTimeoutSeconds)
{
	std::error_code ec;
	const std::filesystem::path inputFile = std::filesystem::absolute(std::filesystem::path(inputPath), ec);
	ec.clear();
	const std::filesystem::path root = std::filesystem::absolute(std::filesystem::path(workDir), ec);
	ec.clear();
	const std::filesystem::path roundtripFile = std::filesystem::absolute(std::filesystem::path(outputPath), ec);
	ec.clear();
	const std::filesystem::path compileOutputFile = std::filesystem::absolute(std::filesystem::path(compileOutputPath), ec);
	ec.clear();
	const std::filesystem::path eideExePath = std::filesystem::absolute(
		std::filesystem::path(eidePath == nullptr || eidePath[0] == '\0' ? kDefaultEidePath : eidePath), ec);
	ec.clear();
	const std::filesystem::path originalDir = root / "original_unpacked";
	const std::filesystem::path roundtripDir = root / "roundtrip_unpacked";

	json summary = {
		{"ok", false},
		{"original_e_path", PathToUtf8(inputFile)},
		{"roundtrip_e_path", PathToUtf8(roundtripFile)},
		{"compile_output_path", PathToUtf8(compileOutputFile)},
		{"eide_path", PathToUtf8(eideExePath)}
	};

	std::filesystem::remove_all(root, ec);
	std::filesystem::create_directories(root, ec);
	std::filesystem::create_directories(roundtripFile.parent_path(), ec);
	std::filesystem::create_directories(compileOutputFile.parent_path(), ec);
	std::filesystem::remove(compileOutputFile, ec);

	std::string stepSummary;
	std::string error;
	if (!DoUnpack(inputPath, originalDir.string(), stepSummary, error)) {
		summary["error"] = "unpack_generate_failed";
		summary["detail"] = error;
		return PrintStringResult("verify-roundtrip-compile", -1, summary.dump().c_str());
	}
	if (!DoPack(originalDir.string(), roundtripFile.string(), stepSummary, error)) {
		summary["error"] = "roundtrip_pack_failed";
		summary["detail"] = error;
		return PrintStringResult("verify-roundtrip-compile", -1, summary.dump().c_str());
	}
	if (!DoUnpack(roundtripFile.string(), roundtripDir.string(), stepSummary, error)) {
		summary["error"] = "roundtrip_unpack_failed";
		summary["detail"] = error;
		return PrintStringResult("verify-roundtrip-compile", -1, summary.dump().c_str());
	}

	std::string compareSummary;
	if (!CompareDirectoryTrees(originalDir, roundtripDir, compareSummary)) {
		summary["error"] = "roundtrip_compare_failed";
		summary["detail"] = compareSummary;
		return PrintStringResult("verify-roundtrip-compile", -1, summary.dump().c_str());
	}
	summary["roundtrip_compare_ok"] = true;
	summary["roundtrip_compare_summary"] = compareSummary;

	const std::string normalizedInputPath = NormalizeUtf8PathForCompare(inputFile);
	const std::string normalizedRoundtripPath = NormalizeUtf8PathForCompare(roundtripFile);
	if (normalizedInputPath == normalizedRoundtripPath) {
		summary["error"] = "roundtrip_path_matches_original";
		return PrintStringResult("verify-roundtrip-compile", -1, summary.dump().c_str());
	}

	const std::vector<IdeInstanceInfo> preLaunchInstances = ScanEideInstances();
	std::set<std::uint32_t> preExistingRoundtripProcessIds;
	for (const auto& instance : preLaunchInstances) {
		if (NormalizeUtf8PathTextForCompare(instance.sourceFilePath) == normalizedRoundtripPath) {
			preExistingRoundtripProcessIds.insert(instance.processId);
		}
	}
	summary["prelaunch_instances"] = BuildInstanceListJson(preLaunchInstances);

	std::uint32_t launchedProcessId = 0;
	if (!LaunchEideProcess(eideExePath, roundtripFile, launchedProcessId, error)) {
		summary["error"] = "roundtrip_launch_failed";
		summary["detail"] = error;
		return PrintStringResult("verify-roundtrip-compile", -1, summary.dump().c_str());
	}
	summary["launched_process_id"] = launchedProcessId;

	IdeInstanceInfo roundtripInstance;
	std::vector<IdeInstanceInfo> lastInstances;
	if (!WaitForRoundtripInstance(
			normalizedRoundtripPath,
			preExistingRoundtripProcessIds,
			launchedProcessId,
			openTimeoutSeconds,
			roundtripInstance,
			lastInstances,
			error)) {
		summary["error"] = error;
		summary["last_scanned_instances"] = BuildInstanceListJson(lastInstances);
		return PrintStringResult("verify-roundtrip-compile", -1, summary.dump().c_str());
	}

	summary["mcp_endpoint"] = roundtripInstance.endpoint;
	summary["ide_process_id"] = roundtripInstance.processId;
	summary["opened_source_file_path"] = roundtripInstance.sourceFilePath;
	summary["project_type"] = roundtripInstance.projectType;

	if (NormalizeUtf8PathTextForCompare(roundtripInstance.sourceFilePath) != normalizedRoundtripPath) {
		summary["error"] = "opened_source_path_mismatch";
		return PrintStringResult("verify-roundtrip-compile", -1, summary.dump().c_str());
	}
	if (NormalizeUtf8PathTextForCompare(roundtripInstance.sourceFilePath) == normalizedInputPath) {
		summary["error"] = "opened_original_project_instead_of_roundtrip";
		return PrintStringResult("verify-roundtrip-compile", -1, summary.dump().c_str());
	}
	if (roundtripInstance.projectType != "ecom") {
		summary["error"] = "roundtrip_project_type_invalid";
		return PrintStringResult("verify-roundtrip-compile", -1, summary.dump().c_str());
	}

	json compileResult;
	if (!TryCompileEcom(roundtripInstance, compileOutputFile, compileResult, error)) {
		summary["error"] = "compile_roundtrip_failed";
		summary["detail"] = error;
		summary["compile_result"] = compileResult;
		return PrintStringResult("verify-roundtrip-compile", -1, summary.dump().c_str());
	}

	summary["compile_result"] = compileResult;
	summary["compile_ok"] = compileResult.value("ok", false);
	summary["output_window_text"] = compileResult.value("output_window_text", std::string());
	summary["caret_page_name"] = compileResult.value("caret_page_name", std::string());
	summary["caret_row"] = compileResult.value("caret_row", -1);
	summary["caret_line_text"] = compileResult.value("caret_line_text", std::string());

	const std::string normalizedCompileOutputPath = NormalizeUtf8PathForCompare(compileOutputFile);
	const std::string normalizedReportedCompilePath =
		NormalizeUtf8PathTextForCompare(compileResult.value("output_path", std::string()));
	if (normalizedCompileOutputPath != normalizedReportedCompilePath) {
		summary["error"] = "compile_output_path_mismatch";
		return PrintStringResult("verify-roundtrip-compile", -1, summary.dump().c_str());
	}
	if (!compileResult.value("output_file_exists", false) ||
		!compileResult.value("output_file_modified_after_compile", false)) {
		summary["error"] = "compile_output_not_updated";
		return PrintStringResult("verify-roundtrip-compile", -1, summary.dump().c_str());
	}

	summary["ok"] = true;
	return PrintStringResult("verify-roundtrip-compile", 0, summary.dump().c_str());
}

void PrintUsage()
{
	std::cout << "EFolderCodec commands:" << std::endl;
	std::cout << "  EFolderCodec unpack <input.e> <output-dir>" << std::endl;
	std::cout << "  EFolderCodec pack <input-dir> <output.e>" << std::endl;
	std::cout << "  EFolderCodec roundtrip <input.e> <work-dir> <output.e>" << std::endl;
	std::cout << "  EFolderCodec verify-roundtrip <input.e> <work-dir> <output.e>" << std::endl;
	std::cout << "  EFolderCodec compile-ecom <input.e> <output.ec> [--eide-path <e5.95.exe>] [--open-timeout-seconds <n>]" << std::endl;
	std::cout << "  EFolderCodec verify-roundtrip-compile <input.e> <work-dir> <roundtrip.e> <output.ec> [--eide-path <e5.95.exe>] [--open-timeout-seconds <n>]" << std::endl;
}

}  // namespace

int main(int argc, char* argv[])
{
	if (argc < 2) {
		PrintUsage();
		return EXIT_FAILURE;
	}

	const std::string command = argv[1];
	if (command == "unpack") {
		if (argc != 4) {
			PrintUsage();
			return EXIT_FAILURE;
		}
		return RunUnpack(argv[2], argv[3]);
	}
	if (command == "pack") {
		if (argc != 4) {
			PrintUsage();
			return EXIT_FAILURE;
		}
		return RunPack(argv[2], argv[3]);
	}
	if (command == "roundtrip") {
		if (argc != 5) {
			PrintUsage();
			return EXIT_FAILURE;
		}
		return RunRoundTrip(argv[2], argv[3], argv[4]);
	}
	if (command == "verify-roundtrip") {
		if (argc != 5) {
			PrintUsage();
			return EXIT_FAILURE;
		}
		return RunVerifyRoundTrip(argv[2], argv[3], argv[4]);
	}
	if (command == "compile-ecom") {
		if (argc < 4) {
			PrintUsage();
			return EXIT_FAILURE;
		}

		const char* eidePath = kDefaultEidePath;
		int openTimeoutSeconds = kDefaultOpenTimeoutSeconds;
		for (int index = 4; index < argc; ++index) {
			const std::string arg = argv[index];
			if (arg == "--eide-path") {
				if (index + 1 >= argc) {
					PrintUsage();
					return EXIT_FAILURE;
				}
				eidePath = argv[++index];
				continue;
			}
			if (arg == "--open-timeout-seconds") {
				if (index + 1 >= argc) {
					PrintUsage();
					return EXIT_FAILURE;
				}
				openTimeoutSeconds = (std::max)(1, std::atoi(argv[++index]));
				continue;
			}
		}
		return RunCompileEcomFile(argv[2], argv[3], eidePath, openTimeoutSeconds);
	}
	if (command == "verify-roundtrip-compile") {
		if (argc < 6) {
			PrintUsage();
			return EXIT_FAILURE;
		}

		const char* eidePath = kDefaultEidePath;
		int openTimeoutSeconds = kDefaultOpenTimeoutSeconds;
		for (int index = 6; index < argc; ++index) {
			const std::string arg = argv[index];
			if (arg == "--eide-path") {
				if (index + 1 >= argc) {
					PrintUsage();
					return EXIT_FAILURE;
				}
				eidePath = argv[++index];
				continue;
			}
			if (arg == "--open-timeout-seconds") {
				if (index + 1 >= argc) {
					PrintUsage();
					return EXIT_FAILURE;
				}
				openTimeoutSeconds = (std::max)(1, std::atoi(argv[++index]));
				continue;
			}

			PrintUsage();
			return EXIT_FAILURE;
		}

		return RunVerifyRoundTripCompile(argv[2], argv[3], argv[4], argv[5], eidePath, openTimeoutSeconds);
	}

	PrintUsage();
	return EXIT_FAILURE;
}
