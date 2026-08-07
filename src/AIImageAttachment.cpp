#include "AIImageAttachment.h"

#include <Windows.h>
#include <bcrypt.h>
#include <wincodec.h>
#include <wrl/client.h>

#include <algorithm>
#include <array>
#include <cctype>
#include <cwctype>
#include <fstream>
#include <iomanip>
#include <iterator>
#include <sstream>

#include "UnicodeTextCodec.h"
#include "..\thirdparty\json.hpp"

#pragma comment(lib, "bcrypt.lib")
#pragma comment(lib, "windowscodecs.lib")

namespace {

using Microsoft::WRL::ComPtr;

std::string Utf8ToLocal(const std::string& text)
{
	return UnicodeTextCodec::Utf8ToLocalPreservingUnicode(text);
}

std::string WideToLocal(const std::wstring& text)
{
	if (text.empty()) {
		return {};
	}
	const int length = WideCharToMultiByte(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
	if (length <= 0) {
		return {};
	}
	std::string utf8(static_cast<size_t>(length), '\0');
	if (WideCharToMultiByte(
			CP_UTF8,
			0,
			text.data(),
			static_cast<int>(text.size()),
			utf8.data(),
			length,
			nullptr,
			nullptr) <= 0) {
		return {};
	}
	return UnicodeTextCodec::Utf8ToLocalPreservingUnicode(utf8);
}

std::filesystem::path LocalTextToPath(const std::string& text)
{
	const std::string utf8 = UnicodeTextCodec::LocalToUtf8RestoringUnicode(text);
	if (utf8.empty()) {
		return {};
	}
	const int length = MultiByteToWideChar(
		CP_UTF8,
		MB_ERR_INVALID_CHARS,
		utf8.data(),
		static_cast<int>(utf8.size()),
		nullptr,
		0);
	if (length <= 0) {
		return std::filesystem::path(text);
	}
	std::wstring wide(static_cast<size_t>(length), L'\0');
	if (MultiByteToWideChar(
			CP_UTF8,
			MB_ERR_INVALID_CHARS,
			utf8.data(),
			static_cast<int>(utf8.size()),
			wide.data(),
			length) <= 0) {
		return std::filesystem::path(text);
	}
	return std::filesystem::path(wide);
}

std::string NormalizeDetail(std::string detail)
{
	std::transform(detail.begin(), detail.end(), detail.begin(), [](unsigned char ch) {
		return static_cast<char>(std::tolower(ch));
	});
	return detail == "low" || detail == "high" ? detail : "auto";
}

std::string Base64Encode(const std::vector<unsigned char>& bytes)
{
	static constexpr char kAlphabet[] =
		"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
	std::string output;
	output.reserve(((bytes.size() + 2) / 3) * 4);
	for (size_t index = 0; index < bytes.size(); index += 3) {
		const unsigned int value =
			(static_cast<unsigned int>(bytes[index]) << 16) |
			(index + 1 < bytes.size() ? static_cast<unsigned int>(bytes[index + 1]) << 8 : 0) |
			(index + 2 < bytes.size() ? static_cast<unsigned int>(bytes[index + 2]) : 0);
		output.push_back(kAlphabet[(value >> 18) & 0x3F]);
		output.push_back(kAlphabet[(value >> 12) & 0x3F]);
		output.push_back(index + 1 < bytes.size() ? kAlphabet[(value >> 6) & 0x3F] : '=');
		output.push_back(index + 2 < bytes.size() ? kAlphabet[value & 0x3F] : '=');
	}
	return output;
}

int Base64Value(unsigned char ch)
{
	if (ch >= 'A' && ch <= 'Z') return ch - 'A';
	if (ch >= 'a' && ch <= 'z') return ch - 'a' + 26;
	if (ch >= '0' && ch <= '9') return ch - '0' + 52;
	if (ch == '+') return 62;
	if (ch == '/') return 63;
	return -1;
}

bool Base64Decode(const std::string& encoded, std::vector<unsigned char>& output)
{
	output.clear();
	unsigned int accumulator = 0;
	int bits = -8;
	for (unsigned char ch : encoded) {
		if (ch == '=') {
			break;
		}
		if (std::isspace(ch)) {
			continue;
		}
		const int value = Base64Value(ch);
		if (value < 0) {
			output.clear();
			return false;
		}
		accumulator = (accumulator << 6) | value;
		bits += 6;
		if (bits >= 0) {
			output.push_back(static_cast<unsigned char>((accumulator >> bits) & 0xFF));
			bits -= 8;
		}
	}
	return !output.empty();
}

bool ReadAllBytes(const std::filesystem::path& path, std::vector<unsigned char>& output)
{
	std::ifstream stream(path, std::ios::binary);
	if (!stream.is_open()) {
		return false;
	}
	output.assign(std::istreambuf_iterator<char>(stream), std::istreambuf_iterator<char>());
	return stream.good() || stream.eof();
}

bool WriteAllBytes(const std::filesystem::path& path, const std::vector<unsigned char>& bytes)
{
	std::ofstream stream(path, std::ios::binary | std::ios::trunc);
	if (!stream.is_open()) {
		return false;
	}
	stream.write(reinterpret_cast<const char*>(bytes.data()), static_cast<std::streamsize>(bytes.size()));
	return stream.good();
}

std::string Sha256Hex(const std::vector<unsigned char>& bytes)
{
	BCRYPT_ALG_HANDLE algorithm = nullptr;
	BCRYPT_HASH_HANDLE hash = nullptr;
	DWORD objectSize = 0;
	DWORD copied = 0;
	std::vector<unsigned char> object;
	std::array<unsigned char, 32> digest = {};
	if (BCryptOpenAlgorithmProvider(&algorithm, BCRYPT_SHA256_ALGORITHM, nullptr, 0) < 0 ||
		BCryptGetProperty(algorithm, BCRYPT_OBJECT_LENGTH, reinterpret_cast<PUCHAR>(&objectSize), sizeof(objectSize), &copied, 0) < 0) {
		if (algorithm != nullptr) BCryptCloseAlgorithmProvider(algorithm, 0);
		return {};
	}
	object.resize(objectSize);
	const NTSTATUS createStatus = BCryptCreateHash(algorithm, &hash, object.data(), objectSize, nullptr, 0, 0);
	const NTSTATUS dataStatus = createStatus >= 0
		? BCryptHashData(hash, const_cast<PUCHAR>(bytes.data()), static_cast<ULONG>(bytes.size()), 0)
		: createStatus;
	const NTSTATUS finishStatus = dataStatus >= 0
		? BCryptFinishHash(hash, digest.data(), static_cast<ULONG>(digest.size()), 0)
		: dataStatus;
	if (hash != nullptr) BCryptDestroyHash(hash);
	BCryptCloseAlgorithmProvider(algorithm, 0);
	if (finishStatus < 0) {
		return {};
	}
	std::ostringstream text;
	text << std::hex << std::setfill('0');
	for (unsigned char value : digest) {
		text << std::setw(2) << static_cast<unsigned int>(value);
	}
	return text.str();
}

class ComApartmentScope {
public:
	ComApartmentScope()
		: result_(CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED))
	{
	}

	~ComApartmentScope()
	{
		if (SUCCEEDED(result_)) {
			CoUninitialize();
		}
	}

	bool IsReady() const
	{
		return SUCCEEDED(result_) || result_ == RPC_E_CHANGED_MODE;
	}

private:
	HRESULT result_ = E_FAIL;
};

bool CreateWicFactory(ComPtr<IWICImagingFactory>& factory)
{
	return SUCCEEDED(CoCreateInstance(
		CLSID_WICImagingFactory,
		nullptr,
		CLSCTX_INPROC_SERVER,
		IID_PPV_ARGS(&factory)));
}

bool EncodeWicSource(
	IWICImagingFactory* factory,
	IWICBitmapSource* source,
	UINT width,
	UINT height,
	const std::filesystem::path& destination,
	const GUID& container,
	const WICPixelFormatGUID& pixelFormat,
	float quality)
{
	ComPtr<IWICFormatConverter> converter;
	if (FAILED(factory->CreateFormatConverter(&converter)) ||
		FAILED(converter->Initialize(source, pixelFormat, WICBitmapDitherTypeNone, nullptr, 0.0, WICBitmapPaletteTypeCustom))) {
		return false;
	}
	ComPtr<IWICStream> stream;
	ComPtr<IWICBitmapEncoder> encoder;
	ComPtr<IWICBitmapFrameEncode> frame;
	ComPtr<IPropertyBag2> properties;
	if (FAILED(factory->CreateStream(&stream)) ||
		FAILED(stream->InitializeFromFilename(destination.c_str(), GENERIC_WRITE)) ||
		FAILED(factory->CreateEncoder(container, nullptr, &encoder)) ||
		FAILED(encoder->Initialize(stream.Get(), WICBitmapEncoderNoCache)) ||
		FAILED(encoder->CreateNewFrame(&frame, &properties))) {
		return false;
	}
	if (properties != nullptr && IsEqualGUID(container, GUID_ContainerFormatJpeg)) {
		PROPBAG2 option = {};
		option.pstrName = const_cast<LPOLESTR>(L"ImageQuality");
		VARIANT value;
		VariantInit(&value);
		value.vt = VT_R4;
		value.fltVal = quality;
		properties->Write(1, &option, &value);
		VariantClear(&value);
	}
	WICPixelFormatGUID targetFormat = pixelFormat;
	return SUCCEEDED(frame->Initialize(properties.Get())) &&
		SUCCEEDED(frame->SetSize(width, height)) &&
		SUCCEEDED(frame->SetPixelFormat(&targetFormat)) &&
		SUCCEEDED(frame->WriteSource(converter.Get(), nullptr)) &&
		SUCCEEDED(frame->Commit()) &&
		SUCCEEDED(encoder->Commit());
}

AIImagePrepareResult PreparePathInternal(
	const std::filesystem::path& sourcePath,
	const std::filesystem::path& assetDirectory,
	const std::string& displayNameLocal,
	const std::string& originalSourceLocal,
	const std::string& detail)
{
	AIImagePrepareResult result;
	ComApartmentScope apartment;
	if (!apartment.IsReady()) {
		result.errorLocal = "初始化 Windows 图像组件失败。";
		return result;
	}
	std::error_code ec;
	if (!std::filesystem::is_regular_file(sourcePath, ec)) {
		result.errorLocal = "图片文件不存在或不是普通文件。";
		return result;
	}
	const std::uint64_t inputSize = std::filesystem::file_size(sourcePath, ec);
	if (ec || inputSize == 0 || inputSize > AIImageAttachmentManager::kMaxInputBytes) {
		result.errorLocal = "图片为空或超过 20 MB 输入限制。";
		return result;
	}

	std::vector<unsigned char> sourceBytes;
	if (!ReadAllBytes(sourcePath, sourceBytes)) {
		result.errorLocal = "无法读取图片文件。";
		return result;
	}
	const std::string hash = Sha256Hex(sourceBytes);
	if (hash.empty()) {
		result.errorLocal = "计算图片哈希失败。";
		return result;
	}

	ComPtr<IWICImagingFactory> factory;
	ComPtr<IWICBitmapDecoder> decoder;
	ComPtr<IWICBitmapFrameDecode> frame;
	if (!CreateWicFactory(factory) ||
		FAILED(factory->CreateDecoderFromFilename(sourcePath.c_str(), nullptr, GENERIC_READ, WICDecodeMetadataCacheOnLoad, &decoder)) ||
		FAILED(decoder->GetFrame(0, &frame))) {
		result.errorLocal = "图片格式无法由 Windows 图像组件解码。";
		return result;
	}
	UINT width = 0;
	UINT height = 0;
	if (FAILED(frame->GetSize(&width, &height)) || width == 0 || height == 0) {
		result.errorLocal = "无法读取图片尺寸。";
		return result;
	}

	UINT outputWidth = width;
	UINT outputHeight = height;
	ComPtr<IWICBitmapSource> prepared;
	if (FAILED(frame.As(&prepared))) {
		result.errorLocal = "无法读取图片像素。";
		return result;
	}
	if (width > AIImageAttachmentManager::kMaxDimension || height > AIImageAttachmentManager::kMaxDimension) {
		const double scale = (std::min)(
			static_cast<double>(AIImageAttachmentManager::kMaxDimension) / width,
			static_cast<double>(AIImageAttachmentManager::kMaxDimension) / height);
		outputWidth = (std::max)(1u, static_cast<UINT>(width * scale));
		outputHeight = (std::max)(1u, static_cast<UINT>(height * scale));
		ComPtr<IWICBitmapScaler> scaler;
		if (FAILED(factory->CreateBitmapScaler(&scaler)) ||
			FAILED(scaler->Initialize(frame.Get(), outputWidth, outputHeight, WICBitmapInterpolationModeFant))) {
			result.errorLocal = "缩放图片失败。";
			return result;
		}
		if (FAILED(scaler.As(&prepared))) {
			result.errorLocal = "无法读取缩放后的图片像素。";
			return result;
		}
	}

	std::filesystem::create_directories(assetDirectory, ec);
	if (ec) {
		result.errorLocal = "无法创建会话图片资源目录。";
		return result;
	}
	std::filesystem::path outputPath = assetDirectory / (std::wstring(hash.begin(), hash.end()) + L".png");
	if (!std::filesystem::exists(outputPath, ec) &&
		!EncodeWicSource(factory.Get(), prepared.Get(), outputWidth, outputHeight, outputPath,
			GUID_ContainerFormatPng, GUID_WICPixelFormat32bppBGRA, 1.0f)) {
		result.errorLocal = "编码 PNG 图片失败。";
		return result;
	}
	std::uint64_t outputSize = std::filesystem::file_size(outputPath, ec);
	std::string mimeType = "image/png";
	if (ec || outputSize > AIImageAttachmentManager::kMaxPreparedBytes) {
		std::filesystem::path jpegPath = assetDirectory / (std::wstring(hash.begin(), hash.end()) + L".jpg");
		if (!std::filesystem::exists(jpegPath, ec) &&
			!EncodeWicSource(factory.Get(), prepared.Get(), outputWidth, outputHeight, jpegPath,
				GUID_ContainerFormatJpeg, GUID_WICPixelFormat24bppBGR, 0.86f)) {
			result.errorLocal = "图片超过通用大小限制，JPEG 压缩失败。";
			return result;
		}
		outputPath = jpegPath;
		mimeType = "image/jpeg";
		outputSize = std::filesystem::file_size(outputPath, ec);
	}
	if (ec || outputSize == 0 || outputSize > AIImageAttachmentManager::kMaxPreparedBytes) {
		result.errorLocal = "处理后的图片仍超过 4 MB 限制。";
		return result;
	}

	result.ok = true;
	result.attachment.id = hash.substr(0, 16);
	result.attachment.fileNameLocal = displayNameLocal.empty() ? WideToLocal(sourcePath.filename().wstring()) : displayNameLocal;
	result.attachment.mimeType = mimeType;
	result.attachment.assetPathLocal = WideToLocal(outputPath.wstring());
	result.attachment.sourcePathLocal = originalSourceLocal;
	result.attachment.detail = NormalizeDetail(detail);
	result.attachment.byteSize = outputSize;
	result.attachment.width = outputWidth;
	result.attachment.height = outputHeight;
	return result;
}

} // namespace

AIImagePrepareResult AIImageAttachmentManager::PrepareFile(
	const std::filesystem::path& sourcePath,
	const std::filesystem::path& assetDirectory,
	const std::string& detail)
{
	std::error_code ec;
	const std::filesystem::path absolute = std::filesystem::absolute(sourcePath, ec).lexically_normal();
	const std::filesystem::path effective = ec ? sourcePath : absolute;
	return PreparePathInternal(
		effective,
		assetDirectory,
		WideToLocal(effective.filename().wstring()),
		WideToLocal(effective.wstring()),
		detail);
}

AIImagePrepareResult AIImageAttachmentManager::PrepareDataUrl(
	const std::string& dataUrlUtf8,
	const std::string& fileNameUtf8,
	const std::filesystem::path& assetDirectory,
	const std::string& detail)
{
	AIImagePrepareResult result;
	const size_t separator = dataUrlUtf8.find(',');
	if (separator == std::string::npos || dataUrlUtf8.rfind("data:image/", 0) != 0 ||
		dataUrlUtf8.substr(0, separator).find(";base64") == std::string::npos) {
		result.errorLocal = "附件不是有效的 Base64 图片 data URL。";
		return result;
	}
	std::vector<unsigned char> bytes;
	if (!Base64Decode(dataUrlUtf8.substr(separator + 1), bytes) || bytes.size() > kMaxInputBytes) {
		result.errorLocal = "图片 Base64 数据无效或超过 20 MB。";
		return result;
	}
	std::error_code ec;
	std::filesystem::create_directories(assetDirectory, ec);
	if (ec) {
		result.errorLocal = "无法创建会话图片资源目录。";
		return result;
	}
	const std::string hash = Sha256Hex(bytes);
	if (hash.empty()) {
		result.errorLocal = "计算图片哈希失败。";
		return result;
	}
	const std::filesystem::path temporary = assetDirectory / (std::wstring(hash.begin(), hash.end()) + L".input");
	if (!WriteAllBytes(temporary, bytes)) {
		result.errorLocal = "无法写入图片临时快照。";
		return result;
	}
	result = PreparePathInternal(
		temporary,
		assetDirectory,
		Utf8ToLocal(fileNameUtf8),
		{},
		detail);
	std::filesystem::remove(temporary, ec);
	return result;
}

bool AIImageAttachmentManager::BuildDataUrl(
	const AIImageAttachment& attachment,
	std::string& outDataUrlUtf8,
	std::string& outErrorLocal)
{
	outDataUrlUtf8.clear();
	outErrorLocal.clear();
	std::vector<unsigned char> bytes;
	if (attachment.assetPathLocal.empty() ||
		!ReadAllBytes(LocalTextToPath(attachment.assetPathLocal), bytes)) {
		outErrorLocal = "图片快照不存在或无法读取。";
		return false;
	}
	if (bytes.empty() || bytes.size() > kMaxPreparedBytes) {
		outErrorLocal = "图片快照为空或超过 4 MB。";
		return false;
	}
	const std::string mime = attachment.mimeType.empty() ? "image/png" : attachment.mimeType;
	outDataUrlUtf8 = "data:" + mime + ";base64," + Base64Encode(bytes);
	return true;
}

bool AIImageAttachmentManager::IsSupportedImagePath(const std::filesystem::path& path)
{
	std::wstring extension = path.extension().wstring();
	std::transform(extension.begin(), extension.end(), extension.begin(), [](wchar_t ch) {
		return static_cast<wchar_t>(std::towlower(ch));
	});
	return extension == L".png" || extension == L".jpg" || extension == L".jpeg" ||
		extension == L".gif" || extension == L".bmp" || extension == L".webp" ||
		extension == L".tif" || extension == L".tiff";
}

std::string AIImageAttachmentManager::BuildSelfTestJson()
{
	nlohmann::json report = {
		{"name", "ai-image-attachment"},
		{"ok", false},
		{"max_dimension", kMaxDimension},
		{"max_attachment_count", kMaxAttachmentCount},
		{"checks", nlohmann::json::object()}
	};
	std::vector<unsigned char> sample = {0x00, 0x01, 0x02, 0xFD, 0xFE, 0xFF};
	const std::string encoded = Base64Encode(sample);
	std::vector<unsigned char> decoded;
	report["checks"]["base64_roundtrip"] = Base64Decode(encoded, decoded) && decoded == sample;
	report["checks"]["extension_detection"] =
		IsSupportedImagePath(L"test.PNG") && !IsSupportedImagePath(L"test.txt");

	const std::filesystem::path tempRoot = std::filesystem::temp_directory_path() /
		(L"AutoLinker-\u56fe\u7247\u81ea\u68c0-" + std::to_wstring(GetCurrentProcessId()) + L"-" + std::to_wstring(GetTickCount64()));
	const std::filesystem::path sourcePath = tempRoot / L"\u6d4b\u8bd5\u56fe\u7247.png";
	const std::filesystem::path assetDirectory = tempRoot / L"\u4f1a\u8bdd\u8d44\u6e90";
	std::error_code ec;
	std::filesystem::create_directories(tempRoot, ec);
	std::vector<unsigned char> pngBytes;
	const bool pngDecoded = Base64Decode(
		"iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAF/gL+X2iXWQAAAABJRU5ErkJggg==",
		pngBytes);
	const bool sourceWritten = !ec && pngDecoded && WriteAllBytes(sourcePath, pngBytes);
	const AIImagePrepareResult prepared = sourceWritten
		? PrepareFile(sourcePath, assetDirectory, "high")
		: AIImagePrepareResult{};
	const AIImagePrepareResult duplicate = prepared.ok
		? PrepareFile(sourcePath, assetDirectory, "low")
		: AIImagePrepareResult{};
	std::string dataUrl;
	std::string dataUrlError;
	const bool dataUrlOk = prepared.ok && BuildDataUrl(prepared.attachment, dataUrl, dataUrlError);
	AIImageAttachment missing = prepared.attachment;
	missing.assetPathLocal = WideToLocal((tempRoot / L"missing.png").wstring());
	std::string missingDataUrl;
	std::string missingError;
	const bool missingRejected = !BuildDataUrl(missing, missingDataUrl, missingError) && !missingError.empty();

	report["checks"]["unicode_path_wic"] = prepared.ok &&
		prepared.attachment.width == 1 && prepared.attachment.height == 1;
	report["checks"]["snapshot_data_url"] = dataUrlOk && dataUrl.rfind("data:image/", 0) == 0;
	report["checks"]["content_deduplication"] = duplicate.ok &&
		duplicate.attachment.assetPathLocal == prepared.attachment.assetPathLocal;
	report["checks"]["missing_snapshot_rejected"] = missingRejected;
	report["ok"] = report["checks"]["base64_roundtrip"].get<bool>() &&
		report["checks"]["extension_detection"].get<bool>() &&
		report["checks"]["unicode_path_wic"].get<bool>() &&
		report["checks"]["snapshot_data_url"].get<bool>() &&
		report["checks"]["content_deduplication"].get<bool>() &&
		report["checks"]["missing_snapshot_rejected"].get<bool>();
	std::filesystem::remove_all(tempRoot, ec);
	return report.dump();
}
