#pragma once

#include <cstddef>
#include <string>
#include <vector>

namespace e571 {

// 当前工程二进制序列化器。
class ProjectBinarySerializer {
public:
	static ProjectBinarySerializer& Instance();

	// 记录经 IDE 真实保存路径验证过的序列化对象上下文。
	void RecordVerifiedSerializerContext(
		void* serializerThis,
		const std::string& sourcePath = std::string());

	// 清空已记录的序列化对象上下文。
	void ClearVerifiedSerializerContext();

	// 注册 Detours 交易完成后的 IDE 原始文件序列化入口。
	void ConfigureFileSerializer(void* serializeToFileFunction);

	bool SerializeCurrentProject(
		std::vector<unsigned char>& outBytes,
		std::string* outError = nullptr,
		std::string* outTrace = nullptr);

	bool WriteCurrentProjectToFile(
		const std::string& outputPath,
		size_t* outBytesWritten = nullptr,
		std::string* outError = nullptr,
		std::string* outTrace = nullptr);

	// 按 IDE 完整 .e 文件格式写入快照，并恢复当前工程的原路径。
	bool WriteCurrentProjectFileSnapshot(
		const std::string& outputPath,
		size_t* outBytesWritten = nullptr,
		std::string* outError = nullptr,
		std::string* outTrace = nullptr);

private:
	ProjectBinarySerializer() = default;
	ProjectBinarySerializer(const ProjectBinarySerializer&) = delete;
	ProjectBinarySerializer& operator=(const ProjectBinarySerializer&) = delete;
};

} // namespace e571
