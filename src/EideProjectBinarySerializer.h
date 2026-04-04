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

	bool SerializeCurrentProject(
		std::vector<unsigned char>& outBytes,
		std::string* outError = nullptr,
		std::string* outTrace = nullptr);

	bool WriteCurrentProjectToFile(
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
