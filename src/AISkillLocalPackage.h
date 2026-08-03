#pragma once

#include <filesystem>
#include <string>

// 本地 AI 技能来源类型。
enum class AISkillLocalSourceKind {
	Directory,
	SkillFile,
	ZipArchive
};

// 已准备好的本地 AI 技能来源；ZIP 来源会持有临时解压目录。
struct AISkillPreparedLocalSource {
	AISkillLocalSourceKind kind = AISkillLocalSourceKind::Directory;
	std::filesystem::path sourcePath;
	std::filesystem::path scanRoot;
	std::filesystem::path directSkillFile;
	std::filesystem::path stagingRoot;
};

// 本地 AI 技能目录和 ZIP 包准备工具。
namespace AISkillLocalPackage {

// 校验来源，并在需要时将 ZIP 安全解压到缓存目录。
bool Prepare(
	const std::filesystem::path& sourcePath,
	const std::filesystem::path& cacheRoot,
	AISkillPreparedLocalSource& outSource,
	std::string& outError);

// 清理 Prepare 创建的临时目录。
void Cleanup(AISkillPreparedLocalSource& source, const std::filesystem::path& cacheRoot);

// 返回适合界面协议使用的来源类型名称。
const char* KindName(AISkillLocalSourceKind kind);

} // namespace AISkillLocalPackage
