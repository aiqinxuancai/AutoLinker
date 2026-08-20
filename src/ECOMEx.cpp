#include "ECOMEx.h"
#include "IDEFacade.h"
#include "Global.h"
#include <filesystem>
#include <vector>

int GetECOMCount() {
	int ecomCount = 0;
	IDEFacade::Instance().GetImportedECOMCount(ecomCount);
	return ecomCount;
}

std::string GetECOMPath(int index) {
	std::string path;
	IDEFacade::Instance().GetImportedECOMPath(index, path);
	return path;
}

BOOL AddECOM(std::string filePath) {
	return IDEFacade::Instance().AddECOM(filePath) ? TRUE : FALSE;
}

BOOL AddECOM2(std::string filePath) {
	return IDEFacade::Instance().AddECOM2(filePath) ? TRUE : FALSE;
}

BOOL RemoveECOM(int index) {
	return IDEFacade::Instance().RemoveECOM(index) ? TRUE : FALSE;
}

BOOL RemoveECOM(std::string filePah) {
	return IDEFacade::Instance().RemoveECOM(filePah) ? TRUE : FALSE;
}

int FindECOMIndex(std::string filePah) {
	return IDEFacade::Instance().FindECOMIndex(filePah);
}

/// <summary>
/// 使用模块名查找模块
/// </summary>
/// <param name="ecomName"></param>
/// <returns></returns>
int FindECOMNameIndex(std::string ecomName) {
	return IDEFacade::Instance().FindECOMNameIndex(ecomName);
}


/// <summary>
/// 根据当前阶段自动切换模块
/// </summary>
/// <param name="useCompileModules">true 切到静态编译版，false 切回动态调试版。</param>
int RunChangeECOM(bool useCompileModules) {
	struct Replacement {
		std::string oldPath;
		std::string newPath;
	};

	std::vector<Replacement> replacements;
	const int ecomCount = GetECOMCount();
	for (int i = 0; i < ecomCount; ++i) {
		std::string item = GetECOMPath(i);
		std::filesystem::path pathObj(item);
		std::string fileName = pathObj.filename().string();
		std::string needChangECOMName;

		if (useCompileModules) {
			needChangECOMName = g_modelManager.getValue(fileName);
		}
		else {
			needChangECOMName = g_modelManager.getKeyFromValue(fileName);
		}

		if (!needChangECOMName.empty()) {
			replacements.push_back({
				item,
				(pathObj.parent_path() / needChangECOMName).string()
			});
		}
	}

	int changedCount = 0;
	for (const auto& replacement : replacements) {
		if (!RemoveECOM(replacement.oldPath)) {
			OutputStringToELog("切换模块失败，无法移除: " + replacement.oldPath);
			continue;
		}
		if (AddECOM2(replacement.newPath)) {
			++changedCount;
			OutputStringToELog("切换模块:" + replacement.oldPath + " -> " + replacement.newPath);
			continue;
		}

		OutputStringToELog("切换模块失败，正在恢复: " + replacement.newPath);
		if (!AddECOM2(replacement.oldPath)) {
			OutputStringToELog("恢复原模块失败: " + replacement.oldPath);
		}
	}
	return changedCount;
}
