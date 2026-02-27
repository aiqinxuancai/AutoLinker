#include "ECOMEx.h"
#include "IDEFacade.h"
#include "Global.h"
#include <filesystem>

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
/// 根据当前是否是调试，自动切换模块
/// </summary>
/// <param name="isDebug"></param>
void RunChangeECOM(bool isDebug) {
	int ecomCount = GetECOMCount();
	for (int i = 0; i < ecomCount; i++) {
		std::string item = GetECOMPath(i);
		std::filesystem::path pathObj(item);
		std::string fileName = pathObj.filename().string();
		std::string needChangECOMName;

		if (isDebug) {
			needChangECOMName = g_modelManager.getValue(fileName);
		}
		else {
			needChangECOMName = g_modelManager.getKeyFromValue(fileName);
		}

		if (!needChangECOMName.empty()) {
			auto newPath = pathObj.parent_path().append(needChangECOMName).string();
			if (RemoveECOM(i) && AddECOM2(newPath)) {
				OutputStringToELog("切换模块:" + item + " -> " + newPath);
			}
		}
	}
}
