#include "ECOMEx.h"
#include "Global.h"
#include <PublicIDEFunctions.h>
#include <filesystem>

int GetEModelCount() {
	int ecomCount;
	int e2 = NESRUNFUNC(FN_GET_NUM_ECOM, (DWORD)(&ecomCount), 0);
	return ecomCount;
}

std::string GetEModelPath(int index) {
	char path[MAX_PATH];
	NESRUNFUNC(FN_GET_ECOM_FILE_NAME, index, (DWORD)path);
	return path;
}

BOOL AddEModel(std::string filePath) {
	BOOL isSuccess;
	NESRUNFUNC(FN_INPUT_ECOM, (DWORD)filePath.c_str(), (DWORD)&isSuccess);
	return isSuccess;
}

BOOL AddEModel2(std::string filePath) {
	BOOL isSuccess;
	NESRUNFUNC(FN_ADD_NEW_ECOM2, (DWORD)filePath.c_str(), (DWORD)&isSuccess);
	return isSuccess;
}



BOOL RemoveEModel(int index) {
	char path[MAX_PATH];
	BOOL isSuccess;
	isSuccess = NESRUNFUNC(FN_REMOVE_SPEC_ECOM, index, 0);
	return isSuccess;
}

BOOL RemoveEModel(std::string filePah) {
	char path[MAX_PATH];
	BOOL isSuccess;
	int index = FindEModelIndex(filePah);
	if (index != -1) {
		NESRUNFUNC(FN_REMOVE_SPEC_ECOM, index, 0);
		return true;
	}

	return false;
}

int FindEModelIndex(std::string filePah) {
	char path[MAX_PATH];
	BOOL isSuccess;
	int ecomCount = GetEModelCount();
	for (int i = 0; i < ecomCount; i++) {
		std::string item = GetEModelPath(i);
		if (item == filePah) {
			return i;
		}
	}
	return -1;
}

/// <summary>
/// 使用模块名查找模块
/// </summary>
/// <param name="ecomName"></param>
/// <returns></returns>
int FindEModelNameIndex(std::string ecomName) {
	char path[MAX_PATH];
	BOOL isSuccess;
	int ecomCount = GetEModelCount();
	for (int i = 0; i < ecomCount; i++) {
		std::string item = GetEModelPath(i);
		std::filesystem::path pathObj(item);
		std::string fileNameWithoutExtension = pathObj.stem().string();
		if (fileNameWithoutExtension == ecomName) {
			return i;
		}
	}
	return -1;
}


/// <summary>
/// 根据当前是否是调试，自动切换模块
/// </summary>
/// <param name="isDebug"></param>
void RunChangeECOM(bool isDebug) {
	int ecomCount = GetEModelCount();
	for (int i = 0; i < ecomCount; i++) {
		std::string item = GetEModelPath(i);
		std::filesystem::path pathObj(item);
		std::string fileName = pathObj.filename().string();
		std::string needChangeModelName;

		if (isDebug) {
			needChangeModelName = g_modelManager.getValue(fileName);
		}
		else {
			needChangeModelName = g_modelManager.getKeyFromValue(fileName);
		}

		if (!needChangeModelName.empty()) {
			auto newPath = pathObj.parent_path().append(needChangeModelName).string();
			RemoveEModel(i);
			AddEModel2(newPath);
			OutputStringToELog("切换模块:" + item + " -> " + newPath);
		}
	}
}



