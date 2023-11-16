#pragma once
#ifndef __ELIBFNE__
#define __ELIBFNE__


#include <windows.h>
#include <tchar.h>
#include <cassert>
#include <lib2.h>
#include <lang.h>
#include <fnshare.h>
#include <fnshare.cpp>

#include <detver.h>
#include <detours.h>
#include <string>

#define LIBARAYNAME "AutoLinker_MessageNotify"


typedef DWORD(* OriginalFunctionWithDebugStart)(DWORD*, DWORD, DWORD);

typedef void (WINAPIV* OriginalFunctionWithBuildStart)();

typedef void(* OriginalCompilerFunction)();


void StartHookCreateFileA();



EXTERN_C INT WINAPI  AutoLinker_MessageNotify(INT nMsg, DWORD dwParam1, DWORD dwParam2);

#ifndef __E_STATIC_LIB
#define LIB_GUID_STR "EA968347-300C-4515-8888-F1D3BA3DF67E" /*GUID串*/
#define LIB_MajorVersion 1 /*库主版本号*/
#define LIB_MinorVersion 2 /*库次版本号*/
#define LIB_BuildNumber 20231009 /*构建版本号*/
#define LIB_SysMajorVer 3 /*系统主版本号*/
#define LIB_SysMinorVer 0 /*系统次版本号*/
#define LIB_KrnlLibMajorVer 3 /*核心库主版本号*/
#define LIB_KrnlLibMinorVer 0 /*核心库次版本号*/
#define LIB_NAME_STR "AutoLinker" /*支持库名*/
#define LIB_DESCRIPTION_STR "AutoLinker" /*功能描述*/
#define LIB_Author "" /*作者名称*/
#define LIB_ZipCode "" /*邮政编码*/
#define LIB_Address "" /*通信地址*/
#define LIB_Phone	"" /*电话号码*/
#define LIB_Fax		"" /*QQ号码*/
#define LIB_Email	 "" /*电子邮箱*/
#define LIB_HomePage "" /*主页地址*/
#define LIB_Other	"" /*其它信息*/
#define LIB_TYPE_COUNT 1 /*命令分类数量*/
#define LIB_TYPE_STR "0000基本命令\0""\0" /*命令分类*/
#endif

#endif
