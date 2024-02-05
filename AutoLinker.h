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
#include <json.hpp>


using json = nlohmann::json;

#define LIBARAYNAME "AutoLinker_MessageNotify"

#define AUTOLINKER_VERSION "0.0.0"

typedef DWORD(* OriginalFunctionWithDebugStart)(DWORD*, DWORD, DWORD);

typedef void (WINAPIV* OriginalFunctionWithBuildStart)();

typedef void(* OriginalCompilerFunction)();

void StartHookCreateFileA();

void ChangeVMProtectModel(bool isLib);

void UpdateCurrentOpenSourceFile();

EXTERN_C INT WINAPI  AutoLinker_MessageNotify(INT nMsg, DWORD dwParam1, DWORD dwParam2);


#ifndef __E_STATIC_LIB
#define LIB_GUID_STR "EA968347-300C-4515-8888-F1D3BA3DF67E" /*GUID��*/
#define LIB_MajorVersion 1 /*�����汾��*/
#define LIB_MinorVersion 0 /*��ΰ汾��*/
#define LIB_BuildNumber 1000000 /*�����汾��*/
#define LIB_SysMajorVer 3 /*ϵͳ���汾��*/
#define LIB_SysMinorVer 7 /*ϵͳ�ΰ汾��*/
#define LIB_KrnlLibMajorVer 5 /*���Ŀ����汾��*/
#define LIB_KrnlLibMinorVer 3 /*���Ŀ�ΰ汾��*/
#define LIB_NAME_STR "AutoLinker" /*֧�ֿ���*/
#define LIB_DESCRIPTION_STR "AutoLinker 0.0.0.1000000��ʵ��ָ����.e�ļ��������������뼰����ʱ.ec�ɶԵ��Զ��л�����д���Ŀ⺯���ȹ��ܡ���Դ��ַ��https://github.com/aiqinxuancai/AutoLinker" /*��������*/
#define LIB_Author "aiqinxuancai" /*��������*/
#define LIB_ZipCode "" /*��������*/
#define LIB_Address "" /*ͨ�ŵ�ַ*/
#define LIB_Phone	"" /*�绰����*/
#define LIB_Fax		"" /*QQ����*/
#define LIB_Email	 "" /*��������*/
#define LIB_HomePage "https://github.com/aiqinxuancai" /*��ҳ��ַ*/
#define LIB_Other	"" /*������Ϣ*/
#define LIB_TYPE_COUNT 1 /*�����������*/
#define LIB_TYPE_STR "0000��������\0""\0" /*�������*/
#endif

#endif
