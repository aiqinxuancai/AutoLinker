// global.h
#pragma once
#ifndef GLOBAL_H
#define GLOBAL_H

//���Կ�ʼ��ַ
extern int g_debugStartAddress;

//���뿪ʼ��ַ
extern int g_compileStartAddress;

extern HWND g_hwnd;

/// <summary>
/// ����ı�
/// </summary>
/// <param name="szbuf"></param>
void OutputStringToELog(std::string szbuf);


/// <summary>
/// ���в��֪ͨ
/// </summary>
/// <param name="code"></param>
/// <param name="p1"></param>
/// <param name="p2"></param>
/// <returns></returns>
INT NESRUNFUNC(INT code, DWORD p1, DWORD p2);

#endif // GLOBAL_H