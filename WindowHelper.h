#pragma once

#include "Windows.h"
#include <string>

//����E���ڲ˵���
HWND FindMenuBar(HWND hParent);

//�����������
HWND FindOutputWindow(HWND hParent);


//��ȡ��ǰԴ�ļ���·�� ���ӱ�������
std::string GetSourceFilePath();


//���ڡ������¼���
void PeekAllMessage();