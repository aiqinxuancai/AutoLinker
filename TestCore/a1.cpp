#include "pch.h"
#include "framework.h"
#include <math.h>
#include <string_view>
#include <string>
#include <algorithm>
#include <variant>


/*

��lib�����滻���Ŀ⺯��Ϊ�ִ�C++�����Ĳ��ԣ��������Ҫ��������������������ʵ��

1.��IDA�в�����ȷ�ĺ���ǩ��������
2.������ں��º��Ŀ�Ŀ�Դ�вο����������µ�ʵ��



ע�⣺
��Ҫ�ر�������/GL����������ᴥ���������ӵ��º�����AutoLinker������Lib˳��


*/


/// <summary>
/// λ������
/// </summary>
/// <param name="pRetData"></param>
/// <param name="nArgCount"></param>
/// <param name="pArgInf"></param>
/// <returns></returns>
extern "C" void _cdecl krnln_fnBAnd(PMDATA_INF pRetData, INT nArgCount, PMDATA_INF pArgInf) {
	int n = pArgInf->m_int;
	for (int i = 1; i < nArgCount; i++)
	{
		n &= pArgInf[i].m_int;
	}
	pRetData->m_int = n;
}


/// <summary>
/// ʹ���ִ�C++����ʵ��Ѱ���ı�����Լ�Ǻ��Ŀ��300%�ٶ�
/// </summary>
/// <param name="pRetData"></param>
/// <param name="nArgCount"></param>
/// <param name="pArgInf"></param>
/// <returns></returns>
extern "C" void __cdecl krnln_fnInStr(PMDATA_INF pRetData, INT nArgCount, PMDATA_INF pArgInf) {

    //TODO �㻹�������������VMP��־

    // ��ȡ�����ַ���
    std::string_view inputString = pArgInf[0].m_pText;
    std::string_view searchString = pArgInf[1].m_pText;

    // ���ַ������
    if (inputString.empty() || searchString.empty()) {
        pRetData->m_int = -1;
        return;
    }

    // ȷ����ʼ������λ��
    size_t searchStartPos = (pArgInf[2].m_dtDataType == _SDT_NULL || pArgInf[2].m_int <= 1) ? 0 : pArgInf[2].m_int - 1;

    // ����ָ���ַ���
    auto searchResult = (pArgInf[3].m_bool) ?
        std::search(
            inputString.begin() + searchStartPos, inputString.end(),
            searchString.begin(), searchString.end(),
            [](char c1, char c2) { return std::tolower(c1) == std::tolower(c2); }
        ) :
        std::search(
            inputString.begin() + searchStartPos, inputString.end(),
            searchString.begin(), searchString.end()
        );

    // ���÷���λ�ã�����ҵ��򷵻�λ�ã����򷵻�-1
    pRetData->m_int = (searchResult != inputString.end()) ? std::distance(inputString.begin(), searchResult) + 1 : -1;
}
