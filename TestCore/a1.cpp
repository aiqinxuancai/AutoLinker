#include "pch.h"
#include "framework.h"
#include <math.h>
#include <string_view>
#include <string>
#include <algorithm>

//LIBAPI(int, krnln_fnBAnd)
//{
//	PMDATA_INF pArg = &ArgInf;
//	ret
//	int n = pArg->m_int;
//	for (int i = 1; i < nArgCount; i++)
//	{
//		n &= pArg[i].m_int;
//	}
//
//	return n;
//
//}


extern "C" void _cdecl krnln_fnBAnd(PMDATA_INF pRetData, INT nArgCount, PMDATA_INF pArgInf) {
	int n = pArgInf->m_int;
	for (int i = 1; i < nArgCount; i++)
	{
		n &= pArgInf[i].m_int;
	}
	pRetData->m_int = n;
}


extern "C" void __cdecl krnln_fnInStr(PMDATA_INF pRetData, INT nArgCount, PMDATA_INF pArgInf) {
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
