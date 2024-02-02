#include "pch.h"
#include "framework.h"
#include <math.h>
#include <string_view>
#include <string>
#include <algorithm>
#include <variant>


/*

此lib用于替换核心库函数为现代C++方法的测试，如果你想要增加其他函数，请自行实现

1.在IDA中查找正确的函数签名并声明
2.你可以在黑月核心库的开源中参考其代码进行新的实现



注意：
需要关闭链接器/GL参数，否则会触发二次链接导致忽略由AutoLinker调整的Lib顺序


*/


/// <summary>
/// 位与运算
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
/// 使用现代C++方法实现寻找文本，大约是核心库的300%速度
/// </summary>
/// <param name="pRetData"></param>
/// <param name="nArgCount"></param>
/// <param name="pArgInf"></param>
/// <returns></returns>
extern "C" void __cdecl krnln_fnInStr(PMDATA_INF pRetData, INT nArgCount, PMDATA_INF pArgInf) {

    //TODO 你还可以在这里添加VMP标志

    // 获取输入字符串
    std::string_view inputString = pArgInf[0].m_pText;
    std::string_view searchString = pArgInf[1].m_pText;

    // 空字符串检查
    if (inputString.empty() || searchString.empty()) {
        pRetData->m_int = -1;
        return;
    }

    // 确定开始搜索的位置
    size_t searchStartPos = (pArgInf[2].m_dtDataType == _SDT_NULL || pArgInf[2].m_int <= 1) ? 0 : pArgInf[2].m_int - 1;

    // 搜索指定字符串
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

    // 设置返回位置，如果找到则返回位置，否则返回-1
    pRetData->m_int = (searchResult != inputString.end()) ? std::distance(inputString.begin(), searchResult) + 1 : -1;
}
