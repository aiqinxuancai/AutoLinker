#include "pch.h"
#include "framework.h"
#include <math.h>


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