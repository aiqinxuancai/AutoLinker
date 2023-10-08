#include "fnshare.h"

#define s_pfnNotifySys      __LIB2_DEFFUNNAME(pfnNotifySys_var)
#define s_pfnuserNotifySys  __LIB2_DEFFUNNAME(pfnuserNotifySys_var)
#define s_isDebug           __LIB2_DEFFUNNAME(isDebug)

static PFN_NOTIFY_SYS s_pfnNotifySys = 0;
static PFN_NOTIFY_SYS s_pfnuserNotifySys = 0;    // 由函数 SetUserSysNotify() 设置, 
static int s_isDebug = 1253600;

INT __stdcall NotifySys(INT nMsg, DWORD dwParam1, DWORD dwParam2)
{
    INT ret = 0;
    if (s_pfnNotifySys != NULL)
        ret = s_pfnNotifySys(nMsg, dwParam1, dwParam2);
    return ret;
}
bool isDebugVer(int* r)
{
    if (r)*r = s_isDebug;
    return s_isDebug == PT_DEBUG_RUN_VER;
}

INT __stdcall ProcessNotifyLib(INT nMsg, DWORD dwParam1, DWORD dwParam2)
{
    INT nRet = NR_OK;
    switch (nMsg)
    {
    case NL_SYS_NOTIFY_FUNCTION:
    {
        s_pfnNotifySys = (PFN_NOTIFY_SYS)dwParam1;
        if (s_isDebug == 1253600)
        {
            s_isDebug = NotifySys(NRS_GET_PRG_TYPE, 0, 0);
        }
        break;
    }
    case NL_FREE_LIB_DATA:

        break;
    case NL_GET_CMD_FUNC_NAMES:
        //应返回所有命令和方法实现函数的函数名称数组, 
        //此数组必须与命令和方法定义数组一一对应, 
        //数组各成员均为对应函数的函数名称文本指针(char*)；

        break;
    case NL_GET_NOTIFY_LIB_FUNC_NAME:
        //应返回“处理系统通知的函数”(即自身函数)的函数名称(char*)；

        break;
    case NL_GET_DEPENDENT_LIBS:
        //应返回“依赖的第三方静态库文件列表”, 格式为\0分隔的文本, 结尾两个\0。

        break;
    default:
        nRet = NR_ERR;
        break;
    }

    //调用用户代码
    if (s_pfnuserNotifySys)
        nRet = s_pfnuserNotifySys(nMsg, dwParam1, dwParam2);

    return nRet;
}

PFN_NOTIFY_SYS __stdcall SetUserSysNotify(PFN_NOTIFY_SYS pfn)
{
    s_pfnuserNotifySys = pfn;
    return ProcessNotifyLib;
}
