/*
    版权声明：
    本文件版权为易语言作者吴涛所有, 仅授权给第三方用作开发易语言支持库, 禁止用于其他任何场合。
*/

//本单元不使用MFC类, 可供编写跨平台的支持库或非窗口组件支持库使用

#ifndef __FN_SHARE_H
#define __FN_SHARE_H

#include "lib2.h"
typedef INT(cdecl* PFN_ON_SYS_NOTIFY) (INT nMsg, DWORD dwParam1, DWORD dwParam2);
#ifndef _private
#define _private  //称识为只私有
#endif




// 向易语言ide发送消息
#define NotifySys __LIB2_DEFFUNNAME(jy_NotifySys)
INT WINAPI NotifySys(INT nMsg, DWORD dwParam1, DWORD dwParam2);


// 从易语言里申请内存, 单位为字节
#define ealloc __LIB2_DEFFUNNAME(jy_ealloc)
inline void* ealloc(int size)
{
    void* pMem = (void*)NotifySys(NRS_MALLOC, size, 0);
    memset(pMem, 0, size);
    return (void*)pMem;
}

// 释放从易语言里申请的内存
#define efree __LIB2_DEFFUNNAME(jy_efree)
inline void efree(void* p)
{
    NotifySys(NRS_MFREE, (DWORD)p, 0);
}



// 是否为调试版, 这个值是支持库加载时取一次
#define isDebugVer __LIB2_DEFFUNNAME(jy_GetVer)
bool isDebugVer(int* r); // 获取当前版本是否为编译版


// 与易语言ide通讯子程序
#define ProcessNotifyLib __LIB2_DEFFUNNAME(jy_ProcessNotifyLib)
INT WINAPI ProcessNotifyLib(INT nMsg, DWORD dwParam1, DWORD dwParam2);


// 用户自己定义的通知函数, 用户在使用以下代码即可保证不修改本文件, 而达到及时与易系统通讯的目的：
#define SetUserSysNotify __LIB2_DEFFUNNAME(jy_NotifySys)
PFN_NOTIFY_SYS WINAPI SetUserSysNotify(PFN_NOTIFY_SYS pfn);


// 拷贝文本, 不使用时需要调用 efree() 来释放
// nTextLen用作指定文本部分的长度(不包含结束零)
// 如果为0, 则取ps的全部长度
#define CloneTextData __LIB2_DEFFUNNAME(jy_CloneTextData)
inline char* CloneTextData(const char* ps, INT nTextLen = 0)
{
    if (ps == NULL || *ps == '\0') return NULL;
    if (nTextLen == 0) nTextLen = lstrlenA(ps);
    if (nTextLen == 0) return NULL;
    char* pd = (char*)ealloc(nTextLen + 1);
    memcpy(pd, ps, nTextLen);
    pd[nTextLen] = '\0';
    return pd;
}

// 拷贝字节集, 生成易语言格式的字节集,1维数组
#define CloneBinData __LIB2_DEFFUNNAME(jy_CloneBinData)
inline LPBYTE CloneBinData(LPBYTE pData, INT nDataSize)
{
    if (nDataSize == 0)
        return NULL;

    LPBYTE pd = (LPBYTE)ealloc(sizeof(INT) * 2 + nDataSize);
    *(LPINT)pd = 1;
    *(LPINT)(pd + sizeof(INT)) = nDataSize;
    memcpy(pd + sizeof(INT) * 2, pData, nDataSize);
    return pd;
}

// 拷贝W版文本, 不使用时需要调用 efree() 来释放
// nTextLen用作指定文本部分的长度(不包含结束零)
// 如果为0, 则取ps的全部长度
#define CloneTextDataW __LIB2_DEFFUNNAME(jy_CloneTextDataW)
inline LPBYTE CloneTextDataW(LPCWSTR ps, INT nTextLen = 0)
{
    if ( ps == NULL || *ps == '\0' ) return NULL;
    if ( nTextLen == 0 ) nTextLen = (int)wcslen(ps);
    if ( nTextLen == 0 ) return NULL;

    int size = nTextLen * sizeof(wchar_t) + sizeof(wchar_t) + 8;
    LPINT pd = (LPINT)ealloc(size);
    pd[0] = 1;
    pd[1] = size - 8;
    memcpy(&pd[2], ps, nTextLen * sizeof(wchar_t));
    return (LPBYTE)pd;
}



// 返回数组的数据部分首地址及成员数目。 返回的地址不需要额外释放
#define GetAryElementInf __LIB2_DEFFUNNAME(jy_GetAryElementInf)
inline LPBYTE GetAryElementInf(void* pAryData, LPINT pnElementCount)
{
    if (pnElementCount)*pnElementCount = 0;
    if (!pAryData)return 0;
    LPINT pnData = (LPINT)pAryData;
    INT nArys = *pnData++;  // 取得维数。
    // 计算成员数目。
    INT nElementCount = 1;
    while (nArys > 0)
    {
        nElementCount *= *pnData++;
        nArys--;
    }

    if (pnElementCount)
        *pnElementCount = nElementCount;
    return (LPBYTE)pnData;
}


// 从字节集参数里获取字节集数据, 返回的字节集多2个字节保存\0\0,如果不是字符串,可以不使用这个, 不使用时需要调用 efree() 来释放
#define GetBinData __LIB2_DEFFUNNAME(jy_GetBinData)
inline LPBYTE GetBinData(LPBYTE pBin, int* len = 0)
{
    if (len)*len = 0;
    if (!pBin)return 0;
    int dwLen = 0;
    LPBYTE pData = GetAryElementInf(pBin, &dwLen);
    LPBYTE pMem = (LPBYTE)ealloc(dwLen + 2);
    memcpy(pMem, pData, dwLen);
    if (len)*len = dwLen + 2;
    return pMem;
}

// 取回数据类型的类别。
#define GetDataTypeType __LIB2_DEFFUNNAME(jy_GetDataTypeType)
inline INT GetDataTypeType(DATA_TYPE dtDataType)
{
#define DTT_IS_NULL_DATA_TYPE   0
#define DTT_IS_SYS_DATA_TYPE    1
#define DTT_IS_USER_DATA_TYPE   2
#define DTT_IS_LIB_DATA_TYPE    3
    if (dtDataType == _SDT_NULL)
        return DTT_IS_NULL_DATA_TYPE;

    DWORD dw = dtDataType & 0xC0000000;
    return dw == DTM_SYS_DATA_TYPE_MASK ? DTT_IS_SYS_DATA_TYPE :
        dw == DTM_USER_DATA_TYPE_MASK ? DTT_IS_USER_DATA_TYPE :
        DTT_IS_LIB_DATA_TYPE;
}

// 申请数组, 根据第二个参数来创建字节数, 默认为4字节
#define allocArray __LIB2_DEFFUNNAME(jy_allocArray)
inline LPBYTE allocArray(int nCount, int elementSize = sizeof(int))
{
    if (elementSize <= 0 || elementSize > 8)elementSize = 4;
    int len = (2 * sizeof(int)) + elementSize * nCount;
    int* p = (int*)ealloc(len);
    p[0] = 1;  // 数组维数。   
    p[1] = nCount;
    return (LPBYTE)p;
}

//根据易编译版本设置默认的等宽字体
#define wkeSetLogfont_CompileLanVer __LIB2_DEFFUNNAME(jy_wkeSetLogfont_CompileLanVer)
inline void wkeSetLogfont_CompileLanVer(LPLOGFONTA pLogfont) //根据易编译版本设置默认的等宽字体
{
    ::GetObjectA(::GetStockObject(DEFAULT_GUI_FONT), sizeof(LOGFONTA), pLogfont);
#ifndef __COMPILE_LANG_VER
    //  #error *** lang.h must be imported!
#endif
//#pragma setlocale(_T("en"))
#if (__COMPILE_LANG_VER == __BIG5_LANG_VER)
    /* 在繁体版Windows中DEFAULT_GUI_FONT中的字体名是“新細明體”, 
       而该字体是变宽的, 所以必须强制改为“細明體”。
       下面为細明體BIG5内码。为了避免被转换工具替换, 必须以这种格式写到程序中。 */
#if _MSC_VER < 1201
    strcpy(pLogfont->lfFaceName, "\xB2\xD3\xA9\xFA\xC5\xE9");
#else
    strcpy_s(pLogfont->lfFaceName, sizeof(pLogfont->lfFaceName) / sizeof(pLogfont->lfFaceName[0]), "\xB2\xD3\xA9\xFA\xC5\xE9");
#endif
#elif (__COMPILE_LANG_VER == __SJIS_JP_LANG_VER)
    /* 在日文版Windows中DEFAULT_GUI_FONT中的字体名是“MS UI Gothic”, 
       而该字体是变宽的, 所以必须强制改为“MS Gothic”。*/
    pLogfont->lfCharSet = SHIFTJIS_CHARSET;
    strcat(pLogfont->lfFaceName, "MS Gothic");
#elif (__COMPILE_LANG_VER == __ENGLISH_LANG_VER)
    strcat(pLogfont->lfFaceName, "Courier New");
#endif
}








//以下由liigo添加

/*
//!!! 以下函数均假设数据类型之成员全部为INT类型(或同样占用4个字节)
//index: 0..*/

inline _private void* _GetPointerByIndex(void* pBase, int index)
{
    return (void*)((INT)pBase + index * sizeof(INT));
}

inline _private INT _GetIntByIndex(void* pBase, INT index)
{
    return *((INT*)_GetPointerByIndex(pBase, index));
}

inline _private void _SetIntByIndex(void* pBase, INT index, INT value)
{
    *((INT*)_GetPointerByIndex(pBase, index)) = value;
}

inline INT _GetIntByIndex(PMDATA_INF pArgInf, INT index)
{
    return _GetIntByIndex(pArgInf->m_pCompoundData, index);
}

inline void SetIntByIndex(PMDATA_INF pArgInf, INT index, INT value)
{
    _SetIntByIndex(pArgInf->m_pCompoundData, index, value);
}

inline void* GetPointerByIndex(PMDATA_INF pArgInf, INT index)
{
    return _GetPointerByIndex(pArgInf->m_pCompoundData, index);
}






//命令信息占位符
#define CMD_INFO_MASK_1 \
        {\
        /*ccname*/    _T("???"),\
        /*egname*/    _T("???"),\
        /*explain*/    NULL,\
        /*category*/-1,\
        /*state*/    CT_IS_HIDED,\
        /*ret*/        _SDT_NULL,\
        /*reserved*/0,\
        /*level*/    LVL_SIMPLE,\
        /*bmp inx*/    0,\
        /*bmp num*/    0,\
        /*ArgCount*/0,\
        /*arg lp*/    NULL,\
        }
#define CMD_INFO_MASK_2        CMD_INFO_MASK_1,CMD_INFO_MASK_1
#define CMD_INFO_MASK_3        CMD_INFO_MASK_2,CMD_INFO_MASK_1
#define CMD_INFO_MASK_4        CMD_INFO_MASK_3,CMD_INFO_MASK_1
#define CMD_INFO_MASK_5        CMD_INFO_MASK_4,CMD_INFO_MASK_1
#define CMD_INFO_MASK_6        CMD_INFO_MASK_5,CMD_INFO_MASK_1
#define CMD_INFO_MASK_7        CMD_INFO_MASK_6,CMD_INFO_MASK_1
#define CMD_INFO_MASK_8        CMD_INFO_MASK_7,CMD_INFO_MASK_1
#define CMD_INFO_MASK_9        CMD_INFO_MASK_8,CMD_INFO_MASK_1
#define CMD_INFO_MASK_10    CMD_INFO_MASK_9,CMD_INFO_MASK_1

#define MASK_CMD_INFO_1        CMD_INFO_MASK_1
#define MASK_CMD_INFO_2        CMD_INFO_MASK_2
#define MASK_CMD_INFO_3        CMD_INFO_MASK_3
#define MASK_CMD_INFO_4        CMD_INFO_MASK_4
#define MASK_CMD_INFO_5        CMD_INFO_MASK_5
#define MASK_CMD_INFO_6        CMD_INFO_MASK_6
#define MASK_CMD_INFO_7        CMD_INFO_MASK_7
#define MASK_CMD_INFO_8        CMD_INFO_MASK_8
#define MASK_CMD_INFO_9        CMD_INFO_MASK_9
#define MASK_CMD_INFO_10    CMD_INFO_MASK_10

//数据类型占位符
#define DATATYPE_INFO_MASK_1 \
    {\
    /*m_szName*/            _T(""),\
    /*m_szEgName*/            _T(""),\
    /*m_szExplain*/            NULL,\
    /*m_nCmdCount*/            0,\
    /*m_pnCmdsIndex*/        NULL,\
    /*m_dwState*/            NULL,\
    /*m_dwUnitBmpID*/        0,\
    /*m_nEventCount*/        0,\
    /*m_pEventBegin*/        NULL,\
    /*m_nPropertyCount*/    0,\
    /*m_pPropertyBegin*/    NULL,\
    /*m_pfnGetInterface*/    NULL,\
    /*m_nElementCount*/        0,\
    /*m_pElementBegin*/        NULL,\
    }
#define DATATYPE_INFO_MASK_2    DATATYPE_INFO_MASK_1,DATATYPE_INFO_MASK_1
#define DATATYPE_INFO_MASK_3    DATATYPE_INFO_MASK_2,DATATYPE_INFO_MASK_1
#define DATATYPE_INFO_MASK_4    DATATYPE_INFO_MASK_3,DATATYPE_INFO_MASK_1
#define DATATYPE_INFO_MASK_5    DATATYPE_INFO_MASK_4,DATATYPE_INFO_MASK_1
#define DATATYPE_INFO_MASK_6    DATATYPE_INFO_MASK_5,DATATYPE_INFO_MASK_1
#define DATATYPE_INFO_MASK_7    DATATYPE_INFO_MASK_6,DATATYPE_INFO_MASK_1
#define DATATYPE_INFO_MASK_8    DATATYPE_INFO_MASK_7,DATATYPE_INFO_MASK_1
#define DATATYPE_INFO_MASK_9    DATATYPE_INFO_MASK_8,DATATYPE_INFO_MASK_1
#define DATATYPE_INFO_MASK_10    DATATYPE_INFO_MASK_9,DATATYPE_INFO_MASK_1

#define MASK_DATATYPE_INFO_1    DATATYPE_INFO_MASK_1
#define MASK_DATATYPE_INFO_2    DATATYPE_INFO_MASK_2
#define MASK_DATATYPE_INFO_3    DATATYPE_INFO_MASK_3
#define MASK_DATATYPE_INFO_4    DATATYPE_INFO_MASK_4
#define MASK_DATATYPE_INFO_5    DATATYPE_INFO_MASK_5
#define MASK_DATATYPE_INFO_6    DATATYPE_INFO_MASK_6
#define MASK_DATATYPE_INFO_7    DATATYPE_INFO_MASK_7
#define MASK_DATATYPE_INFO_8    DATATYPE_INFO_MASK_8
#define MASK_DATATYPE_INFO_9    DATATYPE_INFO_MASK_9
#define MASK_DATATYPE_INFO_10    DATATYPE_INFO_MASK_10

//数据类型成员占位符
#define DATATYPE_ELEMENT_MASK_1 \
    {\
/*m_dtType*/    SDT_INT,\
/*m_pArySpec*/    NULL,\
/*m_szName*/    _WT(""),\
/*m_szEgName*/    _WT(""),\
/*m_szExplain*/    NULL,\
/*m_dwState*/    LES_HIDED,\
/*m_nDefault*/    0,\
    }

#define DATATYPE_ELEMENT_MASK_2        DATATYPE_ELEMENT_MASK_1,DATATYPE_ELEMENT_MASK_1
#define DATATYPE_ELEMENT_MASK_3        DATATYPE_ELEMENT_MASK_2,DATATYPE_ELEMENT_MASK_1
#define DATATYPE_ELEMENT_MASK_4        DATATYPE_ELEMENT_MASK_3,DATATYPE_ELEMENT_MASK_1
#define DATATYPE_ELEMENT_MASK_5        DATATYPE_ELEMENT_MASK_4,DATATYPE_ELEMENT_MASK_1
#define DATATYPE_ELEMENT_MASK_6        DATATYPE_ELEMENT_MASK_5,DATATYPE_ELEMENT_MASK_1
#define DATATYPE_ELEMENT_MASK_7        DATATYPE_ELEMENT_MASK_6,DATATYPE_ELEMENT_MASK_1
#define DATATYPE_ELEMENT_MASK_8        DATATYPE_ELEMENT_MASK_7,DATATYPE_ELEMENT_MASK_1
#define DATATYPE_ELEMENT_MASK_9        DATATYPE_ELEMENT_MASK_8,DATATYPE_ELEMENT_MASK_1
#define DATATYPE_ELEMENT_MASK_10    DATATYPE_ELEMENT_MASK_9,DATATYPE_ELEMENT_MASK_1

#define MASK_DATATYPE_ELEMENT_1        DATATYPE_ELEMENT_MASK_1
#define MASK_DATATYPE_ELEMENT_2        DATATYPE_ELEMENT_MASK_2
#define MASK_DATATYPE_ELEMENT_3        DATATYPE_ELEMENT_MASK_3
#define MASK_DATATYPE_ELEMENT_4        DATATYPE_ELEMENT_MASK_4
#define MASK_DATATYPE_ELEMENT_5        DATATYPE_ELEMENT_MASK_5
#define MASK_DATATYPE_ELEMENT_6        DATATYPE_ELEMENT_MASK_6
#define MASK_DATATYPE_ELEMENT_7        DATATYPE_ELEMENT_MASK_7
#define MASK_DATATYPE_ELEMENT_8        DATATYPE_ELEMENT_MASK_8
#define MASK_DATATYPE_ELEMENT_9        DATATYPE_ELEMENT_MASK_9
#define MASK_DATATYPE_ELEMENT_10    DATATYPE_ELEMENT_MASK_10

//空白函数指针占位符
#define NULL_1     NULL
#define NULL_2     NULL,NULL
#define NULL_3     NULL_2,NULL_1
#define NULL_4     NULL_3,NULL_1
#define NULL_5     NULL_4,NULL_1
#define NULL_6     NULL_5,NULL_1
#define NULL_7     NULL_6,NULL_1
#define NULL_8     NULL_7,NULL_1
#define NULL_9     NULL_8,NULL_1
#define NULL_10     NULL_9,NULL_1

#define    MASK_NULL_1        NULL_1
#define MASK_NULL_2        NULL_2
#define MASK_NULL_3        NULL_3
#define MASK_NULL_4        NULL_4
#define MASK_NULL_5        NULL_5
#define MASK_NULL_6        NULL_6
#define MASK_NULL_7        NULL_7
#define MASK_NULL_8        NULL_8
#define MASK_NULL_9        NULL_9
#define MASK_NULL_10    NULL_10

#endif //__FN_SHARE_H
