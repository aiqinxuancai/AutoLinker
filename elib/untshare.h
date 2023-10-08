/*
    版权声明：
    本文件版权为易语言作者吴涛所有，仅授权给第三方用作开发易语言支持库，禁止用于其他任何场合。
*/

// 本文件修改为不使用MFC, 可以愉快的使用c++了

#ifndef __UNTSHARE_H
#define    __UNTSHARE_H
#include "fnshare.h"



class CPropertyInfo
{
public:
    CPropertyInfo()
    {
        init();
    }

    virtual void init() { ; }
    virtual BOOL Serialize(CPropertyInfo& ar)
    {
        //if (ar.IsLoading() == TRUE)
        //    init();
        return TRUE;
    }

    virtual HGLOBAL SaveData(LPINT pnAllDataSize = NULL)
    {
        HGLOBAL hGlobal = 0;
        //int dwLength = 0;
        //CunshareMem ar;
        //ar.SetMode(TRUE);    //设置为读数据

        //if (Serialize(ar) == TRUE)
        //{
        //    // 走到这就是子类读好数据了

        //    dwLength = ar.size();
        //    hGlobal = ::GlobalAlloc(GMEM_MOVEABLE, dwLength);
        //    if (!hGlobal) throw _SDT_ALL;
        //    void* pMem = ::GlobalLock(hGlobal);
        //    if (!pMem) throw _SDT_ALL;
        //    memcpy(pMem, ar.GetPtr(), dwLength);
        //    ::GlobalUnlock(hGlobal);
        //    // 注意由于Windows系统内存管理按字对齐的原因，最后hGlobal的大小（GloalSize返回）可能会大于dwLength值0-3个字节。
        //}
        //if (pnAllDataSize) *pnAllDataSize = dwLength;

        return hGlobal;
    }
    virtual BOOL LoadData(LPBYTE pData, INT nDataSize)
    {
        if (pData != NULL && nDataSize > 0)
        {
            //CunshareMem ar;
            //ar.add(pData, nDataSize);
            //ar.SetMode(FALSE);
            //return Serialize(ar);
        }

        return TRUE;
    }



};



// 取系统背景色
#define ProcessSysBackColor    __LIB2_DEFFUNNAME(jy_ProcessSysBackColor)
inline COLORREF ProcessSysBackColor(COLORREF clr)
{
#define CLR_DEFAULT             0xFF000000L
    if (clr == CLR_DEFAULT)
        return ::GetSysColor(COLOR_BTNFACE);
    else
        return clr;
}


// 修改窗口样式, dwAddStyle=加入的样式, dwRemoveStyle=删除的样式, ExStyel=是否为扩展样式
#define ModiUnitStyle        __LIB2_DEFFUNNAME(jy_ModiUnitStyle)
inline void    ModiUnitStyle(HWND hWnd, DWORD dwAddStyle, DWORD dwRemoveStyle, BOOL ExStyel = FALSE)
{
    int index = GWL_STYLE;
    if (ExStyel) index = GWL_EXSTYLE;
    DWORD dwOldStyle = GetWindowLongW(hWnd, index);
    DWORD dwNewStyle = (dwOldStyle & ~dwRemoveStyle) | dwAddStyle;
    if (dwNewStyle != dwOldStyle)
        SetWindowLongW(hWnd, index, dwNewStyle);
}

// 设置样式, blSet=是否添加,TRUE=添加,FALSE=删除 , ExStyel = 是否为扩展样式
#define SetUnitStyle        __LIB2_DEFFUNNAME(jy_SetUnitStyle)
inline void    SetUnitStyle(HWND hWnd, BOOL blSet, DWORD dwStyle, BOOL ExStyel = FALSE)
{
    ModiUnitStyle(hWnd, (blSet ? dwStyle : 0), (blSet ? 0 : dwStyle), ExStyel);
}


// 修改边框 0=无边框,1=凹入式,2=凸出式,3=浅凹入式,4=镜框式,5=单线边框
#define ChangeBorder        __LIB2_DEFFUNNAME(jy_ChangeBorder)
inline void    ChangeBorder(HWND hWnd, INT nBorderType)
{
    DWORD dwStyle = NULL, dwExStyle = NULL;
    switch (nBorderType)
    {
    //case 0:        // 无边框
    //    dwExStyle = WS_EX_CLIENTEDGE;
    //    break;
    case 1:        // 凹入式
        dwExStyle = WS_EX_CLIENTEDGE;
        break;
    case 2:        // 凸出式
        dwExStyle = WS_EX_DLGMODALFRAME;
        break;
    case 3:        // 浅凹入式
        dwExStyle = WS_EX_STATICEDGE;
        break;
    case 4:        // 镜框式
        dwExStyle = WS_EX_CLIENTEDGE | WS_EX_DLGMODALFRAME;
        break;
    case 5:
        dwStyle = WS_BORDER;
        break;
    }
    ModiUnitStyle(hWnd, dwExStyle, WS_EX_STATICEDGE | WS_EX_CLIENTEDGE | WS_EX_DLGMODALFRAME, TRUE);
    SetWindowPos(hWnd, 0, 0, 0, 0, 0, SWP_NOSIZE | SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);

    ModiUnitStyle(hWnd, dwStyle, WS_BORDER, FALSE);
    SetWindowPos(hWnd, 0, 0, 0, 0, 0, SWP_NOSIZE | SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);


    //CWnd* pUnit;
    //pUnit->ModifyStyleEx(WS_EX_STATICEDGE | WS_EX_CLIENTEDGE | WS_EX_DLGMODALFRAME,
    //    dwExStyle, SWP_NOSIZE | SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE |
    //    SWP_FRAMECHANGED | SWP_DRAWFRAME);
    //pUnit->ModifyStyle(WS_BORDER, dwStyle,
    //    SWP_NOSIZE | SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE |
    //    SWP_FRAMECHANGED | SWP_DRAWFRAME);
}



// 绝对不会返回NULL或者窗口句柄无效的CWnd*指针。
#define GetWndPtr            __LIB2_DEFFUNNAME(jy_GetWndPtr)
inline void*   __LIB2_DEFFUNNAME(jy_GetWndPtr)(PMDATA_INF pInf)
{
    return (void*)NotifySys(NRS_GET_AND_CHECK_UNIT_PTR, pInf->m_unit.m_dwFormID, pInf->m_unit.m_dwUnitID);
}



// 创建字体
#define MyCreateFont        __LIB2_DEFFUNNAME(jy_MyCreateFont)
inline HFONT   MyCreateFont(LPLOGFONTW* pLogFont)
{

    LOGFONTW lf = { 0 };
    if (pLogFont)
        memcpy(&lf, pLogFont, sizeof(LOGFONTW));
    else
        SystemParametersInfoW(SPI_GETICONTITLELOGFONT, sizeof(LOGFONTW), &lf, 0);  // 获取系统默认字体

    lf.lfCharSet        = DEFAULT_CHARSET;
    lf.lfOutPrecision   = OUT_DEFAULT_PRECIS;
    lf.lfClipPrecision  = CLIP_DEFAULT_PRECIS;
    lf.lfQuality        = PROOF_QUALITY;
    lf.lfPitchAndFamily = DEFAULT_PITCH;

    return CreateFontIndirectW(&lf);
}



// 序列化字符串, 读取或者写入
#define SerializeCString    __LIB2_DEFFUNNAME(jy_SerializeCString)
inline void    SerializeCString(LPCSTR str, CPropertyInfo& ar)
{
    //INT nSize = 0;
    //if (ar.IsLoading() == FALSE)
    //{
    //    LPSTR s;
    //    ar >> s;
    //    str = s;
    //}
    //else
    //{
    //    ar << str.c_str() << (BYTE)0;
    //}
}


// 加载图标, 返回图标数据的图标句柄
#define LoadIco                __LIB2_DEFFUNNAME(jy_LoadIco)
inline HICON   LoadIco(LPBYTE pData, INT nDataSize, BOOL blIsBigIcon)
{
    if (pData == NULL || nDataSize <= 0)  return NULL;

    //CunshareMem file;
    //file.add(pData, nDataSize);

    //try
    //{
    //    WORD* wBuf = 0;
    //    file.read((void**)&wBuf, sizeof(WORD) * 3);
    //    if (wBuf[0] != 0 || (wBuf[1] != 1 && wBuf[1] != 2) || wBuf[2] <= 0)
    //        return NULL;

    //    BOOL blIsIcon = wBuf[1] == 1;
    //    INT nNumImages = (INT)wBuf[2];

    //    if (nNumImages <= 0)  return NULL;

    //    typedef struct
    //    {
    //        BYTE    bWidth;               // Width of the image
    //        BYTE    bHeight;              // Height of the image (times 2)
    //        BYTE    bColorCount;          // Number of colors in image (0 if >=8bpp)
    //        BYTE    bReserved;            // Reserved
    //        WORD    wPlanes;              // Color Planes
    //        WORD    wBitCount;            // Bits per pixel
    //        DWORD   dwBytesInRes;         // how many bytes in this resource?
    //        DWORD   dwImageOffset;        // where in the file is this image
    //    } ICONDIRENTRY, * LPICONDIRENTRY;

    //    LPICONDIRENTRY info = 0;
    //    INT nImageIndex = 0;
    //    for (nImageIndex = 0; nImageIndex < nNumImages; nImageIndex++)
    //    {
    //        file.SetPos(nImageIndex * sizeof(ICONDIRENTRY) + sizeof(WORD) * 3);
    //        file.read((void**)&info, sizeof(ICONDIRENTRY));
    //        if (!info)return NULL;

    //        if (blIsBigIcon == TRUE ? info->bWidth == 32 : info->bWidth == 16)
    //            break;
    //    }
    //    if (nImageIndex == nNumImages)
    //        return NULL;

    //    if (nImageIndex >= nNumImages)
    //        nImageIndex = nNumImages - 1;

    //    LPBYTE pBuf = NULL;

    //    file.SetPos(info->dwImageOffset);
    //    file.read((void**)&pBuf, info->dwBytesInRes);
    //    if (!pBuf)return NULL;

    //    LPBITMAPINFO pBitmapInfo = (LPBITMAPINFO)pBuf;

    //    HICON hIcon = CreateIconFromResourceEx(pBuf, info->dwBytesInRes,
    //        TRUE, 0x00030000, pBitmapInfo->bmiHeader.biWidth,
    //        pBitmapInfo->bmiHeader.biHeight / 2, 0);

    //    if (hIcon == NULL && pBitmapInfo->bmiHeader.biBitCount != 16)
    //        hIcon = CreateIconFromResource(pBuf, info->dwBytesInRes, TRUE, 0x00030000);

    //    return hIcon;
    //}
    //catch (...)
    //{

    //    return NULL;
    //}

    return NULL;
}


//属性信息占位符
#define UNIT_PROPERTY_MASK_1 \
    {\
        /*m_szName*/        _T(""),\
        /*m_szEgName*/      _T(""),\
        /*m_szExplain*/     NULL,\
        /*m_shtType*/       0,\
        /*m_wState*/        UW_IS_HIDED,\
        /*m_szzPickStr*/    NULL,\
    }
#define UNIT_PROPERTY_MASK_2    UNIT_PROPERTY_MASK_1,UNIT_PROPERTY_MASK_1
#define UNIT_PROPERTY_MASK_3    UNIT_PROPERTY_MASK_2,UNIT_PROPERTY_MASK_1
#define UNIT_PROPERTY_MASK_4    UNIT_PROPERTY_MASK_3,UNIT_PROPERTY_MASK_1
#define UNIT_PROPERTY_MASK_5    UNIT_PROPERTY_MASK_4,UNIT_PROPERTY_MASK_1
#define UNIT_PROPERTY_MASK_6    UNIT_PROPERTY_MASK_5,UNIT_PROPERTY_MASK_1
#define UNIT_PROPERTY_MASK_7    UNIT_PROPERTY_MASK_6,UNIT_PROPERTY_MASK_1
#define UNIT_PROPERTY_MASK_8    UNIT_PROPERTY_MASK_7,UNIT_PROPERTY_MASK_1
#define UNIT_PROPERTY_MASK_9    UNIT_PROPERTY_MASK_8,UNIT_PROPERTY_MASK_1
#define UNIT_PROPERTY_MASK_10   UNIT_PROPERTY_MASK_9,UNIT_PROPERTY_MASK_1

#define MASK_UNIT_PROPERTY_1    UNIT_PROPERTY_MASK_1
#define MASK_UNIT_PROPERTY_2    UNIT_PROPERTY_MASK_2
#define MASK_UNIT_PROPERTY_3    UNIT_PROPERTY_MASK_3
#define MASK_UNIT_PROPERTY_4    UNIT_PROPERTY_MASK_4
#define MASK_UNIT_PROPERTY_5    UNIT_PROPERTY_MASK_5
#define MASK_UNIT_PROPERTY_6    UNIT_PROPERTY_MASK_6
#define MASK_UNIT_PROPERTY_7    UNIT_PROPERTY_MASK_7
#define MASK_UNIT_PROPERTY_8    UNIT_PROPERTY_MASK_8
#define MASK_UNIT_PROPERTY_9    UNIT_PROPERTY_MASK_9
#define MASK_UNIT_PROPERTY_10   UNIT_PROPERTY_MASK_10

//事件信息占位符
#define EVENT_INFO2_MASK_1 \
    {\
        /*m_szName*/            _T(""),\
        /*m_szExplain*/         NULL,\
        /*m_dwState*/           EV_IS_HIDED | EV_IS_VER2,\
        /*m_nArgCount*/         0,\
        /*m_pBeginArgInfo*/     NULL,\
        /*m_dtRetDataType*/     _SDT_NULL,\
    }
#define EVENT_INFO2_MASK_2      EVENT_INFO2_MASK_1,EVENT_INFO2_MASK_1
#define EVENT_INFO2_MASK_3      EVENT_INFO2_MASK_2,EVENT_INFO2_MASK_1
#define EVENT_INFO2_MASK_4      EVENT_INFO2_MASK_3,EVENT_INFO2_MASK_1
#define EVENT_INFO2_MASK_5      EVENT_INFO2_MASK_4,EVENT_INFO2_MASK_1
#define EVENT_INFO2_MASK_6      EVENT_INFO2_MASK_5,EVENT_INFO2_MASK_1
#define EVENT_INFO2_MASK_7      EVENT_INFO2_MASK_6,EVENT_INFO2_MASK_1
#define EVENT_INFO2_MASK_8      EVENT_INFO2_MASK_7,EVENT_INFO2_MASK_1
#define EVENT_INFO2_MASK_9      EVENT_INFO2_MASK_8,EVENT_INFO2_MASK_1
#define EVENT_INFO2_MASK_10     EVENT_INFO2_MASK_9,EVENT_INFO2_MASK_1

#define MASK_EVENT_INFO2_1      EVENT_INFO2_MASK_1
#define MASK_EVENT_INFO2_2      EVENT_INFO2_MASK_2
#define MASK_EVENT_INFO2_3      EVENT_INFO2_MASK_3
#define MASK_EVENT_INFO2_4      EVENT_INFO2_MASK_4
#define MASK_EVENT_INFO2_5      EVENT_INFO2_MASK_5
#define MASK_EVENT_INFO2_6      EVENT_INFO2_MASK_6
#define MASK_EVENT_INFO2_7      EVENT_INFO2_MASK_7
#define MASK_EVENT_INFO2_8      EVENT_INFO2_MASK_8
#define MASK_EVENT_INFO2_9      EVENT_INFO2_MASK_9
#define MASK_EVENT_INFO2_10     EVENT_INFO2_MASK_10

#endif //__UNTSHARE_H
