#ifndef __MTYPE_H
#define __MTYPE_H

#include <time.h>

#define	FAR
#define	NEAR
#define WINAPI

typedef long LONG;
typedef int BOOL;
typedef float FLOAT;
typedef double DOUBLE;
typedef int INT;
typedef short SHORT;
typedef unsigned long ULONG;
typedef ULONG* PULONG;
typedef unsigned short USHORT;
typedef USHORT* PUSHORT;
typedef unsigned char UCHAR;
typedef char TCHAR;
typedef UCHAR* PUCHAR;
typedef char* PSZ;
typedef char* LPTSTR;
typedef const char* LPCTSTR;
typedef unsigned long DWORD;
typedef unsigned char BYTE;
typedef unsigned short WORD;
typedef FLOAT* PFLOAT;
typedef BOOL NEAR* PBOOL;
typedef BOOL FAR* LPBOOL;
typedef BYTE NEAR* PBYTE;
typedef BYTE FAR* LPBYTE;
typedef int NEAR* PINT;
typedef int FAR* LPINT;
typedef WORD NEAR* PWORD;
typedef WORD FAR* LPWORD;
typedef long FAR* LPLONG;
typedef DWORD NEAR* PDWORD;
typedef DWORD FAR* LPDWORD;
typedef void FAR* LPVOID;
typedef const void FAR* LPCVOID;
typedef unsigned int UINT;
typedef unsigned int* PUINT;
typedef double DATE;
typedef long long INT64;
typedef INT64 NEAR* PINT64;
typedef INT64 FAR* LPINT64;

/////////////////////////////////

typedef DWORD HGLOBAL;
typedef DWORD HWND;
typedef DWORD HMENU;
typedef DWORD COLORREF;
typedef DWORD HICON;
typedef DWORD HANDLE;
typedef void* HINSTANCE;
typedef void (*FARPROC) ();

typedef struct tagRECT
{
    LONG    left;
    LONG    top;
    LONG    right;
    LONG    bottom;
} RECT, * PRECT, NEAR* NPRECT, FAR* LPRECT;

typedef const RECT FAR* LPCRECT;

typedef struct _RECTL       /* rcl */
{
    LONG    left;
    LONG    top;
    LONG    right;
    LONG    bottom;
} RECTL, * PRECTL, * LPRECTL;

typedef const RECTL FAR* LPCRECTL;

typedef struct tagPOINT
{
    LONG  x;
    LONG  y;
} POINT, * PPOINT, NEAR* NPPOINT, FAR* LPPOINT;

typedef struct _POINTL      /* ptl  */
{
    LONG  x;
    LONG  y;
} POINTL, * PPOINTL;

typedef struct tagSIZE
{
    LONG        cx;
    LONG        cy;
} SIZE, * PSIZE, * LPSIZE;

typedef SIZE               SIZEL;
typedef SIZE* PSIZEL, * LPSIZEL;

typedef struct tagPOINTS
{
    SHORT   x;
    SHORT   y;
} POINTS, * PPOINTS, * LPPOINTS;

/////////////////////////////////

#ifndef NULL
#ifdef __cplusplus
#define NULL    0
#else
#define NULL    ((void *)0)
#endif
#endif

#ifndef FALSE
#define FALSE               0
#endif

#ifndef TRUE
#define TRUE                1
#endif

#ifndef max
#define max(a,b)            (((a) > (b)) ? (a) : (b))
#endif

#ifndef min
#define min(a,b)            (((a) < (b)) ? (a) : (b))
#endif

#ifndef INT_MAX
#define INT_MAX             0x7FFFFFFF
#endif

#define MAKEWORD(a, b)      ((WORD)(((BYTE)(a)) | ((WORD)((BYTE)(b))) << 8))
#define MAKELONG(a, b)      ((LONG)(((WORD)(a)) | ((DWORD)((WORD)(b))) << 16))
#define LOWORD(l)           ((WORD)(l))
#define HIWORD(l)           ((WORD)(((DWORD)(l) >> 16) & 0xFFFF))
#define LOBYTE(w)           ((BYTE)(w))
#define HIBYTE(w)           ((BYTE)(((WORD)(w) >> 8) & 0xFF))

#define _T(string)            string
#define IS_CC(ch)             ((ch & 0x80) != 0)
#define GET_HZ(btLow,btHigh)  (MAKEWORD (btHigh, btLow))

#define MIN_DATE            (-657434L)  // about year 100
#define MAX_DATE            2958465L    // about year 9999

#define ENTRY_LIB_FUNCTION

#ifdef __COMPILE_FNR
#define _WT(text)	                    _T("")
#define _CMDS_COUNT(CmdsCount)          0
#define _CMDS_PTR(CmdsPtr)              NULL
#define _EVENTS_COUNT(EventsCount)      0
#define _EVENTS_PTR(EventsPtr)          NULL
#define _ELEMENTS_COUNT(ElementsCount)  0
#define _ELEMENTS_PTR(ElementsPtr)      NULL
#define _FORM_BASE_CLASS	            CWnd
#else
#define _WT(text)	                    _T(text)
#define _CMDS_COUNT(CmdsCount)          CmdsCount
#define _CMDS_PTR(CmdsPtr)              CmdsPtr
#define _EVENTS_COUNT(EventsCount)      EventsCount
#define _EVENTS_PTR(EventsPtr)          ((PEVENT_INFO2)EventsPtr)
#define _ELEMENTS_COUNT(ElementsCount)  ElementsCount
#define _ELEMENTS_PTR(ElementsPtr)      ElementsPtr
#define _FORM_BASE_CLASS	            CFrameWnd
#endif

/////////////////////////////////

#endif
