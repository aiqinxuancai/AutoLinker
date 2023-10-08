#ifndef __LIB_INF_H
#define __LIB_INF_H

#ifdef __cplusplus
#define EXTERN_C extern "C"
#else
#define EXTERN_C
#endif


//!////////////////////////////////////////////////////////////////////////
//! 以下是处理静态库需要的宏, 静态库尽可能的把函数名改成唯一的,不会跟其他库冲突的
//! 这里定义了几个宏 方法名加上 __LIB2_DEFFUNNAME(fun_name)  这是把 fun_name 替换成其他名字, 调用起来也需要 __LIB2_DEFFUNNAME(fun_name)(...) 进行调用
//! 实现支持库命令方法名可以使用 DEF_EXECUTE_CMD(fun_name)
// TODO 引用lib2.h前必须定义 __E_FNENAME 这个宏, 宏里面是库名前缀, 所有方法会在前面加上 库名_方法名_库名

#define __E_FNENAME AutoLinker        // 定义支持库命令用, 静态库需要改名

#define ____E_FNENAME(quote) #quote                    // 辅助用
#define ______E_FNENAME(quote) ____E_FNENAME(quote)    // 辅助用
#define __E_FNENAME_S ______E_FNENAME(__E_FNENAME)    // 静态库链接时用, 需要拼接成方法名字符串

#ifndef __E_FNENAME
#error "必须先定义 __E_FNENAME 这个宏, 并设置名字, 比如 #define __E_FNENAME jywke"
#endif


#define __LIB2_FNE_NAME_LEFT(leftName) leftName
#define __LIB2_FNE_NAME_RIGHT(rightName) rightName

#define __LIB2_DEFFUNNAME__(fnName) __LIB2_FNE_NAME_LEFT(__E_FNENAME)##_##fnName##_
#define __LIB2_DEFFUNNAME(fnName)   __LIB2_DEFFUNNAME__(fnName)__LIB2_FNE_NAME_LEFT(__E_FNENAME)

// 命令声明使用宏,可以最大限度的解决静态编译命令名冲突问题
#define DEF_EXECUTE_CMD(fnName) EXTERN_C void __LIB2_DEFFUNNAME(fnName) (PMDATA_INF pRetData, INT nArgCount, PMDATA_INF pArgInf)

#if defined(__E_STATIC_LIB) || defined(__COMPILE_FNR)
#define DEF_CMD(name)  __LIB2_DEFFUNNAME(name)  // 对于不同的库, 请做相应的名称修改
#define DEF_NULL_CMD(name)  NULL
#else if defined(__cplusplus)
#define DEF_CMD(name)  _RecordCmdName (__LIB2_DEFFUNNAME(name), __E_FNENAME_S "_"#name"_" __E_FNENAME_S)    // 对于不同的库, 请做相应的名称修改
#define DEF_CMD_NULL(name)  _RecordCmdName(NULL, NULL)                // 定义空命令
#endif

//!////////////////////////////////////////////////////////////////////////



//////////////////////////////////////////////////////////////////////////

// 操作系统类别：

#define __OS_WIN        0x80000000
#define __OS_LINUX      0x40000000
#define __OS_UNIX       0x20000000

#define OS_ALL    (__OS_WIN | __OS_LINUX | __OS_UNIX)

//!////////////////////////////////////////////////////////////////////////

#include <windows.h>
#include <stdio.h>
#include <math.h>

//#include <windef.h>
//#include <tchar.h>
//#include <winnt.h>
//
/*

关于跨操作系统编程支持技术说明：

  1、各支持库需要说明其具有哪些操作系统版本,及支持库中所有命令、数据类型、
     数据类型方法、数据类型事件、数据类型属性所支持的操作系统。以上这些信息,
     在该支持库所具有的所有操作系统版本中必须一致。

  2、为了和以前的支持库相兼容,所有m_nRqSysMajorVer.m_nRqSysMinorVer版本号
     小于3.6的都默认为支持Windows操作系统,包括内部的所有命令、数据类型、
     数据类型方法、数据类型事件、数据类型属性。

  3、所有组件固定属性和固定事件都支持所有操作系统。

  4、对于纯工具库不进行操作系统支持检查。

*/

#ifndef DOUBLE
typedef double DOUBLE;
#endif
#ifndef DATE
typedef double DATE;
#endif
#ifndef SHORT
typedef short SHORT;
#endif
#ifndef DWORD
typedef unsigned long       DWORD;
#endif

//!////////////////////////////////////////////////////////////////////////
#ifdef _DEBUG
#include <assert.h>
#define CHKV_RET(epr)        \
            assert (epr);    \
            if (!(epr))  return;
#define CHKV_RET_VAL(epr,val)        \
            assert (epr);    \
            if (!(epr))  return val;
#define CHKV_BREAK(epr)        \
            assert (epr);    \
            if (!(epr))  break;
#define CHKV_CONTINUE(epr)        \
            assert (epr);    \
            if (!(epr))  continue;
#define DEFAULT_FAILD  default: assert (FALSE); break
#else
#define CHKV_RET(epr)        \
            if (!(epr))  return;
#define CHKV_RET_VAL(epr,val)        \
            if (!(epr))  return val;
#define CHKV_BREAK(epr)        \
            if (!(epr))  break;
#define CHKV_CONTINUE(epr)        \
            if (!(epr))  continue;
#define DEFAULT_FAILD
#endif

// 用作支持编译fnr
#ifdef __COMPILE_FNR
#define _WT(text)                       _T("")
#define _CMDS_COUNT(CmdsCount)          0
#define _CMDS_PTR(CmdsPtr)              NULL
#define _EVENTS_COUNT(EventsCount)      0
#define _EVENTS_PTR(EventsPtr)          NULL
#define _ELEMENTS_COUNT(ElementsCount)  0
#define _ELEMENTS_PTR(ElementsPtr)      NULL
#define _FORM_BASE_CLASS                CWnd
#else
#define _WT(text)                       _T(text)
#define _CMDS_COUNT(CmdsCount)          CmdsCount
#define _CMDS_PTR(CmdsPtr)              CmdsPtr
#define _EVENTS_COUNT(EventsCount)      EventsCount
#define _EVENTS_PTR(EventsPtr)          ((PEVENT_INFO2)EventsPtr)
#define _ELEMENTS_COUNT(ElementsCount)  ElementsCount
#define _ELEMENTS_PTR(ElementsPtr)      ElementsPtr
#define _FORM_BASE_CLASS                CFrameWnd
#endif

typedef    SHORT    DTBOOL;             // SDT_BOOL类型的值类型
typedef    DTBOOL*  PDTBOOL;
#define    BL_TRUE  -1                  // SDT_BOOL类型的真值
#define    BL_FALSE 0                   // SDT_BOOL类型的假值

typedef DATE * PDATE;

//+////////////////////////////////////////////////////////////////////////

// 以下是由系统定义的基本数据类型

#define        _SDT_NULL        (DATA_TYPE)0           // 空数据(内部使用,必须为零)

/*  仅用于库命令定义其参数或返回值的数据类型。
    1、当用于定义库命令参数时,_SDT_ALL可以匹配所有数据类型(数组类型必须符合要求)。
    2、由于定义为返回_SDT_ALL数据类型的库命令不允许返回数组或复合数据类型的数据(即用户或库自定义数据类型但不包含窗口或菜单组件),
        所以_SDT_ALL类型的数据只可能为非数组的系统数据类型或窗口组件、菜单数据类型。*/
#define        _SDT_ALL        (DATA_TYPE)MAKELONG (MAKEWORD (0, 0), 0x8000)    // (内部使用)

    // 数值,匹配所有类型数字,仅用于库命令定义其参数及返回值的数据类型。
//#define _SDT_NUM        (DATA_TYPE)MAKELONG (MAKEWORD (1, 0), 0x8000)        // (内部使用),3.0版本中已经废弃。
#define SDT_BYTE        (DATA_TYPE)MAKELONG (MAKEWORD (1, 1), 0x8000)        // 字节
#define SDT_SHORT       (DATA_TYPE)MAKELONG (MAKEWORD (1, 2), 0x8000)        // 短整数
#define SDT_INT         (DATA_TYPE)MAKELONG (MAKEWORD (1, 3), 0x8000)        // 整数
#define SDT_INT64       (DATA_TYPE)MAKELONG (MAKEWORD (1, 4), 0x8000)        // 长整数
#define SDT_FLOAT       (DATA_TYPE)MAKELONG (MAKEWORD (1, 5), 0x8000)        // 小数
#define SDT_DOUBLE      (DATA_TYPE)MAKELONG (MAKEWORD (1, 6), 0x8000)        // 双精度小数
#define SDT_BOOL        (DATA_TYPE)MAKELONG (MAKEWORD (2, 0), 0x8000)        // 逻辑
#define SDT_DATE_TIME   (DATA_TYPE)MAKELONG (MAKEWORD (3, 0), 0x8000)        // 日期时间
#define SDT_TEXT        (DATA_TYPE)MAKELONG (MAKEWORD (4, 0), 0x8000)        // 文本
#define SDT_BIN         (DATA_TYPE)MAKELONG (MAKEWORD (5, 0), 0x8000)        // 字节集
#define SDT_SUB_PTR     (DATA_TYPE)MAKELONG (MAKEWORD (6, 0), 0x8000)        // 子程序指针
//#define _SDT_VAR_REF    (DATA_TYPE)MAKELONG (MAKEWORD (7, 0), 0x8000)        // 参考,3.0版本中已经废弃。
    
    /* 
    子语句型,仅用于库命令定义其参数的数据类型。其数据长度为两个INT,第一个记录存根子程序地址,第二个记录该子语句所在子程序的变量栈首。
    注意在用作库命令参数时,编译器允许其可以接收所有基本数据类型,所以必须首先判断处理
    调用例子：
        if (pArgInf->m_dtDataType == SDT_BOOL)
            return pArgInf->m_bool;

        if (pArgInf->m_dtDataType == SDT_STATMENT)
        {
            DWORD dwEBP = pArgInf->m_statment.m_dwSubEBP;
            DWORD dwSubAdr = pArgInf->m_statment.m_dwStatmentSubCodeAdr;
            DWORD dwECX, dwEAX;

            _asm
            {
                mov eax, dwEBP
                call dwSubAdr
                mov dwECX, ecx
                mov dwEAX, eax
            }

            if (dwECX == SDT_BOOL)
                return dwEAX != 0;

            // 释放文本或字节集数据所分配的内存。
            if (dwECX == SDT_TEXT || dwECX == SDT_BIN)
                MFree ((void*)dwEAX);
        }

        GReportError ("用作进行条件判断的子语句参数只能接受逻辑型数据");
    */
#define SDT_STATMENT    (DATA_TYPE)MAKELONG (MAKEWORD (8, 0),    0x8000)


//+////////////////////////////////////////////////////////////////////////

// 用作区分系统类型、用户自定义类型、库定义数据类型
#define DTM_SYS_DATA_TYPE_MASK          0x80000000
#define DTM_USER_DATA_TYPE_MASK         0x40000000
#define DTM_LIB_DATA_TYPE_MASK          0x00000000

// 用作细分用户自定义数据类型
#define UDTM_USER_DECLARE_MASK          0x00000000    // 用户自定义复合数据类型
#define UDTM_WIN_DECLARE_MASK           0x10000000    // 用户自定义窗口类型


//+////////////////////////////////////////////////////////////////////////


// 在数据类型中的数组标志,如果某数据类型值此位置1,则表示为此数据类型的数组。
// 本标志仅用作在运行时为具有AS_RECEIVE_VAR_OR_ARRAY或AS_RECEIVE_ALL_TYPE_DATA
// 标志的库命令参数说明其为是否为数组数据,其他场合均未使用。因此其他地方均可以忽略本标志。
#define DT_IS_ARY                       0x20000000

// 在数据类型中的传址标志,如果某数据类型值此位置1,则表示为此数据类型的变量地址。
// 本标志仅用作在运行时为具有AS_RECEIVE_VAR_OR_OTHER标志的库命令参数说明其为是否为
// 变量地址,其他场合均未使用。因此其他地方均可以忽略本标志。
// 本标志与上标志不能共存。
#define DT_IS_VAR                       0x20000000

typedef DWORD DATA_TYPE;
typedef DATA_TYPE* PDATA_TYPE;

//+//////////////////////////////////////////////////////////////

// 程序的版本类型宏
#define PT_EDIT_VER                 1    // 为用作编辑的版本
#define PT_DEBUG_RUN_VER            2    // 为DEBUG调试运行版本
#define PT_RELEASE_RUN_VER          3    // 为RELEASE最终运行版本

// 命令的学习难度级别
#define LVL_SIMPLE                  1    // 初级命令
#define LVL_SECONDARY               2    // 中级命令
#define LVL_HIGH                    3    // 高级命令


//+////////////////////////////////////////////////////////////////////////

// 参数数据提供方式有四种：
//        1、立即数提供方式(不能用于库或用户自定义类型);
//        2、单一变量提供方式(提供基本或自定义数据类型变量、变量数组元素、 自定义变量的成员变量等);
//        3、整个变量数组提供方式(提供整个变量数组的所有元素);
//        4、参数命令提供方式(由命令的返回值来提供参数数据)。
typedef struct
{
    LPCSTR          m_szName;               // 参数名称
    LPCSTR          m_szExplain;            // 参数详细解释
    SHORT           m_shtBitmapIndex;       // 指定图像索引,从1开始,0表示无.
    SHORT           m_shtBitmapCount;       // 图像数目(用作动画).

    DATA_TYPE       m_dtType;

    // 系统基本类型参数的默认指定值(在编辑程序时已被处理,编译时不需再处理)：
    //    1、数字型：直接为数字值(如为小数,只能指定其整数部分, 如为长整数,只能指定不超过INT限值的部分);
    //    2、逻辑型：1等于真,0等于假;
    //    3、文本型：本变量此时为LPSTR指针,指向默认文本串;
    //    4、其它所有类型参数(包括自定义类型)一律无默认指定值。
    INT            m_nDefault;

#define AS_HAS_DEFAULT_VALUE            (1 << 0)    // 本参数有默认值,默认值在m_nDefault中说明。本参数在编辑程序时已被处理,编译时不需再处理。
#define AS_DEFAULT_VALUE_IS_EMPTY       (1 << 1)    // 本参数有默认值,默认值为空,与AS_HAS_DEFAULT_VALUE标志互斥,运行时所传递过来的参数数据类型可能为_SDT_NULL。
#define AS_RECEIVE_VAR                  (1 << 2)    // 为本参数提供数据时只能提供单一变量,而不能提供整个变量数组、立即数或命令返回值。运行时所传递过来的参数数据肯定是内容不为数组的变量地址。
#define AS_RECEIVE_VAR_ARRAY            (1 << 3)    // 为本参数提供数据时只能提供整个变量数组,而不能提供单一变量、立即数或命令返回值。
#define AS_RECEIVE_VAR_OR_ARRAY         (1 << 4)    // 为本参数提供数据时只能提供单一变量或整个变量数组,而不能提供立即数或命令返回值。如果具有此标志,则传递给库命令参数的数据类型将会通过DT_IS_ARY来标志其是否为数组。
#define AS_RECEIVE_ARRAY_DATA           (1 << 5)    // 为本参数提供数据时只能提供数组数据,如不指定本标志,默认为只能提供非数组数据。如指定了本标志,运行时所传递过来的参数数据肯定为数组。
#define AS_RECEIVE_ALL_TYPE_DATA        (1 << 6)    // 为本参数提供数据时可以同时提供非数组或数组数据,与上标志互斥。如果具有此标志,则传递给库命令参数的数据类型将会通过DT_IS_ARY来标志其是否为数组。
#define AS_RECEIVE_VAR_OR_OTHER         (1 << 9)    // 为本参数提供数据时可以提供单一变量或立即数或命令返回值,不能提供数组。如果具有此标志,则传递给库命令参数的数据类型将会通过DT_IS_VAR来标志其是否为变量地址。
    DWORD        m_dwState;                // 参数MASK

}ARG_INFO, *PARG_INFO;

//////////////////////////////////////////////////////////////////////////

#ifndef __GCC_
struct CMD_INFO
#else
typedef struct
#endif
{
    LPCSTR          m_szName;           // 命令中文名称
    LPCSTR          m_szEgName;         // 命令英文名称,可以为空或NULL。
    LPCSTR          m_szExplain;        // 命令详细解释
    SHORT           m_shtCategory;      // 全局命令的所属类别,从1开始。对象成员命令的此值为-1。

#define CT_IS_HIDED                 (1 << 2)    // 本命令是否为隐含命令(即不需要由用户直接插入的命令,如循环结束命令、被废弃但为了保持兼容性又要存在的命令)。
#define CT_IS_ERROR                 (1 << 3)    // 本命令在本库中不能使用,具有此标志一定隐含,主要用作在不同语言版本的相同库中使用, 即：A命令在A语言版本库中可能需要实现并使用,但在B语言版本库中可能就不需要。如果程序中使用了具有此标志的命令,则只能支持该程序调入和编译,而不能支持运行。如具有此标志,本命令可以不实现其执行部分。
#define CT_DISABLED_IN_RELEASE      (1 << 4)    // 本命令在RELEASE版本程序中不被使用(相当于空命令),本类型命令必须无返回值。
#define CT_ALLOW_APPEND_NEW_ARG     (1 << 5)    // 在本命令的参数表的末尾是否可以添加新的参数,新参数等同于参数表中的最后一个参数, 本类型命令的参数必须最少有一个。
#define CT_RETRUN_ARY_TYPE_DATA     (1 << 6)    // 用于说明m_dtRetValType,说明是否为返回数组数据。

    /* 说明本命令为某数据类型的复制函数(执行将另一同类型数据类型数据复制到本对象时所需要的额外操作,在编译此数据类型对象的赋值语句时编译器先释放该对象的原有内容,
    然后生成调用此命令的代码,而不会再生成其它任何常规赋值及复制代码)。
        1、此命令必须仅接受一个同数据类型参数而且不返回任何数据。
        2、执行本命令时对象的内容数据为未初始化状态,命令内必须负责初始化其全部成员数据。
        3、所提供过来的待复制数据类型参数数据可能为全零状态(由编译器自动生成的对象初始代码设置), 复制时需要对此情况进行区别处理。
    */
#define CT_IS_OBJ_COPY_CMD          (1 << 7)    //! 注意每个数据类型最多只能有一个方法具有此标记。

    /* 说明本命令为某数据类型的析构函数(执行当该数据类型数据销毁时所需要的全部操作,
    当对象超出其作用区域时编译器仅仅生成调用此命令的代码,而不会生成任何常规销毁代码)。
        1、此命令必须没有任何参数而且不返回任何数据。
        2、此命令被执行时对象的内容数据可能为全零状态(由编译器自动生成的对象初始代码设置), 释放时需要对此情况进行区别处理。
    */
#define CT_IS_OBJ_FREE_CMD          (1 << 8)   //! 注意每个数据类型最多只能有一个方法具有此标记。

    /*
    说明本命令为某数据类型的构造函数(执行当该数据类型数据初始化时所需要的全部操作,
        1、此命令必须没有任何参数而且不返回任何数据。
        2、此命令被执行时对象的内容数据为全零状态。
        3、指定类型成员(复合数据类型成员、数组成员),必须按照对应格式继续进行下一步初始化。
    */
#define CT_IS_OBJ_CONSTURCT_CMD     (1 << 9)   //! 注意每个数据类型最多只能有一个方法具有此标记。
#define _CMD_OS(os)     ((os) >> 16)  // 用作转换os类型以便加入到m_wState。
#define _TEST_CMD_OS(m_wState,os)    ((_CMD_OS (os) & m_wState) != 0) // 用作测试指定命令是否支持指定操作系统。
    /* ! 千万注意：如果返回值类型为 _SDT_ALL , 绝对不能返回数组(即CT_RETRUN_ARY_TYPE_DATA
       置位)或复合数据类型的数据(即用户或库自定义数据类型但不包含窗口或菜单组件),
       因为无法自动删除复合类型中所分配的额外空间(如文本型或者字节集型成员等).
       由于通用型数据只可能通过库命令返回,因此所有_SDT_ALL类型的数据只可能为非数组的
       系统数据类型或窗口组件、菜单数据类型。
   */
    WORD            m_wState;           // 标记, CT_开头常量
    DATA_TYPE       m_dtRetValType;     // 返回值类型,使用前注意转换HIWORD为0的内部数据类型值到程序中使用的数据类型值。
    WORD            m_wReserved;        // 保留字段
    SHORT           m_shtUserLevel;     // 命令的用户学习难度级别,本变量的值为级别宏。

#ifndef __GCC_
    BOOL IsInObj()
    {
        return m_shtCategory == -1;
    }
#endif

    SHORT           m_shtBitmapIndex;   // 指定图像索引,从1开始,0表示无.
    SHORT           m_shtBitmapCount;   // 图像数目(用作动画).
    INT             m_nArgCount;        // 命令的参数数目
    PARG_INFO       m_pBeginArgInfo;    // 参数起始地址
#ifndef __GCC_
};
#else
} CMD_INFO;
#endif
typedef CMD_INFO* PCMD_INFO;

//////////////////////////////////////////////////////////////////////////

// 数据类型使用的结构。
typedef struct
{
    DATA_TYPE   m_dtType;       // 子数据组件的数据类型, 如果位于枚举数据类型中,则本成员必须为SDT_INT。 
    LPBYTE      m_pArySpec;     // 数组指定串,如果不为数组,此值为NULL。注意绝对不能指定某维数上限为0的数组。如果位于枚举数据类型中,则本成员必须为NULL。
    LPCSTR      m_szName;       // 子数据组件的中文变量名称,如果所属的数据类型只有一个子数据组件,则此值应该为NULL。
    LPCSTR      m_szEgName;     // 子数据组件的英文变量名称,可以为空或NULL。
    LPCSTR      m_szExplain;    // 详细说明


#define LES_HAS_DEFAULT_VALUE    (1 << 0)    //! 如果位于枚举数据类型中,则本成员默认具有LES_HAS_DEFAULT_VALUE标记。子数据成员有默认值,默认值在m_nDefault中说明。
#define LES_HIDED                (1 << 1)    // 本子数据成员被隐藏。
    DWORD m_dwState;    //! 如果位于枚举数据类型中,则本成员默认具有LES_HAS_DEFAULT_VALUE标记。子数据成员有默认值,默认值在m_nDefault中说明。

    // 系统基本类型子数据组件(非数组)的默认指定值：
    //        1、数字型：直接为数字值(如为小数,只能指定其整数部分,如为长整数,只能指定不超过INT限值的部分);
    //        2、逻辑型：1等于真,0等于假;
    //        3、文本型：本变量此时为LPSTR指针,指向默认文本串;
    //        4、其它所有类型参数(包括自定义类型)一律无默认指定值。
    INT m_nDefault;    // !!! 如果位于枚举数据类型中,则本成员为具体的枚举数值。

} LIB_DATA_TYPE_ELEMENT;
typedef LIB_DATA_TYPE_ELEMENT* PLIB_DATA_TYPE_ELEMENT;

//////////////////////////////////////////////////////////////////////////

// 固定属性的数目
#define    FIXED_WIN_UNIT_PROPERTY_COUNT    8

// 每个固定属性定义
#define    FIXED_WIN_UNIT_PROPERTY    \
    {    "左边",    "left",        NULL,    UD_INT,     _PROP_OS (OS_ALL),    NULL },    \
    {    "顶边",    "top",         NULL,    UD_INT,     _PROP_OS (OS_ALL),    NULL },    \
    {    "宽度",    "width",       NULL,    UD_INT,     _PROP_OS (OS_ALL),    NULL },    \
    {    "高度",    "height",      NULL,    UD_INT,     _PROP_OS (OS_ALL),    NULL },    \
    {    "标记",    "tag",         NULL,    UD_TEXT,    _PROP_OS (OS_ALL),    NULL },    \
    {    "可视",    "visible",     NULL,    UD_BOOL,    _PROP_OS (OS_ALL),    NULL },    \
    {    "禁止",    "disable",     NULL,    UD_BOOL,    _PROP_OS (OS_ALL),    NULL },    \
    {    "鼠标指针","MousePointer", NULL,    UD_CURSOR,  _PROP_OS (OS_ALL),    NULL }


// 组件的设置属性,注意不要增加数组属性支持。
typedef struct
{
    LPCSTR m_szName;    // 属性名称,注意为利于在属性表中同时设置多对象的属性,属性名称必须高度一致。
    LPCSTR m_szEgName;    // 属性英文名
    LPCSTR m_szExplain;    // 属性解释。

// 注意：数据格式仅在PFN_NOTIFY_PROPERTY_CHANGED中用作通知。
#define UD_PICK_SPEC_INT        1000    // 数据为INT值,用户只能选择,不能编辑。
#define UD_INT                  1001    // 数据为INT值
#define UD_DOUBLE               1002    // 数据为DOUBLE值
#define UD_BOOL                 1003    // 数据为BOOL值
#define UD_DATE_TIME            1004    // 数据为DATE值
#define UD_TEXT                 1005    // 数据为字符串
#define UD_PICK_INT             1006    // 数据为INT值,用户只能选择,不能编辑。
#define UD_PICK_TEXT            1007    // 数据为字符串,用户只能选择,不能编辑。
#define UD_EDIT_PICK_TEXT       1008    // 数据为字符串,用户可以编辑。
#define UD_PIC                  1009    // 为图片文件内容
#define UD_ICON                 1010    // 为图标文件内容
#define UD_CURSOR               1011    // 第一个INT记录鼠标指针类型,具体值见LoadCursor函数。如为-1,则为自定义鼠标指针,此时后跟相应长度的鼠标指针文件内容。
#define UD_MUSIC                1012    // 为声音文件内容
#define UD_FONT                 1013    // 为一个LOGFONT数据结构,不能再改。
#define UD_COLOR                1014    // 数据为COLORREF值。
#define UD_COLOR_TRANS          1015    // 数据为COLORREF值,允许透明颜色(用CLR_DEFAULT代表)。
#define UD_FILE_NAME            1016    // 数据为文件名字符串。此时m_szzPickStr中的数据为: 对话框标题\0 + 文件过滤器串\0 + 默认后缀\0 + "1"(取保存文件名)或"0"(取读入文件名)\0
#define UD_COLOR_BACK           1017    // 数据为COLORREF值,允许系统默认背景颜色(用CLR_DEFAULT代表)。
#define UD_IMAGE_LIST           1023    // 图片组,数据结构为：DWORD: 标志数据：为 IMAGE_LIST_DATA_MARK, COLORREF: 透明颜色(可以为CLR_DEFAULT), 后面为图片组数据.用CImageList::Read和CImageList::Write读写。
#define IMAGE_LIST_DATA_MARK    (MAKELONG ('IM', 'LT'))
#define UD_CUSTOMIZE            1024

#define UD_BEGIN                UD_PICK_SPEC_INT
#define UD_END                  UD_CUSTOMIZE

    SHORT m_shtType;    // 属性的数据类型。

#define UW_HAS_INDENT           (1 << 0)    // 在属性表中显示时向外缩进一段,一般用于子属性。
#define UW_GROUP_LINE           (1 << 1)    // 在属性表中本属性下显示分组底封线。
#define UW_ONLY_READ            (1 << 2)    // 只读属性,设计时不可用,运行时不能写。
#define UW_CANNOT_INIT          (1 << 3)    // 设计时不可用,但运行时可以正常读写。与上标志互斥。
#define UW_IS_HIDED             (1 << 4)    // 3.2 新增 隐藏但可用。
//!!! 注意高位包含 __OS_xxxx 宏用于指定本属性所支持的操作系统。
#define _PROP_OS(os)            ((os) >> 16)// 用作转换os类型以便加入到m_wState。
#define _TEST_PROP_OS(m_wState,os)    ((_PROP_OS (os) & m_wState) != 0) // 用作测试指定属性是否支持指定操作系统。
    WORD m_wState;        // UW_开头常量

    // 顺序记录所有的备选文本(除开UD_FILE_NAME),以一个空串结束。
    // 当m_nType为UP_PICK_INT、UP_PICK_TEXT、UD_EDIT_PICK_TEXT、UD_FILE_NAME时不为NULL。
    // 当m_nType为UD_PICK_SPEC_INT时,每一项备选文本的格式为 数值文本 + "\0" + 说明文本 + "\0" 。
    LPCSTR m_szzPickStr;
} UNIT_PROPERTY, * PUNIT_PROPERTY;

//////////////////////////////////////////////////////////////////////////


///////////////////////////////// 3.2 以前所使用的事件定义信息第一个版本

typedef struct
{
#define EAS_IS_BOOL_ARG (1 << 0)    // 为逻辑型参数,如无此标志,默认为整数型参数
    LPCSTR m_szName;        // 参数名称
    LPCSTR m_szExplain;     // 参数详细解释
    DWORD  m_dwState;       // EAS_IS_BOOL_ARG,为逻辑型参数, 如无此标志,默认为整数型参数
}EVENT_ARG_INFO, * PEVENT_ARG_INFO;

typedef struct
{
#define EV_IS_HIDED         (1 << 0)// 本事件是否为隐含事件(即不能被一般用户所使用或被废弃但为了保持兼容性又要存在的事件)。
    //#define EV_IS_MOUSE_EVENT    (1 << 1)    // 3.2被废弃。
#define EV_IS_KEY_EVENT     (1 << 2)
#define EV_RETURN_INT       (1 << 3)    // 返回一个整数
#define EV_RETURN_BOOL      (1 << 4)    // 返回一个逻辑值,与上标志互斥。
#define _EVENT_OS(os)       ((os) >> 1)    //! 注意高位包含 __OS_xxxx 宏用于指定本事件所支持的操作系统。用作转换os类型以便加入到m_dwState。
#define _TEST_EVENT_OS(m_dwState,os) ((_EVENT_OS (os) & m_dwState) != 0) // 用作测试指定事件是否支持指定操作系统。
    LPCSTR m_szName;        // 事件名称
    LPCSTR m_szExplain;     // 事件详细解释
    DWORD  m_dwState;       // 返回值类型, EV_ 开头常量, _EVENT_OS() 绝对不能定义成返回文本、字节集、复合类型等需要空间释放代码的数据类型。
    INT    m_nArgCount;     // 事件的参数数目
    PEVENT_ARG_INFO m_pEventArgInfo;    // 事件参数地址
}EVENT_INFO, * PEVENT_INFO;

//////////////////////////////////////////////////////////////////////////

///////////////////////////////// 3.2 及以后所使用的事件定义信息第二个版本

// 注意首部必须完全包含 EVENT_ARG_INFO 。
typedef struct
{
#define EAS_BY_REF (1 << 1)     // 不使用 (1 << 0)
    LPCSTR m_szName;            // 参数名称
    LPCSTR m_szExplain;         // 参数详细解释
    DWORD  m_dwState;           // EAS_BY_REF, 是否需要以参考方式传值,如果置位,则支持库中抛出事件的代码必须确保其能够被系统所访问(即分配内存的方法和数据的格式必须符合要求)。
    DATA_TYPE m_dtDataType;     // 类型
}EVENT_ARG_INFO2, * PEVENT_ARG_INFO2;


//////////////////////////////////////////////////////////////////////////
// 注意首部必须完全包含 EVENT_INFO 。
typedef struct
{
    LPCSTR m_szName;            // 事件名称
    LPCSTR m_szExplain;         // 事件详细解释
    // 以下基本状态值宏与 EVENT_INFO 中的定义相同。
    // #define EV_IS_HIDED        (1 << 0)    // 本事件是否为隐含事件(即不能被一般用户所使用或被废弃但为了保持兼容性又要存在的事件)。
    // #define EV_IS_KEY_EVENT    (1 << 2)    
    // 注意高位包含 __OS_xxxx 宏用于指定本事件所支持的操作系统。
    // #define _EVENT_OS(os)    ((os) >> 1)  // 用作转换os类型以便加入到m_dwState。
    // #define _TEST_EVENT_OS(m_dwState,os)    ((_EVENT_OS (os) & m_dwState) != 0) // 用作测试指定事件是否支持指定操作系统。
#define EV_IS_VER2 (1 << 31)    // 表示本结构为EVENT_INFO2,!!!使用本结构时必须加上此状态值。

    DWORD  m_dwState;           // 返回值类型, 最好带上 EV_IS_VER2, EV_ 开头常量, _EVENT_OS() 绝对不能定义成返回文本、字节集、复合类型等需要空间释放代码的数据类型。
    INT    m_nArgCount;         // 事件的参数数目
    PEVENT_ARG_INFO2 m_pEventArgInfo;   // 事件参数
    DATA_TYPE m_dtRetDataType;          //! 如果该数据类型有额外的数据需要释放,需要由支持库中抛出事件的代码负责将其释放。
}EVENT_INFO2, *PEVENT_INFO2;

//////////////////////////////////////////////////////////////////////////

typedef DWORD  HUNIT;

// 通用接口指针。
typedef void (WINAPI* PFN_INTERFACE) ();

// 取指定的接口。
#define ITF_CREATE_UNIT                     1       // 创建组件
#define ITF_PROPERTY_UPDATE_UI              2       // 指定属性目前可否被修改
#define ITF_DLG_INIT_CUSTOMIZE_DATA         3       // 使用对话框设置附加定制数据
#define ITF_NOTIFY_PROPERTY_CHANGED         4       // 通知某属性数据被用户修改
#define ITF_GET_ALL_PROPERTY_DATA           5       // 取全部属性数据
#define ITF_GET_PROPERTY_DATA               6       // 取某属性数据
#define ITF_GET_ICON_PROPERTY_DATA          7       // 取窗口的图标属性数据(仅用于窗口)
#define ITF_IS_NEED_THIS_KEY                8       // 询问组件是否需要指定的按键信息,用作组件截获处理默认为系统处理的按键,如TAB、SHIFT+TAB、UP、DOWN等。
#define ITF_LANG_CNV                        9       // 组件数据语言转换
#define ITF_MSG_FILTER                      11      // 消息过滤
#define ITF_GET_NOTIFY_RECEIVER             12      // 取组件的附加通知接收者(PFN_ON_NOTIFY_UNIT)
typedef INT(WINAPI * PFN_ON_NOTIFY_UNIT) (INT nMsg, DWORD dwParam1, DWORD dwParam2);
#define NU_GET_CREATE_SIZE_IN_DESIGNER    0
// 取设计时组件被单击放置到窗体上时的默认创建尺寸.
// dwParam1: 类型: INT*, 返回宽度(单位像素)
// dwParam2: 类型: INT*, 返回高度(单位像素)
// 成功返回1,失败返回0.
typedef PFN_INTERFACE(WINAPI * PFN_GET_INTERFACE) (INT nInterfaceNO);

//+////////////////////////////////////////////////////////////////////////

// 创建组件,成功时返回窗口句柄,phUnit中返回组件句柄,失败返回NULL。
// hDesignWnd仅在blInDesignMode为真时才有效,为设计窗体的窗口句柄。
typedef HUNIT(WINAPI* PFN_CREATE_UNIT) (LPBYTE pAllData, INT nAllDataSize,
    DWORD dwStyle, HWND hParentWnd, UINT uID, HMENU hMenu, INT x, INT y, INT cx, INT cy,
    DWORD dwWinFormID, DWORD dwUnitID,            // 用作通知到系统
    HWND hDesignWnd, BOOL blInDesignMode);


// dwState可以为以下宏值的组合。
#define CNV_NULL        0
#define CNV_FONTNAME    (1 << 0)    // 为转换字体名的语言(由于可能变长, ps中必须保证有足够的空间存放转换后的字体名)。
typedef void (WINAPI * PFN_CNV)(char* ps, DWORD dwState, int nParam); 

// 组件数据语言转换,返回包含转换后组件数据的HGLOBAL句柄,失败返回NULL。
// nParam必须原值传递给fnCnv的对应参数。
typedef HGLOBAL(WINAPI* PFN_LANG_CNV) (LPBYTE pAllData, LPINT pnAllDataSize, PFN_CNV fnCnv, int nParam);  

//! 以下函数仅在设计时使用。

// 如果指定属性目前可以被操作,返回真,否则返回假。
typedef BOOL(WINAPI* PFN_PROPERTY_UPDATE_UI) (HUNIT hUnit, INT nPropertyIndex);

// 如有附加定制数据,使用对话框设置这些数据,根据设置修改内部数据及外形,
// 如果需要重新创建才能修改外形,返回真。如果pblModified不为NULL,在其中返回修改状态。
typedef BOOL(WINAPI* PFN_DLG_INIT_CUSTOMIZE_DATA) (HUNIT hUnit, INT nPropertyIndex, BOOL* pblModified, LPVOID pResultExtraData);


// 用作记录某属性的属性值。
union UNIT_PROPERTY_VALUE
{
    INT         m_int;          // UD_INT、UD_PICK_INT、UD_PICK_SPEC_INT
    DOUBLE      m_double;       // UD_DOUBLE
    BOOL        m_bool;         // UD_BOOL
    DATE        m_dtDateTime;   // UD_DATE_TIME
    COLORREF    m_clr;          // UD_COLOR、UD_COLOR_TRANS、UD_COLOR_BACK
    LPCSTR      m_szText;       // UD_TEXT、UD_PICK_TEXT、UD_EDIT_PICK_TEXT、UD_DATA_SOURCE_NAME、UD_DATA_PROVIDER_NAME、UD_DSCOL_NAME、UD_ODBC_CONNECT_STR、UD_ODBC_SELECT_STR
    LPCSTR      m_szFileName;   // UD_FILE_NAME

    struct
    {
        LPBYTE m_pData;
        INT    m_nDataSize;
    } m_data;                // UD_PIC、UD_ICON、UD_CURSOR、UD_MUSIC、UD_FONT、UD_CUSTOMIZE、UD_IMAGE_LIST


#ifndef __GCC_
    UNIT_PROPERTY_VALUE()
    {
        memset(this, 0, sizeof(UNIT_PROPERTY_VALUE));
    }

    BOOL IsSame(INT nDataType, UNIT_PROPERTY_VALUE& val)
    {
        switch (nDataType)
        {
        case UD_INT:
        case UD_PICK_INT:
        case UD_PICK_SPEC_INT:
            return m_int == val.m_int;
        case UD_DOUBLE:
            return m_double == val.m_double;
        case UD_BOOL:
            return m_bool == val.m_bool;
        case UD_DATE_TIME:
            return m_dtDateTime == val.m_dtDateTime;
        case UD_COLOR:
        case UD_COLOR_TRANS:
        case UD_COLOR_BACK:
            return m_clr == val.m_clr;
        case UD_TEXT:
        case UD_PICK_TEXT:
        case UD_EDIT_PICK_TEXT:
            /*case UD_DATA_SOURCE_NAME:
            case UD_DATA_PROVIDER_NAME:
            case UD_DSCOL_NAME:
            case UD_ODBC_CONNECT_STR:
            case UD_ODBC_SELECT_STR:*/
            return m_szText == NULL && val.m_szText == NULL ||
                m_szText != NULL && val.m_szText != NULL &&
                strcmp(m_szText, val.m_szText) == 0;
        case UD_FILE_NAME:
            return m_szFileName == NULL && val.m_szFileName == NULL ||
                m_szFileName != NULL && val.m_szFileName != NULL &&
                strcmp(m_szFileName, val.m_szFileName) == 0;
        case UD_PIC:
        case UD_ICON:
        case UD_CURSOR:
        case UD_MUSIC:
        case UD_FONT:
        case UD_CUSTOMIZE:
        case UD_IMAGE_LIST:
            if (m_data.m_nDataSize == val.m_data.m_nDataSize)
            {
                if (m_data.m_nDataSize == 0)
                    return TRUE;
                else
                    return memcmp(m_data.m_pData, val.m_data.m_pData, m_data.m_nDataSize) == 0;
            }
            break;
        default:
#ifdef _DEBUG
            assert(FALSE);
#endif
            break;
        }

        return FALSE;
    }
#endif
};
typedef union UNIT_PROPERTY_VALUE* PUNIT_PROPERTY_VALUE;

// 通知某一般属性(非定制属性)数据被用户修改,根据设置修改内部数据及外形,如果确实需要重新创建才能修改外形,返回真。任何情况下如果有相应提示信息,返回到ppszTipText中。
//! 注意：必须进行值的合法性校验。
typedef BOOL(WINAPI* PFN_NOTIFY_PROPERTY_CHANGED) (HUNIT hUnit, INT nPropertyIndex, PUNIT_PROPERTY_VALUE pPropertyVaule, LPCSTR* ppszTipText);


// 取某属性数据到pPropertyVaule中,成功返回真,否则返回假。
// !! 1、对于运行时可能被用户操作改变(不考虑直接调用API函数)的属性,在运行时必须取其现行实际值。
// !! 2、对于取回的文本或字节集等分配有额外内存的数据指针,必须自行释放,易程序并不会将其释放。
typedef BOOL(WINAPI* PFN_GET_PROPERTY_DATA) (HUNIT hUnit, INT nPropertyIndex, PUNIT_PROPERTY_VALUE pPropertyVaule);

// 取全部属性数据
typedef HGLOBAL(WINAPI* PFN_GET_ALL_PROPERTY_DATA) (HUNIT hUnit);

// 取窗口的图标属性数据(仅用于窗口)
typedef HGLOBAL(WINAPI* PFN_GET_ICON_PROPERTY_DATA) (LPBYTE pAllData, INT nAllDataSize);

// 询问组件是否需要指定的按键信息,如果需要,返回真,否则返回假。
typedef BOOL(WINAPI* PFN_IS_NEED_THIS_KEY) (HUNIT hUnit, WORD wKey);

// 消息过滤(仅对消息过滤窗口组件有效,即具有LDT_MSG_FILTER_CONTROL标记),被过滤掉返回真,否则返回假。
typedef BOOL(WINAPI* PFN_MESSAGE_FILTER) (void* pMsg); // Windows操作系统中pMsg为MSG*指针。

//////////////////////////////////////////////////////////////////////////

#define    UNIT_BMP_SIZE                24                      // 组件标志位图的宽度和高度。
#define    UNIT_BMP_BACK_COLOR          (RGB(192, 192, 192))    // 组件标志位图的背景颜色。

//! 注意m_pElementBegin与m_pPropertyBegin只能有一个不为NULL,m_nElementCount与m_nPropertyCount只能有一个不为0。
typedef struct  // 库定义数据类型结构
{
    LPCSTR m_szName;            // 数据类型的中文名称(如：“整数”,“高精度数”等等)。
    LPCSTR m_szEgName;          // 数据类型的英文名称(如：“int”,“double”等等),可为空或NULL。
    LPCSTR m_szExplain;         // 数据类型的详细解释,如无则可为NULL。
    INT m_nCmdCount;            // 本类型中提供的成员命令的数目(可为0)。
    LPINT m_pnCmdsIndex;        // 顺序记录本类型中所有成员命令在库的命令表中的索引值,可为NULL。

#define LDT_IS_HIDED                (1 << 0)    // 本类型是否为隐含类型(即不能由用户直接用作定义的类型,如被废弃但为了保持兼容性又要存在的类型)。
#define LDT_IS_ERROR                (1 << 1)    // 本类型在本库中不能使用,具有此标志一定隐含。即使具有此标志,本类型的类型数据也必须完整定义。
#define LDT_WIN_UNIT                (1 << 6)    // 是否为窗口或可在窗口内使用的组件,如此标志置位则m_nElementCount必为0,
#define LDT_IS_CONTAINER            (1 << 7)    // 是否为容器组件,如有此标志,LDT_WIN_UNIT必置位。
#define LDT_IS_TAB_UNIT             (1 << 8)    // 是否为TAB控件(选择夹) 4.0新增
#define LDT_IS_FUNCTION_PROVIDER    (1 << 15)   // 仅用作提供功能的窗口组件(如时钟),如此标志置位则LDT_WIN_UNIT必置位。具有此标志的组件尺寸固定为32*32,并且在运行时无可视外形。
#define LDT_CANNOT_GET_FOCUS        (1 << 16)   // 仅用作窗口组件,如此标志置位则表示此组件不能接收输入焦点,不能TAB键停留。
#define LDT_DEFAULT_NO_TABSTOP      (1 << 17)   // 仅用作窗口组件,如此标志置位则表示此组件默认不停留TAB键,使用本标志必须上标志置位。
#define LDT_ENUM                    (1 << 22)   // 是否为枚举数据类型。3.7新增。
#define LDT_MSG_FILTER_CONTROL      (1 << 5)    // 是否为消息过滤组件。

#define _DT_OS(os)     (os)  //! 注意高位包含 __OS_xxxx 宏用于指定本数据类型所支持的操作系统。 用作转换os类型以便加入到m_dwState。
#define _TEST_DT_OS(m_dwState,os)    ((_DT_OS (os) & m_dwState) != 0) // 用作测试指定数据类型是否支持指定操作系统。
    DWORD m_dwState;    // LDT_ 开头常量 _DT_OS() | LDT_

    ////////////////////////////////////////////
    //! 以下变量只有在为窗口、菜单组件且不为枚举数据类型时才有效。
    DWORD               m_dwUnitBmpID;      // 指定在库中的组件图像资源ID,0为无。在OCX包装库中,m_dwUnitBmpID指定备用图像资源ID。
    INT                 m_nEventCount;      // 组件事件数量
    PEVENT_INFO2        m_pEventBegin;      // 定义本组件的所有事件。
    INT                 m_nPropertyCount;   // 组件属性数量
    PUNIT_PROPERTY      m_pPropertyBegin;   // 指向组件的所有属性地址
    PFN_GET_INTERFACE   m_pfnGetInterface;  // 组件交互函数

    ////////////////////////////////////////////
    //! 以下变量只有在不为窗口、菜单组件或为枚举数据类型时才有效。
    INT    m_nElementCount;    // 本数据类型中子数据组件的数目。如为窗口、菜单组件,此变量值必为0。
    PLIB_DATA_TYPE_ELEMENT m_pElementBegin;  // 指向子数据成员数组。
} LIB_DATA_TYPE_INFO, * PLIB_DATA_TYPE_INFO;



//////////////////////////////////////////////////////////////////////////

typedef struct  // 库常量数据结构
{
#define    CT_NULL 0
#define    CT_NUM  1    // value sample: 3.1415926
#define    CT_BOOL 2    // value sample: 1
#define    CT_TEXT 3    // value sample: "abc"
    LPCSTR m_szName;    // 常量名
    LPCSTR m_szEgName;  // 英文名
    LPCSTR m_szExplain; // 说明
    SHORT  m_shtLayout; // 必须为 1
    SHORT  m_shtType;   // 常量类型, CT_NULL=空常量, CT_NUM=数值型,double保存值, CT_BOOL=逻辑型, CT_TEXT=文本型
    LPCSTR m_szText;    // CT_TEXT
    DOUBLE m_dbValue;   // CT_NUM、CT_BOOL
} LIB_CONST_INFO, * PLIB_CONST_INFO;

//+////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////// 常用数据结构。

typedef union
{
    BYTE    m_byte;
    SHORT   m_short;
    INT     m_int;
    INT64   m_int64;
    FLOAT   m_float;
    DOUBLE  m_double;
}SYS_NUM_VALUE, * PSYS_NUM_VALUE;

typedef struct
{
    DWORD m_dwFormID;
    DWORD m_dwUnitID;
}MUNIT, * PMUNIT;

typedef struct
{
    DWORD m_dwStatmentSubCodeAdr;    // 记录可重调用语句子程序代码的进入地址。
    DWORD m_dwSubEBP;    // 记录该子语句所处子程序的EBP指针,以便访问该子程序中的局部变量。
}STATMENT_CALL_DATA, * PSTATMENT_CALL_DATA;

#ifndef __GCC_
#pragma pack (push, old_value)   // 保存VC++编译器结构对齐字节数。
#pragma pack (1)    // 设置为以一字节对齐。
#endif

typedef struct
{
    union
    {
        BYTE   m_byte;            // SDT_BYTE
        SHORT  m_short;            // SDT_SHORT
        INT    m_int;            // SDT_INT
        DWORD  m_uint;            // (DWORD)SDT_INT
        INT64  m_int64;            // SDT_INT64
        FLOAT  m_float;            // SDT_FLOAT
        DOUBLE m_double;        // SDT_DOUBLE
        DATE   m_date;            // SDT_DATE_TIME
        BOOL   m_bool;            // SDT_BOOL
        char* m_pText;            // SDT_TEXT,不可能为NULL。为了避免修改到常量段(m_pText有可能指向常量段区域)中的数据,只可读取而不可更改其中的内容
        LPBYTE m_pBin;            // SDT_BIN,不可能为NULL,只可读取而不可更改其中的内容。
        DWORD  m_dwSubCodeAdr;    // SDT_SUB_PTR,记录子程序代码地址。
        STATMENT_CALL_DATA m_statment;    // SDT_STATMENT数据类型。
        MUNIT  m_unit;            // 窗口组件、菜单数据类型。
        void* m_pCompoundData;    // 复合数据类型数据指针,指针所指向数据的格式请见 run.h 。可以直接更改其中的数据成员,但是如果需要必须首先释放该成员。
        void* m_pAryData;        // 数组数据指针,指针所指向数据的格式请见 run.h 。注意如果为文本或字节集数组,则成员数据指针可能为NULL。只可读取而不可更改其中的内容。

        //! 为指向变量地址的指针,仅当传入参数到库命令实现函数时才有用。
        BYTE*   m_pByte;    // SDT_BYTE*
        SHORT*  m_pShort;    // SDT_SHORT*
        INT*    m_pInt;        // SDT_INT*
        DWORD*  m_pUInt;    // ((DWORD)SDT_INT)*
        INT64*  m_pInt64;    // SDT_INT64*
        FLOAT*  m_pFloat;    // SDT_FLOAT*
        DOUBLE* m_pDouble;    // SDT_DOUBLE*
        DATE*   m_pDate;    // SDT_DATE_TIME*
        BOOL*   m_pBool;    // SDT_BOOL*
        char**  m_ppText;    // SDT_TEXT,*m_ppText可能为NULL。注意写入新值之前必须释放前值,即：MFree (*m_ppText)。不可直接更改*m_ppText所指向的内容,只能释放原指针后换入新指针。
        LPBYTE* m_ppBin;    // SDT_BIN,*m_ppBin可能为NULL。注意写入新值之前必须释放前值,即：MFree (*m_ppBin)。不可直接更改*m_ppBin所指向的内容,只能释放原指针后换入新指针。
        DWORD*  m_pdwSubCodeAdr;            // SDT_SUB_PTR,子程序代码地址变量。
        PSTATMENT_CALL_DATA m_pStatment;    // SDT_STATMENT数据类型变量。
        PMUNIT m_pUnit;                // 窗口组件、菜单数据类型变量。
        void** m_ppCompoundData;    // 复合数据类型变量。可以直接更改其中的数据成员,但是如果需要必须首先释放该成员。
        void** m_ppAryData;            // 数组数据变量,注意：1、写入新值之前必须释放原值(使用NRS_FREE_VAR通知)。2、变量如果为文本或字节集数组,则成员数据指针可能为NULL。不可直接更改*m_ppAryData所指向的内容,只能释放原指针后换入新指针。
    };

    // 1、当用作传入参数时,如果该参数具有 AS_RECEIVE_VAR_OR_ARRAY 或AS_RECEIVE_ALL_TYPE_DATA 标志, 且为数组数据,则包含数组标志 DT_IS_ARY, 这也是 DT_IS_ARY 标志的唯一使用场合。
    // 2、当用作传递参数数据时,如果为空白数据,则为 _SDT_NULL 。
    DATA_TYPE m_dtDataType;
} MDATA_INF;
typedef MDATA_INF* PMDATA_INF;

#ifndef __GCC_
#pragma pack (pop, old_value)    // 恢复VC++编译器结构对齐字节数。
#endif


//////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////// 通知用数据结构。

/*/////////////*/
// 库可以通知系统的码值：

#ifndef __GCC_
struct MDATA
{
    MDATA()
    {
        m_pData = NULL;
        m_nDataSize = 0;
    }
#else
typedef struct
{
#endif

    LPBYTE    m_pData;
    INT        m_nDataSize;

#ifndef __GCC_
};
#else
} MDATA;
#endif
typedef MDATA* PMDATA;

// 用作第一类事件通知。
#ifndef __GCC_
struct EVENT_NOTIFY
#else
typedef struct
#endif
{
    // 记录事件的来源
    DWORD  m_dwFormID;        // 窗口 ID
    DWORD  m_dwUnitID;        // 窗口组件 ID
    INT    m_nEventIndex;    // 事件索引
    INT    m_nArgCount;        // 事件所传递的参数数目,最多 5 个。
    INT    m_nArgValue[5];    // 记录各参数值,SDT_BOOL 型参数值为 1 或 0。

    //!!! 注意下面两个成员在没有定义返回值的事件处理中无效。
    BOOL  m_blHasRetVal;    // 用户事件处理子程序处理完毕事件后是否有返回值
    INT   m_nRetVal;        // 用户事件处理子程序处理完毕事件后的返回值,逻辑值用数值 0(假) 和 1(真) 返回。

    //////////////////////////////////////////////////////////////////////////

#ifndef __GCC_
    EVENT_NOTIFY(DWORD dwFormID, DWORD dwUnitID, INT nEventIndex)
    {
        memset(this, 0, sizeof(EVENT_NOTIFY));

        m_dwFormID = dwFormID;
        m_dwUnitID = dwUnitID;
        m_nEventIndex = nEventIndex;

        m_nArgCount = 0;
        m_blHasRetVal = FALSE;
        m_nRetVal = 0;
    }
    EVENT_NOTIFY()
    {
        memset(this, 0, sizeof(EVENT_NOTIFY));
    }
};
#else
} EVENT_NOTIFY;
#endif
typedef EVENT_NOTIFY* PEVENT_NOTIFY;

typedef struct
{
#define EAV_IS_POINTER  (1 << 0)
#define EAV_IS_WINUNIT  (1 << 1)    //! 补充说明m_inf.m_dtDataType数据类型是否为窗口单元, 注意如果m_inf.m_dtDataType为窗口单元,此标记必须置位。
    MDATA_INF m_inf;

    //!!! 注意如果 m_inf.m_dtDataType 为文本型、字节集型、库定义数据类型(除开窗口单元及菜单数据类型), 必须传递指针,即本标记必须置位。
    DWORD m_dwState;    // m_inf中是否为指针数据。
}EVENT_ARG_VALUE, * PEVENT_ARG_VALUE;

// 用作第二类事件通知。
#ifndef __GCC_
struct EVENT_NOTIFY2
#else
typedef struct
#endif
{
    // 记录事件的来源
    DWORD m_dwFormID;    // 窗口 ID
    DWORD m_dwUnitID;    // 窗口组件 ID
    INT m_nEventIndex;    // 事件索引

#define MAX_EVENT2_ARG_COUNT    12
    INT m_nArgCount;    // 事件所传递的参数数目,最多 MAX_EVENT2_ARG_COUNT 个。
    EVENT_ARG_VALUE m_arg[MAX_EVENT2_ARG_COUNT];  // 记录各参数值。

    //!!! 注意下面成员在没有定义返回值的事件处理中无效。
    BOOL m_blHasRetVal;        // 用户事件处理子程序处理完毕事件后是否有返回值
    MDATA_INF m_infRetData;    // 记录用户事件处理子程序处理完毕事件后的返回值,注意其中的m_infRetData.m_dtDataType成员未被使用。

    //////////////////////////////////////////////////////////////////////////

#ifndef __GCC_
    EVENT_NOTIFY2(DWORD dwFormID, DWORD dwUnitID, INT nEventIndex)
    {
        memset(this, 0, sizeof(EVENT_NOTIFY2));
        m_dwFormID = dwFormID;
        m_dwUnitID = dwUnitID;
        m_nEventIndex = nEventIndex;
    }
    EVENT_NOTIFY2()
    {
        memset(this, 0, sizeof(EVENT_NOTIFY2));
    }
};
#else
} EVENT_NOTIFY2;
#endif
typedef EVENT_NOTIFY2* PEVENT_NOTIFY2;

typedef struct
{
    HICON m_hBigIcon;
    HICON m_hSmallIcon;
} APP_ICON, * PAPP_ICON;

//////////////////////////////////////////////////////////////////////////

//++ NES_ 宏为仅被易编辑环境处理的通知。

// 取易编辑环境主窗口的句柄,用作支持支持库的AddIn。
#define NES_GET_MAIN_HWND               1

// 通知易编辑环境运行指定的功能,返回一个BOOL值。
// dwParam1为功能号。
// dwParam2为一个双DWORD数组指针,分别提供功能参数1和2。
#define NES_RUN_FUNC                    2

// 通知易编辑环境显示一个对话框列出指定图片组内的所有图片,并返回用户所选择图片的索引号。
//        dwParam1为所欲浏览的有效的图片组句柄。
//        dwParam2如果不为NULL,则系统认为此为一个编辑框HWND窗口句柄,当用户做出有效选择后,系统将自动更改此编辑框的内容并将焦点转移上去。
// 返回用户所选择图片的索引号(-1表示用户选择无图片),如果用户未选择,则返回-2。
#define NES_PICK_IMAGE_INDEX_DLG        7


//++ NAS_ 宏为既被易编辑环境又被易运行环境处理的通知。

// 通知系统创建并返回程序的图标。
// dwParam1为PAPP_ICON指针。
#define NAS_GET_APP_ICON                1000

// 返回指定库定义数据类型的PLIB_DATA_TYPE_INFO定义信息指针。
// dwParam1为欲检查的数据类型。
// 如果该数据类型无效或者不为库定义数据类型,则返回NULL,否则返回PLIB_DATA_TYPE_INFO指针。
#define NAS_GET_LIB_DATA_TYPE_INFO      1002

// dwParam1为图片数据指针,dwParam2为图片数据尺寸。
// 如果成功返回非NULL的HBITMAP句柄(注意使用完毕后释放),否则返回NULL。
#define NAS_GET_HBITMAP                 1003

// 返回当前系统或运行环境所支持的语言ID,具体ID值请见lang.h
#define NAS_GET_LANG_ID                1004

// 返回当前系统或运行环境的版本号,LOWORD为主版本号,HIWORD为次版本号。
#define NAS_GET_VER                    1005

/* 返回当前开发或运行环境的某一类目录或文件名,目录名以“\”结束。
    dwParam1: 指定所需要的目录,可以为以下值：
        A、开发及运行环境下均有效的目录:
            1: 开发或运行环境系统所处的目录;
        B、开发环境下有效的目录(仅开发环境中有效):
            1001: 系统例程和支持库例程所在目录名
            1002: 系统工具所在目录
            1003: 系统帮助信息所在目录
            1004: 保存所有登记到系统中易模块的目录
            1005: 支持库所在的目录
            1006: 安装工具所在目录
        C、运行环境下有效的目录(仅运行环境中有效):
            2001: 用户EXE文件所处目录;
            2002: 用户EXE文件名;
    dwParam2: 接收缓冲区地址,尺寸必须为MAX_PATH。*/
#define NAS_GET_PATH                1006

// 通过指定HWND句柄创建一个CWND对象,返回其指针,记住此指针必须通过调用NRS_DELETE_CWND_OBJECT来释放
// dwParam1为HWND句柄
// 成功返回CWnd*指针,失败返回NULL
#define NAS_CREATE_CWND_OBJECT_FROM_HWND    1007

// 删除通过NRS_CREATE_CWND_OBJECT_FROM_HWND创建的CWND对象
// dwParam1为欲删除的CWnd对象指针
#define NAS_DELETE_CWND_OBJECT              1008

// 取消通过NRS_CREATE_CWND_OBJECT_FROM_HWND创建的CWND对象与其中HWND的绑定
// dwParam1为CWnd对象指针
// 成功返回HWND,失败返回0
#define NAS_DETACH_CWND_OBJECT              1009

// 获取通过NRS_CREATE_CWND_OBJECT_FROM_HWND创建的CWND对象中的HWND
// dwParam1为CWnd对象指针
// 成功返回HWND,失败返回0
#define NAS_GET_HWND_OF_CWND_OBJECT         1010

// 将指定HWND与通过NRS_CREATE_CWND_OBJECT_FROM_HWND创建的CWND对象绑定起来
// dwParam1为HWND
// dwParam2为CWnd对象指针
// 成功返回1,失败返回0
#define NAS_ATTACH_CWND_OBJECT              1011

// 如果指定窗口为易语言窗口或易语言组件,返回真,否则返回假。
// dwParam1为欲测试的HWND.
#define    NAS_IS_EWIN                            1014


//++ NRS_ 宏为仅能被易运行环境处理的通知。

// 通知系统指定的组件已经被销毁。
// dwParam1为dwFormID
// dwParam2为dwUnitID
#define NRS_UNIT_DESTROIED                    2000

// 转换其它数值格式到整数。
// dwParam1为 PMDATA_INF 指针,其 m_dtDataType 必须为数值型。返回转换后的整数值。
#define NRS_CONVERT_NUM_TO_INT              2001

// 取当前命令行文本
// 返回命令行文本指针,有可能为空串。
#define NRS_GET_CMD_LINE_STR                2002

// 取当前执行文件所处目录名称
// 返回当前执行文件所处目录文本指针。
#define NRS_GET_EXE_PATH_STR                2003

// 取当前执行文件名称
// 返回当前执行文件名称文本指针。
#define NRS_GET_EXE_NAME                    2004

// 取组件对象指针
// dwParam1为WinForm的ID
// dwParam2为WinUnit的ID
// 成功返回有效的组件对象CWnd*指针,失败返回0。
#define NRS_GET_UNIT_PTR                    2006

// 取组件对象指针
// dwParam1为WinForm的ID
// dwParam2为WinUnit的ID
// 成功返回有效的组件对象CWnd*指针,失败报告运行时错误并退出程序。
#define NRS_GET_AND_CHECK_UNIT_PTR            2007

// 以第一类方式通知系统产生了事件。
// dwParam1为PEVENT_NOTIFY指针。
// 如果返回 0 ,表示此事件已被系统抛弃,否则表示系统已经成功传递此事件到用户事件处理子程序。
#define NRS_EVENT_NOTIFY                    2008

// 通知系统发送所有待处理事件。
#define    NRS_DO_EVENTS                        2018

// dwParam1为WinForm的ID
// dwParam2为WinUnit的ID
// 成功返回有效的 DATA_TYPE ,失败返回 0 。
#define NRS_GET_UNIT_DATA_TYPE              2022

// 释放指定数组数据。
// dwParam1为该数据的DATA_TYPE,只能为系统数据类型。
// dwParam2为指向该数组数据的指针。
#define NRS_FREE_ARY                        2023

// 分配指定空间的内存,所有与易程序交互的内存都必须使用本通知分配。
// dwParam1为欲需求内存字节数。
// dwParam2如为0,则如果分配失败就自动报告运行时错并退出程序。如不为0,则如果分配失败就返回NULL。返回所分配内存的首地址。
#define NRS_MALLOC                          2024

// 释放已分配的指定内存。
// dwParam1为欲释放内存的首地址。
#define NRS_MFREE                           2025

// 重新分配内存。
// dwParam1为欲重新分配内存尺寸的首地址。
// dwParam2为欲重新分配的内存字节数。
// 返回所重新分配内存的首地址,失败自动报告运行时错并退出程序。
#define NRS_MREALLOC                        2026

// 通知系统已经产生运行时错误。
// dwParam1为char*指针,说明错误文本。
#define    NRS_RUNTIME_ERR                        2027

// 通知系统退出用户程序。
// dwParam1为退出代码,该代码将被返回到操作系统。
#define    NRS_EXIT_PROGRAM                    2028

// 返回当前用户程序的类型,为PT_DEBUG_RUN_VER(调试版)或PT_RELEASE_RUN_VER(发布版)。
#define NRS_GET_PRG_TYPE                    2030

// 以第二类方式通知系统产生了事件。
// dwParam1为PEVENT_NOTIFY2指针。
// 如果返回 0 ,表示此事件已被系统抛弃,否则表示系统已经成功传递此事件到用户事件处理子程序。
#define NRS_EVENT_NOTIFY2                    2031

// 返回当前程序的窗体数目。
#define NRS_GET_WINFORM_COUNT               2032

// 返回指定窗体的窗口句柄,如果该窗体尚未被载入,返回NULL。
// dwParam1为窗体索引。
#define NRS_GET_WINFORM_HWND                2033

// 返回指定HBITMAP的图片数据,成功返回包含BMP图片数据的HGLOBAL句柄,失败返回NULL。
// dwParam1为欲获取其图片数据的HBITMAP。
#define NRS_GET_BITMAP_DATA                 2034

// 通知系统释放指定的DTP_COM_OBJECT类型COM对象。
// dwParam1为该COM对象的地址指针。
#define NRS_FREE_COMOBJECT                  2035

// 当选择夹子夹被切换后, 使用本消息通知易系统
#define NRS_CHK_TAB_VISIBLE                 2039


//////////////////////////////////////////////////////////////////////////
//++ 系统可以通知库的码值：

    // 告知库通知系统用的函数指针,在装载支持库前通知,可能有多次,后通知的值应该覆盖前面所通知的值,忽略返回值。
    // 库可将此函数指针记录下来以便在需要时使用它通知信息到系统。
    // dwParam1: (PFN_NOTIFY_SYS)
#define NL_SYS_NOTIFY_FUNCTION              1

// 通知支持库释放资源准备退出及释放指定的附加数据。
#define NL_FREE_LIB_DATA                    6

// 返回所有命令实现函数的的函数名称数组(char*[]), 支持静态编译的动态库必须处理
#define NL_GET_CMD_FUNC_NAMES               14

// 返回处理系统通知的函数名称(PFN_NOTIFY_LIB函数名称), 支持静态编译的动态库必须处理
#define NL_GET_NOTIFY_LIB_FUNC_NAME         15

// 返回静态库所依赖的其它静态库文件名列表(格式为\0分隔的文本,结尾两个\0), 支持静态编译的动态库必须处理
// kernel32.lib user32.lib gdi32.lib 等常用的系统库不需要放在此列表中
// 返回NULL或NR_ERR表示不指定依赖文件
#define NL_GET_DEPENDENT_LIBS               16

// 当支持库在IDE中被用户取消选择后发送本通知
#define NL_UNLOAD_FROM_IDE                  17

// 返回NR_DELAY_FREE表示本支持库需要被延迟释放
#define NR_DELAY_FREE                       2015

// 当IDE已经准备好接受用户操作后,此通知被发送到所有具有LBS_IDE_PLUGIN标志的支持库.
// 注意: 当用户在支持库配置对话框中取消已选择支持库或选择新的支持库后,本通知会被再次发送,支持库内需要处理好此种情况.
#define NL_IDE_READY                    18

// 当用户在IDE中单击鼠标右键即将弹出菜单前,此通知被发送到所有具有LBS_IDE_PLUGIN标志的支持库.
// dwParam1: 欲弹出菜单句柄(HMENU)
// dwParam2: 欲弹出菜单在易语言主程序(e.exe)中的资源ID
#define NL_RIGHT_POPUP_MENU_SHOW        19

// 当IDE准备显示新加入成员对话框之前,此通知被发送到所有具有LBS_IDE_PLUGIN标志的支持库.
// 返回1表示该通知被处理且IDE根据返回内容添加所指定的新成员,返回2表示该通知被处理且IDE中止添加新成员,返回0表示未被处理.
// dwParam1: 提供欲加入成员的类型(ADD_NEW_ELEMENT_TYPE)
// dwParam2:
//   如果dwParam1为NAN_NEW_VAR,则为一个指向RES_NEW_VAR_INFO的指针;
//   如果dwParam1为NAN_NEW_SUB,则为一个指向欲加入子程序名称的文本指针;
//   如果dwParam1为NAN_NEW_CONST_RES,则为一个指向RES_NEW_CONST_RES_INFO的指针.
#define NL_ADD_NEW_ELEMENT                20


typedef enum
{
    NAN_NEW_VAR,            // 变量
    NAN_NEW_SUB,            // 子程序
    NAN_NEW_CONST_RES       // 常量资源
}
ADD_NEW_ELEMENT_TYPE;       // 欲加入新成员的类型
typedef enum
{
    NVET_GLOBAL_VAR,        // 全局变量
    NVET_MOD_VAR,           // 模块变量
    NVET_LOCAL_VAR          // 局部变量
}NEW_VAR_EXTENT_TYPE;       // 新变量的范围类型

typedef struct
{
    const TCHAR* m_szNewVarName;    // in 新变量的名称
    DATA_TYPE m_dtNewVar;    // out 用作返回新变量的数据类型
    NEW_VAR_EXTENT_TYPE m_enNewVarExtentType;    // out 用作返回新变量的范围类型
}RES_NEW_VAR_INFO;

typedef struct
{
    const TCHAR* m_szNewConstName;          // in 新常量的名称
    TCHAR m_acConstInitialValue[256];       // out 用作返回新常量的初始值
}RES_NEW_CONST_RES_INFO;

//+////////////////////////////////////////////////////////////////////////

#define NR_OK           0
#define NR_ERR          -1


typedef INT(WINAPI* PFN_NOTIFY_LIB) (INT nMsg, DWORD dwParam1, DWORD dwParam2);    // 此函数用作系统通知库有关事件。
typedef INT(WINAPI* PFN_NOTIFY_SYS) (INT nMsg, DWORD dwParam1, DWORD dwParam2);    // 此函数用作库通知系统有关事件。



/* 所有命令和方法实现函数的原型。
   1、必须是 CDECL 调用方式;
   2、pRetData 用作返回数据;
   3、如果指定库命令返回数据类型不为 _SDT_ALL ,可以不填充 pRetData->m_dtDataType,如果为 _SDT_ALL ,则必须填写;
   4、pArgInf 提供参数数据本身,所指向的 MDATA_INF 描述每个输入参数,数目等同于 nArgCount 。*/
typedef void (*PFN_EXECUTE_CMD) (PMDATA_INF pRetData, INT nArgCount, PMDATA_INF pArgInf);

typedef INT(WINAPI* PFN_RUN_ADDIN_FN) (INT nAddInFnIndex);        // 运行支持库中ADDIN功能的函数
typedef INT(WINAPI* PFN_SUPER_TEMPLATE) (INT nTemplateIndex);    // 创建库中提供的超级模板程序的函数

//////////////////////////////////////////////////////////////////////////

#define LIB_FORMAT_VER    20000101    // 库格式号

typedef struct
{
    DWORD  m_dwLibFormatVer;        // 库格式号,应该等于LIB_FORMAT_VER。主要用途如下：譬如 krnln.fnX 库,在做其它与此同名但功能完全不一致的库时,应当改变此格式号,以防止错误装载。
    LPCSTR m_szGuid;                // 对应于本库的唯一GUID串,不能为NULL或空,库的所有版本此串都应相同。如果为ActiveX控件,此串记录其CLSID。
    INT    m_nMajorVersion;         // 本库的主版本号,必须大于0。
    INT    m_nMinorVersion;         // 本库的次版本号。
    INT    m_nBuildNumber;          // 构建版本号,无需对此版本号作任何处理。本版本号仅用作区分相同正式版本号的系统软件(譬如仅仅修改了几个 BUG,不值得升级正式版本的系统软件)。任何公布过给用户使用的版本其构建版本号都应该不一样。赋值时应该顺序递增。
    INT    m_nRqSysMajorVer;        // 所需要的易语言系统的主版本号。
    INT    m_nRqSysMinorVer;        // 所需要的易语言系统的次版本号。
    INT    m_nRqSysKrnlLibMajorVer; // 所需要的系统核心支持库的主版本号。
    INT    m_nRqSysKrnlLibMinorVer; // 所需要的系统核心支持库的次版本号。

    LPCSTR m_szName;    // 库名,不能为NULL或空。
    INT    m_nLanguage;    // 库所支持的语言。
    LPCSTR m_szExplain;    // 库详细解释

#define LBS_FUNC_NO_RUN_CODE        (1 << 2)    // 本库仅为声明库,没有对应功能的支持代码,因此不能运行。
#define LBS_NO_EDIT_INFO            (1 << 3)    // 本库内无供编辑用的信息(编辑信息主要为：各种名称、解释字符串等)。
#define LBS_IS_DB_LIB               (1 << 5)    // 本库是否为数据库操作支持库。
#define LBS_LIB_INFO2               (1 << 7)    // 本信息结构实际上是否为LIB_INFO2
#define LBS_IDE_PLUGIN              (1 << 8)    // 本支持库是否为易语言IDE插件. 注意: 只有设置了此标志的支持库,才会接收到来自易语言IDE的NL_IDE_READY和NL_RIGHT_POPUP_MENU_SHOW等相关通知.

//!!! 注意高位包含 __OS_xxxx 宏用于表明本支持库具有的所有操作系统版本。
#define _LIB_OS(os)     (os)    // 用作转换os类型以便加入到m_dwState。
#define _TEST_LIB_OS(m_dwState,os)    ((_LIB_OS (os) & m_dwState) != 0) // 用作测试支持库是否具有指定操作系统的版本。
    DWORD  m_dwState;           // _LIB_OS() | LBS_ 开头常量, 如果是插件,必须包含 LBS_IDE_PLUGIN

    LPCSTR m_szAuthor;          // 作者
    LPCSTR m_szZipCode;         // 邮编
    LPCSTR m_szAddress;         // 地址
    LPCSTR m_szPhoto;           // 电话
    LPCSTR m_szFax;             // 传真
    LPCSTR m_szEmail;           // 邮箱
    LPCSTR m_szHomePage;        // 主页
    LPCSTR m_szOther;           // 其他信息

    //////////////////
    INT                 m_nDataTypeCount;   // 本库中自定义数据类型的数目。
    PLIB_DATA_TYPE_INFO m_pDataType;        // 本库中所有的自定义数据类型。可以参考使用系统核心支持库中的自定义数据类型,系统核心支持库在程序的库登记数组中的索引值加1后的值为1。
    INT                 m_nCategoryCount;   // 全局命令类别数目,可为0。
    LPCSTR              m_szzCategory;      // 全局命令类别说明表


    INT                 m_nCmdCount;        // 本库中提供的所有命令(全局命令及对象命令)的数目(可为0)。
    PCMD_INFO           m_pBeginCmdInfo;    // 可为NULL
    PFN_EXECUTE_CMD*    m_pCmdsFunc;        // 指向每个命令的实现代码首地址,可为NULL


    // 有关AddIn功能的说明,两个字符串说明一个功能。第一个为功能名称(仅限一行20字符,如果希望自行初始位置而不被自动加入到工具菜单,
    // 则名称应该以@开始,此时会接收到值为 -(nAddInFnIndex + 1) 的调用通知),
    // 第二个为功能详细介绍(仅限一行60字符),最后由两个空串结束。
    PFN_RUN_ADDIN_FN    m_pfnRunAddInFn;    // 可为NULL
    LPCSTR              m_szzAddInFnInfo;   // 功能详细介绍（仅限一行60字符），最后由两个空串结束。

    PFN_NOTIFY_LIB      m_pfnNotify;        // 不能为NULL,和易语言通讯的函数

    // 超级模板暂时保留不用。
    // 有关SuperTemplate的说明,两个字符串说明一个SuperTemplate。
    // 第一个为SuperTemplate名称(仅限一行30字符),第二个为详细介绍(不限),最后由两个空串结束。
    PFN_SUPER_TEMPLATE  m_pfnSuperTemplate;   // 可为NULL
    LPCSTR              m_szzSuperTemplateInfo;              // 详细介绍(不限长度), 最后由两个空串结束。

    INT                 m_nLibConstCount;            // 本库预先定义的所有常量数量。
    PLIB_CONST_INFO     m_pLibConst;    // 本库预先定义的所有常量。

    LPCSTR              m_szzDependFiles;        // 可为NULL, 本库正常运行所需要依赖的其他支持文件

}LIB_INFO, * PLIB_INFO;

#define    FUNCNAME_GET_LIB_INFO        "GetNewInf"     // 此名称必须固定(3.0以前为"GetLibInf")
typedef PLIB_INFO(WINAPI* PFN_GET_LIB_INFO) ();    // 取支持库信息函数
typedef INT(*PFN_ADD_IN_FUNC) ();

// 当LBS_LIB_INFO2被置位时FUNCNAME_GET_LIB_INFO实际返回的是此信息结构
typedef struct
{
    // 支持库授权信息
    LPCSTR m_szHardwareCode;  // // 提供给支持库购买者的本计算机硬件特征码,如无则为NULL或空文本.
    LPCSTR m_szBuyingTips;  // 提供给支持库购买者的提示信息,如无则为NULL或空文本.
    LPCSTR m_szBuyingURL;  // 如果不为NULL或空文本,则提供一个用作接受购买信息的URL,用户选择购买本支持库时系统将自动调用浏览器打开此URL,并将相关用户信息参数附在其后.
    LPCSTR m_szLicenseToUserName;  // 本支持库所授权给的用户名称,如果为NULL/空文本,说明当前使用者尚未得到授权.
}LIB_INFO2, * PLIB_INFO2;

//////////////////////////////////////////////////////////////////////////

#define LIB_BMP_RESOURCE        "LIB_BITMAP"        // 支持库中提供的图像资源的名称
#define LIB_BMP_CX              16                  // 每一图像资源的宽度
#define LIB_BMP_CY              13                  // 每一图像资源的高度
#define LIB_BMP_BKCOLOR         RGB(255, 255, 255)  // 图像资源的底色

////////////////////////////

#define    WU_GET_WND_PTR       (WM_APP + 2)        // 用作在非核心支持库中的窗口组件支持事件反馈



//取时间部分,参数1为只有时间部分的DATE数据
inline void GetTimePart(DATE dt, INT& nHour, INT& nMinute, INT& nSecond)
{
    INT nSecCount = (int)DATE(86400 * fabs(dt) + 0.5);//总秒数
    nHour = nSecCount / 3600;
    nSecCount %= 3600;
    nMinute = nSecCount / 60;
    nSecond = nSecCount % 60;
}

//取日期部分,参数1为只有日期部分的DATE数据
//1899年12月30日0时00分00秒 为0基准,单位为天数的小数值
inline void GetDatePart(DATE dt, INT& nYear, INT& nMonth, INT& nDay)
{
    if (dt == 0 || dt == 1)
    {
        nYear = 1899;
        nMonth = 12;
        nDay = 30 + (int)dt;
        return;
    }
    INT nCountDay;
    INT nStatrYear;
    INT nYearDay;
    INT nStep;
    if (dt > 0)
    {
        nCountDay = (int)dt - 1;//到1900年1月1日为基准
        nStatrYear = 1900;
        nStep = 1;
    }
    else
    {
        nCountDay = (int)fabs(dt - 1);//到1899年12月31日为基准
        nStatrYear = 1899;
        nStep = -1;
    }

    while (1)
    {
        nYearDay = 365;
        if ((nStatrYear % 400 == 0) || (nStatrYear % 100 != 0) && (nStatrYear % 4 == 0))//闰年
            nYearDay++;
        if (nCountDay <= nYearDay)//在这年的范围内
        {
            nYear = nStatrYear;
            if (nStep == -1)
            {
                nCountDay = nYearDay - nCountDay;
            }

            break;
        }
        nCountDay -= nYearDay;
        nStatrYear += nStep;
    }
    DWORD nMonDay[12] = { 31,28,31,30,31,30,31,31,30,31,30,31 };
    if ((nYear % 400 == 0) || (nYear % 100 != 0) && (nYear % 4 == 0))//闰年
        nMonDay[1] = 29; //二月

    for (INT i = 0; i < 12; i++)
    {
        if (nCountDay <= (int)nMonDay[i])
        {
            nMonth = i + 1;
            nDay = nCountDay;
            break;
        }
        nCountDay -= nMonDay[i];

    }

}
// 格式化时间文本, 需要外部传递缓冲区和缓冲区尺寸进来, 返回 sprintf的返回值
inline int DateTimeFormat(LPWSTR strValue, int size, DATE dtDt, BOOL bOnlyDatePart)
{
    DATE dd, dt;
    INT nYear, nMonth, nDay, nHour, nMinute, nSecond;
    dt = modf(dtDt, &dd);


    GetTimePart(dt, nHour, nMinute, nSecond);
    GetDatePart(dd, nYear, nMonth, nDay);


    wchar_t strFormat[128];
    //格式化时间部分
    LPCWSTR strFormatTime = L"";
    if ( nSecond )
        strFormatTime = L"%02d:%02d:%02d";
    else if ( nMinute && nSecond == 0 )
        strFormatTime = L"%02d:%02d";
    else if ( nHour && nMinute == 0 && nSecond == 0 )
        strFormatTime = L"%02d";

#if _MSC_VER < 1201

    wcscpy(strFormat, L"%d-%02d-%02d ");
    if ( bOnlyDatePart == FALSE )
        wcscat(strFormat, strFormatTime);

    return swprintf(strValue, strFormat, nYear, nMonth, nDay, nHour, nMinute, nSecond);
#else

    wcscpy_s(strFormat, 128, L"%d-%02d-%02d ");
    if ( bOnlyDatePart == FALSE )
        wcscat_s(strFormat, 128, strFormatTime);

    return swprintf_s(strValue, size, strFormat, nYear, nMonth, nDay, nHour, nMinute, nSecond);

#endif
}

// 格式化时间文本, 需要外部传递缓冲区和缓冲区尺寸进来, 返回 sprintf的返回值
inline int DateTimeFormat(LPSTR strValue, int size, DATE dtDt, BOOL bOnlyDatePart)
{
    DATE dd, dt;
    INT nYear, nMonth, nDay, nHour, nMinute, nSecond;
    dt = modf(dtDt, &dd);


    GetTimePart(dt, nHour, nMinute, nSecond);
    GetDatePart(dd, nYear, nMonth, nDay);


    char strFormat[128];
    //格式化时间部分
    LPCSTR strFormatTime = "";

    if ( nSecond )
        strFormatTime = "%02d:%02d:%02d";
    else if ( nMinute && nSecond == 0 )
        strFormatTime = "%02d:%02d";
    else if ( nHour && nMinute == 0 && nSecond == 0 )
        strFormatTime = "%02d";

#if _MSC_VER < 1201

    strcpy(strFormat, "%d-%02d-%02d ");
    if (bOnlyDatePart == FALSE)
        strcat(strFormat, strFormatTime);
    return sprintf(strValue, strFormat, nYear, nMonth, nDay, nHour, nMinute, nSecond);
#else

    strcpy_s(strFormat, 128, "%d-%02d-%02d ");
    if (bOnlyDatePart == FALSE)
        strcat_s(strFormat, 128, strFormatTime);

    return sprintf_s(strValue, size, strFormat, nYear, nMonth, nDay, nHour, nMinute, nSecond);

#endif
}
//取指定时间 参数时间值都是正确的
//1899年12月30日0时00分00秒 为0基准,单位为天数的小数值
inline double GetSpecDateTime(INT nYear, INT nMonth, INT nDay, INT nHour, INT nMinute, INT nSecond)
{

    double nSecCount = (double)nHour * 3600;//总秒数
    nSecCount += ((double)nMinute * 60);
    nSecCount += nSecond;
    double dTime = nSecCount / 86400;
    double dDate = 0;
    INT nStep;


    if (nYear == 1899)
    {
        if (nMonth == 12)
        {
            if (nDay >= 30)
            {
                dDate = (double)nDay - 30;
                return (dDate + dTime);
            }

        }

        nStep = -1;
    }
    else if (nYear > 1899)
    {
        nStep = 1;

    }
    else
        nStep = -1;



    INT nStatrYear;
    INT nYearDay;
    INT nCountDay = 0;
    if (nStep == 1)
        nStatrYear = 1900;//到1900年1月1日为基准
    else
        nStatrYear = 1899;



    while (1)
    {

        if (nStatrYear == nYear)
        {
            DWORD nMonDay[12] = { 31,28,31,30,31,30,31,31,30,31,30,31 };
            if ((nYear % 400 == 0) || (nYear % 100 != 0) && (nYear % 4 == 0))//闰年
                nMonDay[1] = 29; //二月
            if (nStep == 1)
            {
                for (INT i = 0; i < 12; i++)
                {
                    if (i == nMonth - 1)
                    {
                        nCountDay += nDay;
                        break;
                    }
                    nCountDay += nMonDay[i];
                }
            }
            else {

                for (INT i = 11; i >= 0; i--)
                {
                    if (i == nMonth - 1)
                    {
                        nCountDay += (nMonDay[i] - nDay);

                        break;
                    }
                    nCountDay += nMonDay[i];
                }

            }
            break;

        }
        nYearDay = 365;
        if ((nStatrYear % 400 == 0) || (nStatrYear % 100 != 0) && (nStatrYear % 4 == 0))//闰年
            nYearDay++;
        nCountDay += nYearDay;
        nStatrYear += nStep;

    }

    if (nStep == 1)
    {
        dDate = (double)nCountDay + 1.0;
        return (dDate + dTime);
    }


    dDate = -((double)nCountDay - 1.0 + dTime);
    return dDate;

}



#endif

//#define SDT_TYPE_OBJECT _SDT_NULL
