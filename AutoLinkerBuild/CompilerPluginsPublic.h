
#ifndef __COMPILER_PLUSINS_PUBLIC_H__
#define __COMPILER_PLUSINS_PUBLIC_H__

typedef enum
{
    // 修改待编译程序的相关信息
    //   dwParam1: COMPILE_PROCESSOR_PRG_INFO* 指向新的程序信息,必须有效且不为NULL.
    PCFN_SET_PRG_INFO = 1,

    // 修改待编译程序内指定数值型常量的值,成功返回1,失败返回0.
    //   dwParam1: const TCHAR* 指向待修改数值常量的名称
    //   dwParam2: INT 所欲修改到的常量值
    PCFN_SET_PRG_NUM_CONST,

    // 修改待编译程序内指定日期时间型常量的值,成功返回1,失败返回0.
    //   dwParam1: const TCHAR* 指向待修改日期时间型常量的名称
    //   dwParam2: DATE* 指向所欲修改到的常量值,必须有效且不为NULL.
    PCFN_SET_PRG_DATE_TIME_CONST,

    // 修改待编译程序内指定逻辑型常量的值,成功返回1,失败返回0.
    //   dwParam1: const TCHAR* 指向待修改逻辑型常量的名称
    //   dwParam2: DWORD 所欲修改到的常量值,为1表示真,为0表示假.
    PCFN_SET_PRG_BOOL_CONST,

    // 修改待编译程序内指定文本型常量的值,成功返回1,失败返回0.
    //   dwParam1: const TCHAR* 指向待修改文本型常量的名称
    //   dwParam2: const TCHAR* 指向所欲修改到的常量文本值
    PCFN_SET_PRG_TEXT_CONST,

    // 修改待编译程序内指定图片或声音类常量的值,成功返回TRUE,失败返回FALSE.
    //   dwParam1: const TCHAR* 指向待修改图片或声音类常量的名称
    //   dwParam2: BYTE* 首先为一个INT记录图片或声音数据的长度(如为空,此值为0),后跟相应长度的图片或声音字节数据.必须有效且不为NULL.
    PCFN_SET_PRG_BIN_CONST,

    // 返回待编译程序内指定数值/逻辑型型常量的值.
    // 对于数值常量,只能返回其整数部分,对于逻辑型常量,为真返回TRUE,否则返回FALSE.
    // 如果未找到所指定名称的常量或者该常量不为数值/逻辑型,则直接返回0.
    //   dwParam1: const TCHAR* 指向欲访问数值/逻辑型常量的名称
    PCFN_GET_PRG_CONST,
}
CALL_BACK_FUNC_NO;

// pCallBackData: COMPILE_PROCESSOR_INFO.m_pCallBackData值
typedef DWORD (WINAPI *FN_COMPILE_CALLBACK) (const void* pCallBackData,
        const DWORD dwFuncNO, const DWORD dwParam1, const DWORD dwParam2);

typedef struct
{
	const TCHAR* m_szPrgName;	// 程序名称
	const TCHAR* m_szExplain;	// 有关本程序的描述和备注

	const TCHAR* m_szAuthor;
	const TCHAR* m_szZipCode;
	const TCHAR* m_szAddress;
	const TCHAR* m_szPhoto;
	const TCHAR* m_szFax;
	const TCHAR* m_szEmail;
	const TCHAR* m_szHomePage;
	const TCHAR* m_szCopyright;

	INT m_nPrgMajorVersion;
	INT m_nPrgMinorVersion;
	INT m_nBuildNO1, m_nBuildNO2;

    const TCHAR* m_szSaveToFileName;  // 编译结果所保存到的文件全路径名,为空表示需要用户手工输入.
}
COMPILE_PROCESSOR_PRG_INFO;

typedef struct
{
    FN_COMPILE_CALLBACK m_fnCallBack;  // 提供给插件用作回调编译器以完成相应功能
    void* m_pCallBackData;  // 编译器提供给插件,插件在调用m_fnCallBack时必须原值返回.

    //---------------------------------------------------

    BOOL m_blCompileReleaseVersion;  // 为真表明为编译发布版本,否则为编译调试版本.
    BOOL m_blCompileEPackage;  // 为真表明为正在编译易包
    const TCHAR* m_szPrgFileName;  // 被编译程序的文件名称(包括路径和后缀),如无(该程序尚未保存)则为空.
    const TCHAR* m_szPurePrgFileName;  // 被编译程序的纯粹文件名称(不包括路径和后缀),如无(该程序尚未保存)则为空.

    //---------------------------------------------------

    DWORD	m_dwSysPrgVersion;  // 系统软件的版本号.
	INT		m_nSysType;			// 系统软件的类型号
	INT		m_nLanguageVer;		// 系统软件的语言版本.
    // 所编译程序的类型（!!!已有宏值绝对不可改变，所有宏值必须大于零）。
    #define AT_WIN_WINDOWS_EXE      0
    #define AT_WIN_CONSOLE_EXE      1
    #define AT_WIN_DLL              2
    #define AT_WIN_ECOM             1000
    #define AT_LINUX_CONSOLE_ELF    10000
    #define AT_LINUX_ECOM           11000
    INT m_nAppType;

    COMPILE_PROCESSOR_PRG_INFO m_infPrg;  // 被编译程序配置信息
}
COMPILE_PROCESSOR_INFO;

// 编译前插件的对外接口函数原型. 返回真表示处理成功,返回假表示处理失败.
//   fnCallBack: 用作与编译器交互
//   pInf: 提供相关编译时信息,必定不为NULL.
typedef BOOL (WINAPI *FN_COMPILE_PROCESSOR) (const COMPILE_PROCESSOR_INFO* pInf);

// 插件输出函数的名称
#define COMPILER_PLUGINS_EXPORT_NAME  "CompileProcessor"

#endif
