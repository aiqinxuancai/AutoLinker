
#ifndef __PUBLIC_IDE_FUNCTIONS_H__
#define __PUBLIC_IDE_FUNCTIONS_H__

// IDE接口功能表
// 注意: 所有功能均通过 NES_RUN_FUNC 通知调用(见lib2.h),如果该功能被处理返回真,未被处理则返回假.

//-------------------------------------------------------------------  相关数据结构和宏定义

// FN_REPLACE_ALL2功能使用
typedef struct
{
    const char* m_szFind;  // 所欲寻找的文本
    const char* m_szReplace;  // 所欲替换的文本
    BOOL m_blCase;  // 是否区分字母大小写
}
REPLACE_ALL2_PARAM;

// FN_ADD_TAB功能使用
typedef struct
{
    HWND m_hWnd;
    HICON m_hIcon;
    LPTSTR m_szCaption;
    LPTSTR m_szToolTip;
}
ADD_TAB_INF;

//------------------------------------  FN_GET_PRG_TEXT和FN_GET_PRG_HELP功能使用

#define VT_DLL_CMD_NAME              51  // DllCmd名称
#define VT_DLL_CMD_RET_TYPE          52  // DllCmd返回值类型
#define VT_DLL_CMD_EXPLAIN           53  // DllCmd备注
#define VT_DLL_CMD_EXPORT            54  // DllCmd公开
#define VT_DLL_CMD_DEFINE_BEGIN      VT_DLL_CMD_NAME
#define VT_DLL_CMD_DEFINE_END        VT_DLL_CMD_EXPORT
#define VT_DLL_LIB_FILE_NAME         VT_DLL_CMD_DEFINE_END + 50   // Dll库文件名
#define VT_DLL_CMD_IN_LIB_NAME       VT_DLL_LIB_FILE_NAME + 51    // DllCmd在库中的名称
#define VT_DLL_CMD_ARG_NAME          VT_DLL_CMD_IN_LIB_NAME + 51  // DllCmd参数名称
#define VT_DLL_CMD_ARG_TYPE          VT_DLL_CMD_IN_LIB_NAME + 52  // DllCmd参数类型
#define VT_DLL_CMD_ARG_POINTER_TYPE  VT_DLL_CMD_IN_LIB_NAME + 53  // DllCmd参数参考类型
#define VT_DLL_CMD_ARG_ARY_TYPE      VT_DLL_CMD_IN_LIB_NAME + 54  // DllCmd参数数组类型
#define VT_DLL_CMD_ARG_EXPLAIN       VT_DLL_CMD_IN_LIB_NAME + 55  // DllCmd参数备注
#define VT_DLL_CMD_ARG_DEFINE_BEGIN  VT_DLL_CMD_ARG_NAME
#define VT_DLL_CMD_ARG_DEFINE_END    VT_DLL_CMD_ARG_EXPLAIN

#define VT_USER_TYPE_NAME            VT_DLL_CMD_ARG_DEFINE_END + 51  // 自定义数据类型名称
#define VT_USER_TYPE_EXPLAIN         VT_DLL_CMD_ARG_DEFINE_END + 52  // 自定义数据类型备注
#define VT_USER_TYPE_EXPORT          VT_DLL_CMD_ARG_DEFINE_END + 53  // 自定义数据类型公开
#define VT_USER_TYPE_DEFINE_BEGIN    VT_USER_TYPE_NAME
#define VT_USER_TYPE_DEFINE_END      VT_USER_TYPE_EXPORT
#define VT_ELEMENT_NAME              VT_USER_TYPE_DEFINE_END + 50  // 数据类型成员名称
#define VT_ELEMENT_TYPE              VT_USER_TYPE_DEFINE_END + 51  // 数据类型成员类型
#define VT_ELEMENT_ARY_TYPE          VT_USER_TYPE_DEFINE_END + 52  // 数据类型成员数组类型
#define VT_ELEMENT_EXPLAIN           VT_USER_TYPE_DEFINE_END + 53  // 数据类型成员备注
#define VT_ELEMENT_POINTER_TYPE      VT_USER_TYPE_DEFINE_END + 54  // 数据类型成员参考类型
#define VT_ELEMENT_DEFINE_BEGIN      VT_ELEMENT_NAME
#define VT_ELEMENT_DEFINE_END        VT_ELEMENT_POINTER_TYPE

#define VT_GLOBAL_VAR_NAME           VT_ELEMENT_DEFINE_END + 50  // 全局变量名称
#define VT_GLOBAL_VAR_TYPE           VT_ELEMENT_DEFINE_END + 51  // 全局变量类型
#define VT_GLOBAL_VAR_ARY_TYPE       VT_ELEMENT_DEFINE_END + 52  // 全局变量数组类型
#define VT_GLOBAL_VAR_EXPLAIN        VT_ELEMENT_DEFINE_END + 53  // 全局变量备注
#define VT_GLOBAL_VAR_EXPORT         VT_ELEMENT_DEFINE_END + 54  // 全局变量公开
#define VT_GLOBAL_VAR_DEFINE_BEGIN   VT_GLOBAL_VAR_NAME
#define VT_GLOBAL_VAR_DEFINE_END     VT_GLOBAL_VAR_EXPORT

#define VT_CONST_RES_NAME            VT_GLOBAL_VAR_DEFINE_END + 50  // 常量资源名称
#define VT_CONST_RES_VAULE           VT_GLOBAL_VAR_DEFINE_END + 51  // 常量资源值
#define VT_CONST_RES_EXPLAIN         VT_GLOBAL_VAR_DEFINE_END + 52  // 常量资源备注
#define VT_CONST_RES_EXPORT          VT_GLOBAL_VAR_DEFINE_END + 53  // 常量资源公开
#define VT_CONST_RES_DEFINE_BEGIN    VT_CONST_RES_NAME
#define VT_CONST_RES_DEFINE_END      VT_CONST_RES_EXPORT

#define VT_PIC_RES_NAME              VT_CONST_RES_DEFINE_END + 50  // 图片资源名称
#define VT_PIC_RES_VAULE             VT_CONST_RES_DEFINE_END + 51  // 图片资源内容
#define VT_PIC_RES_EXPLAIN           VT_CONST_RES_DEFINE_END + 52  // 图片资源备注
#define VT_PIC_RES_EXPORT            VT_CONST_RES_DEFINE_END + 53  // 图片资源公开
#define VT_PIC_RES_DEFINE_BEGIN      VT_PIC_RES_NAME
#define VT_PIC_RES_DEFINE_END        VT_PIC_RES_EXPORT

#define VT_SOUND_RES_NAME            VT_PIC_RES_DEFINE_END + 50  // 声音资源名称
#define VT_SOUND_RES_VAULE           VT_PIC_RES_DEFINE_END + 51  // 声音资源内容
#define VT_SOUND_RES_EXPLAIN         VT_PIC_RES_DEFINE_END + 52  // 声音资源备注
#define VT_SOUND_RES_EXPORT          VT_PIC_RES_DEFINE_END + 53  // 声音资源公开
#define VT_SOUND_RES_DEFINE_BEGIN    VT_SOUND_RES_NAME
#define VT_SOUND_RES_DEFINE_END      VT_SOUND_RES_EXPORT

#define VT_MOD_NAME                  VT_SOUND_RES_DEFINE_END + 50  // 程序集名称
#define VT_MOD_EXPLAIN               VT_SOUND_RES_DEFINE_END + 51  // 程序集备注
#define VT_MOD_BASE_CLASS            VT_SOUND_RES_DEFINE_END + 52  // 基类
#define VT_MOD_EXPORT                VT_SOUND_RES_DEFINE_END + 53  // 输出
#define VT_MOD_DEFINE_BEGIN          VT_MOD_NAME
#define VT_MOD_DEFINE_END            VT_MOD_EXPORT

#define VT_MOD_VAR_NAME              VT_MOD_DEFINE_END + 49  // 程序集变量名称
#define VT_MOD_VAR_TYPE              VT_MOD_DEFINE_END + 50  // 程序集变量类型
#define VT_MOD_VAR_ARY_TYPE          VT_MOD_DEFINE_END + 51  // 程序集变量数组类型
#define VT_MOD_VAR_EXPLAIN           VT_MOD_DEFINE_END + 52  // 程序集变量备注
#define VT_MOD_VAR_DEFINE_BEGIN      VT_MOD_VAR_NAME
#define VT_MOD_VAR_DEFINE_END        VT_MOD_VAR_EXPLAIN

#define VT_SUB_NAME                  VT_MOD_VAR_DEFINE_END + 51  // 子程序名称
#define VT_SUB_RET_TYPE              VT_MOD_VAR_DEFINE_END + 52  // 子程序返回值类型
#define VT_SUB_EPK_NAME              VT_MOD_VAR_DEFINE_END + 53  // 子程序所处易包名称
#define VT_SUB_EXPLAIN               VT_MOD_VAR_DEFINE_END + 54  // 子程序备注
#define VT_SUB_EXPORT                VT_MOD_VAR_DEFINE_END + 55  // 此子程序是否被输出
#define VT_SUB_DEFINE_BEGIN          VT_SUB_NAME
#define VT_SUB_DEFINE_END            VT_SUB_EXPORT

#define VT_SUB_ARG_NAME              VT_SUB_DEFINE_END + 50  // 子程序参数名称
#define VT_SUB_ARG_TYPE              VT_SUB_DEFINE_END + 51  // 子程序参数类型
#define VT_SUB_ARG_POINTER_TYPE      VT_SUB_DEFINE_END + 52  // 子程序参数参考类型
#define VT_SUB_ARG_NULL_TYPE         VT_SUB_DEFINE_END + 53  // 子程序参数NULL类型
#define VT_SUB_ARG_ARY_TYPE          VT_SUB_DEFINE_END + 54  // 子程序参数数组类型
#define VT_SUB_ARG_EXPLAIN           VT_SUB_DEFINE_END + 55  // 子程序参数备注
#define VT_SUB_ARG_DEFINE_BEGIN      VT_SUB_ARG_NAME
#define VT_SUB_ARG_DEFINE_END        VT_SUB_ARG_EXPLAIN

#define VT_SUB_VAR_NAME              VT_SUB_ARG_DEFINE_END + 51  // 子程序局部变量名称
#define VT_SUB_VAR_TYPE              VT_SUB_ARG_DEFINE_END + 52  // 子程序局部变量类型
#define VT_SUB_VAR_STATIC_TYPE       VT_SUB_ARG_DEFINE_END + 53  // 子程序局部变量静态类型
#define VT_SUB_VAR_ARY_TYPE          VT_SUB_ARG_DEFINE_END + 54  // 子程序局部变量数组类型
#define VT_SUB_VAR_EXPLAIN           VT_SUB_ARG_DEFINE_END + 55  // 子程序局部变量备注
#define VT_SUB_VAR_DEFINE_BEGIN      VT_SUB_VAR_NAME
#define VT_SUB_VAR_DEFINE_END        VT_SUB_VAR_EXPLAIN
#define VT_SUB_PRG_ITEM              VT_SUB_VAR_DEFINE_END + 51  // 子程序中的命令或参数行.

typedef struct
{
    int m_nRowIndex;   // 指定欲获取文本的行索引,为-1表示使用当前光标所处行.
    int m_nColIndex;   // 指定欲获取文本的列索引,为-1表示使用当前光标所处列.
    char* m_pBuf;      // 提供一个缓冲区地址,为NULL时功能将在m_nBufSize中返回所需缓冲区的尺寸.
    int m_nBufSize;    // 如果m_pBuf不为NULL,用户需要在此处提供为m_pBuf缓冲区准备的尺寸,否则仅用作返回所需缓冲区的尺寸.
    int m_nType;       // 在其中返回当前位置类型类型,为以上的"VT_xxx"宏值.
    BOOL m_blIsTitle;  // 在其中返回当前位置是否为表格标题栏
}
GET_PRG_TEXT_PARAM;

//-------------------------------------------------------------------
// 功能类别(最高位记录该功能是否不需要记录,最高字节的后7位记录大类,次高字节记录子类):

#define FNT_SYS  0x00000000
#define FNT_MOVE  0x01000000
    #define FNST_MOVE_CARET    (FNT_MOVE | 0x010000)
    #define FNST_SCROLL_WND    (FNT_MOVE | 0x020000)
#define FNT_EDIT  0x02000000
    #define FNST_BLOCK_OPER    (FNT_EDIT | 0x010000)
    #define FNST_MODIFY        (FNT_EDIT | 0x020000)
    #define FNST_OTHER_MODIFY  (FNT_EDIT | 0x030000)
#define FNT_DOCUMENT  0x03000000
    #define FNST_FILE          (FNT_DOCUMENT | 0x010000)
#define FNT_OTHER  0x04000000
    #define FNST_STATE_OPER    (FNT_OTHER | 0x010000)
    #define FNST_OTHER         (FNT_OTHER | 0x020000)
#define FNT_EAPP  0x05000000
    #define FNST_DEBUG         (FNT_EAPP | 0x010000)
    #define FNST_RUN           (FNT_EAPP | 0x020000)
    #define FNST_ADDIN         (FNT_EAPP | 0x030000)

#define GetFuncParentType(dwFuncNO)  ((dwFuncNO) & 0x7f000000)
#define GetFuncType(dwFuncNO)  ((dwFuncNO) & 0x7fff0000)

//-------------------------------------------------------------------  功能表

// 空功能:
#define FN_NULL  0

//------------------------------------  光标移动功能:

#define FN_MOVE_UP                  (FNST_MOVE_CARET | 1)   // 光标上移一行
#define FN_MOVE_DOWN                (FNST_MOVE_CARET | 2)   // 光标下移一行
#define FN_MOVE_TOP                 (FNST_MOVE_CARET | 3)   // 光标移到顶部
#define FN_MOVE_BOTTOM              (FNST_MOVE_CARET | 4)   // 光标移到底部
#define FN_MOVE_FIRST_COL           (FNST_MOVE_CARET | 5)   // 光标移到表格首列
#define FN_MOVE_LAST_COL            (FNST_MOVE_CARET | 6)   // 光标移到表格尾列
#define FN_MOVE_PREV_COL            (FNST_MOVE_CARET | 7)   // 光标移到表格前一列
#define FN_MOVE_NEXT_COL            (FNST_MOVE_CARET | 8)   // 光标移到表格后一列
#define FN_MOVE_SCROLL_UP           (FNST_MOVE_CARET | 9)   // 窗口向上滚动一行
#define FN_MOVE_SCROLL_DOWN         (FNST_MOVE_CARET | 10)  // 窗口向下滚动一行
#define FN_MOVE_PAGEUP              (FNST_MOVE_CARET | 11)  // 窗口向上翻一页
#define FN_MOVE_PAGEDOWN            (FNST_MOVE_CARET | 12)  // 窗口向下翻一页
#define FN_MOVE_PAGE_HOME           (FNST_MOVE_CARET | 13)  // 光标到当前页首部
#define FN_MOVE_PAGE_END            (FNST_MOVE_CARET | 14)  // 光标到当前页底部
#define FN_MOVE_BLK_SEL_UP          (FNST_MOVE_CARET | 27)  // 同上,不过附带块选择.
#define FN_MOVE_BLK_SEL_DOWN        (FNST_MOVE_CARET | 28)
#define FN_MOVE_BLK_SEL_TOP         (FNST_MOVE_CARET | 29)
#define FN_MOVE_BLK_SEL_BOTTOM      (FNST_MOVE_CARET | 30)
#define FN_MOVE_BLK_SEL_PAGEUP      (FNST_MOVE_CARET | 31)
#define FN_MOVE_BLK_SEL_PAGEDOWN    (FNST_MOVE_CARET | 32)
#define FN_MOVE_BLK_SEL_PAGE_HOME   (FNST_MOVE_CARET | 33)
#define FN_MOVE_BLK_SEL_PAGE_END    (FNST_MOVE_CARET | 34)
#define FN_MOVE_BLK_SEL_ALL         (FNST_MOVE_CARET | 35)  // 全部选择
#define FN_MOVE_TO_PARENT_CMD       (FNST_MOVE_CARET | 36)  // 移动到父命令
#define FN_MOVE_UP_FAR_ITEM         (FNST_MOVE_CARET | 38)  // 到配套的前一个远端语句
#define FN_MOVE_DOWN_FAR_ITEM       (FNST_MOVE_CARET | 39)  // 到配套的后一个远端语句
#define FN_MOVE_OPEN_ARG            (FNST_MOVE_CARET | 40)  // 打开语句参数表
#define FN_MOVE_CLOSE_ARG           (FNST_MOVE_CARET | 41)  // 关闭语句参数表
#define FN_MOVE_OPEN_SPEC_ROW_ARG   (FNST_MOVE_CARET | 42)  // 打开指定行的语句参数表. 功能参数1: 所指定行索引位置
#define FN_MOVE_CLOSE_SPEC_ROW_ARG  (FNST_MOVE_CARET | 43)  // 关闭指定行的语句参数表. 功能参数1: 所指定行索引位置
#define FN_MOVE_OPEN_ALL_ARG        (FNST_MOVE_CARET | 44)  // 打开所有参数表
#define FN_MOVE_PREV_UNIT           (FNST_MOVE_CARET | 45)  // 到前一单元
#define FN_MOVE_NEXT_UNIT           (FNST_MOVE_CARET | 46)  // 到后一单元
#define FN_MOVE_SPEC_SUB            (FNST_MOVE_CARET | 47)  // 跳到当前语句行所调用的子程序或DLL命令处
#define FN_MOVE_OPEN_SPEC_SUB       (FNST_MOVE_CARET | 48)  // 打开当前光标所处子程序
#define FN_MOVE_CLOSE_SPEC_SUB      (FNST_MOVE_CARET | 49)  // 关闭当前光标所处子程序
#define FN_MOVE_OPEN_SPEC_SUB_GRP   (FNST_MOVE_CARET | 50)  // 打开当前块选择区域所涉及到的所有子程序
#define FN_MOVE_CLOSE_SPEC_SUB_GRP  (FNST_MOVE_CARET | 51)  // 关闭当前块选择区域所涉及到的所有子程序
#define FN_MOVE_EDIT_CARET_PREV     (FNST_MOVE_CARET | 54)  // 光标左移一个字符
#define FN_MOVE_EDIT_CARET_NEXT     (FNST_MOVE_CARET | 55)  // 光标右移一个字符
#define FN_MOVE_EDIT_CARET_TO_BEGIN (FNST_MOVE_CARET | 56)  // 光标到行首
#define FN_MOVE_EDIT_CARET_TO_END   (FNST_MOVE_CARET | 57)  // 光标到行尾
#define FN_SKIP_LEFT_WORD           (FNST_MOVE_CARET | 58)  // 左移一个单词
#define FN_SKIP_RIGHT_WORD          (FNST_MOVE_CARET | 59)  // 右移一个单词
#define FN_MOVE_BLK_SEL_LEFT        (FNST_MOVE_CARET | 60)  // 同上,不过附带块选择.
#define FN_MOVE_BLK_SEL_RIGHT       (FNST_MOVE_CARET | 61)
#define FN_MOVE_BLK_SEL_HOME        (FNST_MOVE_CARET | 62)
#define FN_MOVE_BLK_SEL_END         (FNST_MOVE_CARET | 63)
#define FN_MOVE_BLK_SEL_LEFT_WORD   (FNST_MOVE_CARET | 64)
#define FN_MOVE_BLK_SEL_RIGHT_WORD  (FNST_MOVE_CARET | 65)
#define FN_MOVE_CARET               (FNST_MOVE_CARET | 66)  // 将光标移动到指定位置. 功能参数1: 行索引位置; 功能参数2: 列索引位置.
#define FN_MOVE_BACK_SUB            (FNST_MOVE_CARET | 67)  // 跳回先前调用FN_MOVE_SPEC_SUB跳转前所处的子程序

//------------------------------------  窗口滚动功能:

#define FN_SCROLL_LEFT              (FNST_SCROLL_WND | 1)
#define FN_SCROLL_RIGHT             (FNST_SCROLL_WND | 2)
#define FN_SCROLL_UP                (FNST_SCROLL_WND | 3)
#define FN_SCROLL_DOWN              (FNST_SCROLL_WND | 4)
#define FN_SCROLL_PAGELEFT          (FNST_SCROLL_WND | 5)
#define FN_SCROLL_PAGERIGHT         (FNST_SCROLL_WND | 6)
#define FN_SCROLL_PAGEUP            (FNST_SCROLL_WND | 7)
#define FN_SCROLL_PAGEDOWN          (FNST_SCROLL_WND | 8)
#define FN_SCROLL_SPEC_HORZ_POS     (FNST_SCROLL_WND | 9)
#define FN_SCROLL_SPEC_VERT_POS     (FNST_SCROLL_WND | 10)

//------------------------------------  块操作功能:

#define FN_BLK_ADD_DEF              (FNST_BLOCK_OPER | 1)  // 增加所指定块选择区域. 功能参数1: 顶部行索引位置; 功能参数2: 底部行索引位置.
#define FN_BLK_REMOVE_DEF           (FNST_BLOCK_OPER | 2)  // 删除所指定块选择区域. 功能参数1: 顶部行索引位置; 功能参数2: 底部行索引位置.
#define FN_BLK_CLEAR_ALL_DEF        (FNST_BLOCK_OPER | 3)  // 删除所有块选择区域

//------------------------------------  其他编辑功能(不作修改):

#define FN_EDIT_COPY                (FNST_OTHER_MODIFY | 1)  // 复制到剪贴板

//------------------------------------  编辑功能:

#define FN_INSERT_NEW               (FNST_MODIFY | 1)   // 插入一个新当前单元
#define FN_INSERT_NEW_AT_NEXT       (FNST_MODIFY | 2)   // 在后方插入一个新当前单元
#define FN_REMOVE                   (FNST_MODIFY | 3)   // 删除光标当前所处字符或者当前所选择块
#define FN_EDIT_CUT                 (FNST_MODIFY | 4)   // 剪切到剪贴板
#define FN_EDIT_PASTE               (FNST_MODIFY | 5)   // 将剪贴板内容粘贴到当前光标所处位置
#define FN_INSERT_NEW_SUB           (FNST_MODIFY | 9)   // 插入一个新子程序
#define FN_INSERT_NEW_DATA_TYPE     (FNST_MODIFY | 10)  // 插入一个新用户自定义数据类型
#define FN_INSERT_NEW_GLOBAL_VAR    (FNST_MODIFY | 11)  // 插入一个新全局变量
#define FN_INSERT_NEW_DLL_CMD       (FNST_MODIFY | 12)  // 插入一个新DLL命令
#define FN_INSERT_NEW_MOD           (FNST_MODIFY | 13)  // 插入一个新程序集
#define FN_INSERT_NEW_ARG           (FNST_MODIFY | 14)  // 插入一个新子程序参数
#define FN_INSERT_NEW_LOCAL_VAR     (FNST_MODIFY | 15)  // 插入一个新局部变量
#define FN_INSERT_NEW_CONST_RES     (FNST_MODIFY | 16)  // 插入一个新常量
#define FN_INSERT_NEW_PIC_RES       (FNST_MODIFY | 17)  // 插入一个新图片资源
#define FN_INSERT_NEW_SOUND_RES     (FNST_MODIFY | 18)  // 插入一个新声音资源
#define FN_UNDO                     (FNST_MODIFY | 50)  // 撤销
#define FN_REDO                     (FNST_MODIFY | 51)  // 重做
#define FN_FORCE_PROCESS            (FNST_MODIFY | 54)  // 强制预编译当前光标所处语句
#define FN_INSERT_NEW_WIN_FORM      (FNST_MODIFY | 55)  // 插入一个新窗体
#define FN_REMOVE_SPEC_UNIT         (FNST_MODIFY | 62)  // 删除窗体上当前所有被选择的组件
#define FN_MOVE_LEFT_ONE            (FNST_MODIFY | 63)  // 将窗体上当前所有被选择的组件左移一个像素. 下同.
#define FN_MOVE_RIGHT_ONE           (FNST_MODIFY | 64)
#define FN_MOVE_UP_ONE              (FNST_MODIFY | 65)
#define FN_MOVE_DOWN_ONE            (FNST_MODIFY | 66)
#define FN_MOVE_LEFT_ONE_GRID       (FNST_MODIFY | 67)
#define FN_MOVE_RIGHT_ONE_GRID      (FNST_MODIFY | 68)
#define FN_MOVE_UP_ONE_GRID         (FNST_MODIFY | 69)
#define FN_MOVE_DOWN_ONE_GRID       (FNST_MODIFY | 70)
#define FN_RESIZE_LEFT_ONE          (FNST_MODIFY | 71)  // 将窗体上当前所有被选择的组件向左扩展一个像素. 下同.
#define FN_RESIZE_RIGHT_ONE         (FNST_MODIFY | 72)
#define FN_RESIZE_UP_ONE            (FNST_MODIFY | 73)
#define FN_RESIZE_DOWN_ONE          (FNST_MODIFY | 74)
#define FN_RESIZE_LEFT_ONE_GRID     (FNST_MODIFY | 75)
#define FN_RESIZE_RIGHT_ONE_GRID    (FNST_MODIFY | 76)
#define FN_RESIZE_UP_ONE_GRID       (FNST_MODIFY | 77)
#define FN_RESIZE_DOWN_ONE_GRID     (FNST_MODIFY | 78)
#define FN_HORZ_CENTER              (FNST_MODIFY | 79)  // 将窗体上当前所有被选择的组件水平居中. 下同
#define FN_VERT_CENTER              (FNST_MODIFY | 80)
#define FN_UNIT_GO_TOP              (FNST_MODIFY | 82)  // 将窗体上当前所有被选择的组件移动到重叠组件的顶层
#define FN_UNIT_GO_BOTTOM           (FNST_MODIFY | 83)  // 将窗体上当前所有被选择的组件移动到重叠组件的底层
#define FN_MENU_MODI                (FNST_MODIFY | 85)  // 启动当前窗体的菜单设计器
#define FN_INPUT_PRG                (FNST_MODIFY | 88)  // 询问程序文件名,将该程序文件合并进当前程序.
#define FN_APP_SETUP                (FNST_MODIFY | 89)  // 调用程序设置对话框
#define FN_ADD_FIELD_CONST          (FNST_MODIFY | 90)  // 将EDB数据库的字段名转换为常量
#define FN_LEFT_ALIGN               (FNST_MODIFY | 93)  // 将窗体上当前所有被选择的组件左对齐. 下同
#define FN_TOP_ALIGN                (FNST_MODIFY | 94)
#define FN_RIGHT_ALIGN              (FNST_MODIFY | 95)
#define FN_BOTTOM_ALIGN             (FNST_MODIFY | 96)
#define FN_HORZ_MIDDLE              (FNST_MODIFY | 97)
#define FN_VERT_MIDDLE              (FNST_MODIFY | 98)
#define FN_HORZ_AVG                 (FNST_MODIFY | 99)
#define FN_VERT_AVG                 (FNST_MODIFY | 100)
#define FN_EQU_WIDTH                (FNST_MODIFY | 101)
#define FN_EQU_HEIGHT               (FNST_MODIFY | 102)
#define FN_EQU_WIDTHHEIGHT          (FNST_MODIFY | 103)
#define FN_INSERT_NEW_CLASS_MOD     (FNST_MODIFY | 106)  // 插入新的类程序集
#define FN_INSERT_TEXT              (FNST_MODIFY | 109)  // 插入文本到程序编辑窗口的当前光标所处字符位置. 功能参数1(CHAR*): 指向欲插入文本的指针; 功能参数2(BOOL): 该文本是否为用户从键盘输入
#define FN_PRE_COMPILE              (FNST_MODIFY | 112)  // 预编译当前程序. 功能参数1(BOOL*): 指向一个BOOL变量的指针,用作返回是否预编译成功.
#define FN_SET_AND_COMPILE_PRG_ITEM_TEXT  (FNST_MODIFY | 113)  // 设置当前光标所处行列位置的整体内容. 功能参数1(CHAR*): 指向欲设置文本; 功能参数2(BOOL): 是否立即进行预编译.
#define FN_REPLACE_ALL              (FNST_MODIFY | 114)
#define FN_REMARK_STATMENT          (FNST_MODIFY | 115)  // 注释当前语句
#define FN_UNREMARK_STATMENT        (FNST_MODIFY | 116)  // 取消当前语句的注释
#define FN_INSERT_FILE_NAME         (FNST_MODIFY | 117)  // 询问并插入文件名文本到当前光标处
#define FN_LOCK_UNIT                (FNST_MODIFY | 118)  // 加锁窗体上当前被选择组件
#define FN_UNLOCK_UNIT              (FNST_MODIFY | 119)  // 解锁窗体上当前被选择组件
#define FN_REMOVE_SPEC_ECOM         (FNST_MODIFY | 122)  // 删除指定索引位置处被引用的模块. 功能参数1: 所指定模块的索引位置
#define FN_INSERT_NEW_LONGTEXT      (FNST_MODIFY | 126)  // 插入新的长文本资源(仅在常量编辑窗口有效)
#define FN_HORZ_UNIT_MIDDLE         (FNST_MODIFY | 127)  // 将窗体中的所有被选择组件水平居中. 下同
#define FN_VERT_UNIT_MIDDLE         (FNST_MODIFY | 128)
#define FN_INPUT_PRG2               (FNST_MODIFY | 133)  // 将所指定程序文件合并进当前程序. 功能参数1: 指向欲合并进来的程序文件路径名称
#define FN_REPLACE_ALL2             (FNST_MODIFY | 134)  // 替换当前编辑窗口内的所有指定文本. 功能参数1(REPLACE_ALL2_PARAM*): 指定替换参数
#define FN_ADD_NEW_ECOM2            (FNST_MODIFY | 135)  // 加入所指定文件名的易模块. 功能参数1(char*): 所指定易模块文件名; 功能参数2(BOOL*): 在该指针内返回是否加入成功

//------------------------------------  文件操作:

#define FN_NEW_FILE                 (FNST_FILE | 1)  // 新建文件
#define FN_OPEN_FILE                (FNST_FILE | 2)  // 打开文件
#define FN_CLOSE_FILE               (FNST_FILE | 3)  // 关闭文件
#define FN_SAVE_FILE                (FNST_FILE | 4)  // 保存文件
#define FN_SAVE_AS_FILE             (FNST_FILE | 5)  // 文件另存为
#define _FN_AUTO_BAK_FILE           (FNST_FILE | 7)  // 自动备份
#define FN_OPEN_FILE2               (FNST_FILE | 8)  // 打开所指定名称文件. 功能参数1(char*): 所欲打开的文件名

//------------------------------------  其它:

#define FN_ABOUT                    (FNST_OTHER | 1)   // 关于
#define FN_SWITCH_WORK_BAR          (FNST_OTHER | 2)   // 开关工作工具条
#define FN_SWITCH_OUTPUT_BAR        (FNST_OTHER | 3)   // 开关输出工具条
#define FN_SWITCH_CMB_BAR           (FNST_OTHER | 4)   // 开关组合工具条
#define FN_SWITCH_UNIT_BAR          (FNST_OTHER | 5)   // 开关窗口组件工具条
#define FN_VIEW_DATA_TYPE_TAB       (FNST_OTHER | 8)   // 打开用户自定义数据类型编辑窗口
#define FN_VIEW_GLOBAL_VAR_TAB      (FNST_OTHER | 9)   // 打开全局变量编辑窗口
#define FN_VIEW_DLLCMD_TAB          (FNST_OTHER | 10)  // 打开DLL命令编辑窗口
#define FN_VIEW_CONST_TAB           (FNST_OTHER | 11)  // 打开常量编辑窗口
#define FN_VIEW_PIC_TAB             (FNST_OTHER | 12)  // 打开图片资源编辑窗口
#define FN_VIEW_SOUND_TAB           (FNST_OTHER | 13)  // 打开声音资源编辑窗口
#define FN_TEST_WIN                 (FNST_OTHER | 14)  // 所设计窗口预览
#define FN_SET_READONLY             (FNST_OTHER | 15)  // 翻转当前程序的只读状态
#define FN_SYS_SETUP                (FNST_OTHER | 16)  // 打开系统设置对话框
#define FN_FIND_BEGIN               (FNST_OTHER | 17)  // 打开寻找对话框
#define FN_FIND_NEXT                (FNST_OTHER | 18)  // 寻找下一个
#define FN_FIND_PREV                (FNST_OTHER | 20)  // 寻找上一个
#define FN_COPY_HELP_TO_CLIP        (FNST_OTHER | 21)  // 将支持库工作夹中当前所选择命令的帮助文本拷贝到剪贴板
#define FN_WRITE_HELP_TO_FILE       (FNST_OTHER | 22)  // 将支持库工作夹中当前所选择命令的帮助文本写到文件
#define FN_HELP                     (FNST_OTHER | 23)  // 打开帮助
#define FN_DBSTRUCT_MANGER          (FNST_OTHER | 25)  // 打开数据库结构管理器
#define FN_DBREC_MANGER             (FNST_OTHER | 26)  // 打开数据库记录管理器
#define FN_SHOW_PROPERTY            (FNST_OTHER | 27)  // 显示属性工作夹
#define FN_DB_CNV                   (FNST_OTHER | 28)  // 打开数据库转换工具
#define FN_RPT_EDITOR               (FNST_OTHER | 29)  // 打开报表编辑器工具
#define FN_LIB_SETUP                (FNST_OTHER | 31)  // 打开支持库设置对话框
#define FN_LIB_MANGER               (FNST_OTHER | 32)  // 打开支持库管理器工具
#define FN_OCX_PACKAGER             (FNST_OTHER | 34)  // 打开OCX包装工具
#define FN_REFRUSH_LIB              (FNST_OTHER | 35)  // 重新载入所有支持库
#define FN_DB_LANG_CONV             (FNST_OTHER | 39)  // 打开数据库语言转换工具
#define FN_DAY_HELP                 (FNST_OTHER | 40)  // 打开每日帮助工具
#define FN_ESN                      (FNST_OTHER | 41)  // 打开易语言知识库
#define FN_WIZARD_MANG              (FNST_OTHER | 42)  // 询问并执行所指定文件名向导
#define FN_SEARCH_REPLACE           (FNST_OTHER | 43)  // 打开寻找替换对话框
#define FN_FIND_ALL                 (FNST_OTHER | 44)  // 打开寻找所有对话框
#define FN_USER_RPT                 (FNST_OTHER | 45)
#define FN_TUTORIAL                 (FNST_OTHER | 47)  // 打开易语言多媒体教程
#define FN_SORT_LIB_TREE            (FNST_OTHER | 48)  // 开关支持库工作夹的排序状态
#define FN_GOTO_SPEC_UNIT_DATA_TYPE (FNST_OTHER | 49)  // 在支持库命令夹中跳转到当前窗体设计器中所选择组件所对应的数据类型项目上
#define FN_RELINK                   (FNST_OTHER | 50)  // 重新进行名称关联操作
#define FN_GOTO_LAST_MODI_PLACE     (FNST_OTHER | 51)  // 跳转到最近一次进行了修改操作的文档位置
#define FN_SET_BOOKMARK             (FNST_OTHER | 52)  // 开关当前光标所在行的书签标志
#define FN_GOTO_BOOKMARK            (FNST_OTHER | 53)  // 打开书签跳转对话框
#define FN_PREV_BOOKMARK            (FNST_OTHER | 54)  // 到上一个书签
#define FN_NEXT_BOOKMARK            (FNST_OTHER | 55)  // 到下一个书签
#define FN_EVA_DESIGNER             (FNST_OTHER | 56)
#define FN_EXTEND_ALL_SUB           (FNST_OTHER | 57)  // 打开所有被收缩或隐藏的子程序
#define FN_VREPORT                  (FNST_OTHER | 58)

//------------------------------------  调试:

#define FN_STEP_INTO                (FNST_DEBUG | 1)   // 单步跟踪进入
#define FN_STEP                     (FNST_DEBUG | 2)   // 单步跟踪
#define FN_STEP_OUT                 (FNST_DEBUG | 3)   // 跟踪跳出
#define FN_RUN_TO_CURSOR            (FNST_DEBUG | 4)   // 执行到光标处
#define FN_VIEW_VAR                 (FNST_DEBUG | 5)   // 显示变量修改对话框
#define FN_SET_BREAK_POINTER        (FNST_DEBUG | 6)   // 设置断点
#define FN_CLEAR_ALL_BREAK_POINTER  (FNST_DEBUG | 7)   // 清除所有断点
#define FN_SHOW_NEXT_STATMENT       (FNST_DEBUG | 8)   // 跳到当前所正在执行的语句
#define FN_ADV_BREAKPOINT           (FNST_DEBUG | 10)  // 设置条件断点

//------------------------------------  运行:

#define FN_COMPILE                             (FNST_RUN | 1)   // 编译
#define FN_COMPILE_AND_RUN                     (FNST_RUN | 2)   // 编译执行
#define FN_END_RUN                             (FNST_RUN | 3)   // 结束执行
#define FN_PUBLISH                             (FNST_RUN | 5)   // 编译发布
#define FN_COMPILE_STANDARD                    (FNST_RUN | 6)   // 独立编译
#define FN_COMPILE_STATIC                      (FNST_RUN | 7)   // 静态编译
#define FN_COMPILE_EPACKAGE                    (FNST_RUN | 8)   // 编译易包
#define FN_COMPILE_WINDOWS_EXE_AND_RUN         (FNST_RUN | 9)   // 编译为Windows程序并执行
#define FN_COMPILE_WINDOWS_CONOLE_EXE_AND_RUN  (FNST_RUN | 10)  // 编译为Windows控制台程序并执行
#define FN_COMPILE_WINDOWS_DLL_AND_RUN         (FNST_RUN | 11)  // 编译为Windows动态链接库并执行
#define FN_COMPILE_WINDOWS_ECOM_AND_RUN        (FNST_RUN | 12)  // 编译为Windows易模块并执行
#define FN_COMPILE_WINDOWS_EXE                 (FNST_RUN | 13)  // 编译为Windows程序
#define FN_COMPILE_WINDOWS_CONOLE_EXE          (FNST_RUN | 14)  // 编译为Windows控制台程序
#define FN_COMPILE_WINDOWS_DLL                 (FNST_RUN | 15)  // 编译为Windows动态链接库
#define FN_COMPILE_WINDOWS_ECOM                (FNST_RUN | 16)  // 编译为Windows易模块
#define FN_STATIC_COMPILE_WINDOWS_EXE          (FNST_RUN | 17)  // 静态编译为Windows程序
#define FN_STATIC_COMPILE_WINDOWS_CONOLE_EXE   (FNST_RUN | 18)  // 静态编译为Windows控制台程序
#define FN_STATIC_COMPILE_WINDOWS_DLL          (FNST_RUN | 19)  // 静态编译为Windows动态链接库

//------------------------------------  辅助:

typedef ADD_TAB_INF* PADD_TAB_INF;
#define FN_ADD_TAB  (FNST_ADDIN | 1)  // 在输出工具条中增加一个指定的工作夹. 功能参数1(ADD_TAB_INF*): 工作夹信息

/* 在 *(int*)功能参数1 中返回当前活动编辑窗口的类型,为以下值之一:
    0: 无
    1: 程序集
    2: 用户自定义数据类型
    3: 全局变量
    4: DLL命令
    5: 窗体设计
    6: 常量资源
    7: 图片资源
    8: 声音资源
功能参数1(int*): 提供填写返回数据的整数变量地址  */
#define FN_GET_ACTIVE_WND_TYPE  (FNST_ADDIN | 2)

// 导入易模块
//   功能参数1(char*): 欲导入的易模块文件名
//   功能参数2(BOOL*): 在其中返回是否导入成功
#define FN_INPUT_ECOM  (FNST_ADDIN | 3)

// 检查指定的功能目前是否可用
//   功能参数1: 欲检查的功能号
//   功能参数2(BOOL*): 在其中返回所检查功能是否可用
#define FN_IS_FUNC_ENABLED  (FNST_ADDIN | 4)

// 检查当前系统剪贴板中是否有易程序数据
//   功能参数1(int*): 如有则在其中返回大于0的数据长度,否则返回0.
#define FN_CLIP_GET_EPRG_DATA_SIZE  (FNST_ADDIN | 5)

// 从当前系统剪贴板中读出易程序数据
//   功能参数1(BYTE*): 指向数据接收缓冲区的指针,其尺寸必须等于使用FN_CLIP_GET_EPRG_DATA_SIZE所获得的尺寸.
//   功能参数2(BOOL*): 在其中返回是否读出成功
#define FN_CLIP_GET_EPRG_DATA  (FNST_ADDIN | 6)

// 将易程序数据存放到当前系统剪贴板中
//   功能参数1(BYTE*): 指向易程序数据的指针.
//   功能参数2: 数据尺寸.
#define FN_CLIP_SET_EPRG_DATA  (FNST_ADDIN | 7)

// 返回当前光标所处行索引位置
//   功能参数1(int*): 在其中填入行索引位置
#define FN_GET_CARET_ROW_INDEX  (FNST_ADDIN | 8)

// 返回当前光标所处列索引位置
//   功能参数1(int*): 在其中填入列索引位置
#define FN_GET_CARET_COL_INDEX  (FNST_ADDIN | 9)

// 返回所指定位置处的类型及程序文本
// 功能参数1(GET_PRG_TEXT_PARAM*): 提供相关参数信息
#define FN_GET_PRG_TEXT  (FNST_ADDIN | 10)

// 返回所指定位置处的类型及帮助信息文本
// 功能参数1(GET_PRG_TEXT_PARAM*): 提供相关参数信息
#define FN_GET_PRG_HELP  (FNST_ADDIN | 11)

// 返回当前程序内所导入的易模块数目
//   功能参数1(int*): 提供填写返回数据的整数变量地址
#define FN_GET_NUM_ECOM   (FNST_ADDIN | 12)

// 返回所指定索引位置处易模块的全路径文件名
//   功能参数1(int): 提供欲查询易模块的索引位置,必须大于等于0小于FN_GET_NUM_ECOM的返回值.
//   功能参数2(char*): 提供用作填写返回文本的缓冲区地址,缓冲区尺寸必须大于等于256个字符.
#define FN_GET_ECOM_FILE_NAME  (FNST_ADDIN | 13)

// 返回当前系统内已经选择的支持库数目
//   功能参数1(int*): 提供填写返回数据的整数变量地址
#define FN_GET_NUM_LIB    (FNST_ADDIN | 14)

// 返回所指定索引位置处支持库的信息文本
//   功能参数1(int): 提供欲查询支持库的索引位置,必须大于等于0小于FN_GET_NUM_LIB的返回值.
//   功能参数2(char**): 提供填写返回文本指针的变量地址
#define FN_GET_LIB_INFO_TEXT  (FNST_ADDIN | 15)

#endif

