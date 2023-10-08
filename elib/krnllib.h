#ifndef __KRNL_LIB_H
#define __KRNL_LIB_H

////////////////////

#define		DTC_WIN_FORM			1
//#define		DTC_RPT_FORM			2	// 已废止
#define		DTC_MENU				3
#define		DTC_FONT				4
#define		DTC_EDIT				5
#define		DTC_PIC_BOX				6
#define		DTC_SHAPE_BOX			7
#define		DTC_DRAW_PANEL			8
#define		DTC_GROUP_BOX			9
#define		DTC_LABEL				10
#define		DTC_BUTTON				11
#define		DTC_CHECK_BOX			12
#define		DTC_RADIO_BOX			13
#define		DTC_COMBO_BOX			14
#define		DTC_LIST_BOX			15
#define		DTC_CHKLIST_BOX			16
#define		DTC_HSCROLL_BAR			17
#define		DTC_VSCROLL_BAR			18
#define		DTC_PROCESS_BAR			19
#define		DTC_SLIDER_BAR			20
#define		DTC_TAB					21
#define		DTC_ANIMATE				22
#define		DTC_DATE_TIME_PICKER	23
#define		DTC_MONTH_CALENDAR		24
#define		DTC_DRIVER_BOX			25
#define		DTC_DIR_BOX				26
#define		DTC_FILE_BOX			27
#define		DTC_COLOR_PICKER		28
#define		DTC_HYPER_LINKER		29
#define		DTC_SPIN				30
#define		DTC_COMMON_DLG			31
#define		DTC_TIMER				32
#define		DTC_PRINTER				33
#define		DTC_FIELD_INF			34
#define		DTC_HTML_VIEWER			35
#define		DTC_UDP					36
#define		DTC_SOCK_CLIENT			37
#define		DTC_SOCK_SERVER			38
#define		DTC_SERIAL_PORT			39
#define		DTC_PRINT_INF			40
#define		DTC_GRID				41
#define		DTC_DATA_SOURCE			42
#define		DTC_NPROVIDER			43
#define		DTC_DBPROVIDER			44
#define		DTC_RGN_BUTTON			45
#define		DTC_ODBC_DB				46
#define		DTC_ODBCPROVIDER		47
#define		DTC_COM_OBJECT  		48
#define     DTC_VARIANT             49
#define     DTC_VARIANT_TYPE        50
#define     DTC_ERROR_CTL           51

/////////////
#define         DTP_WIN_FORM         ( MAKELONG( DTC_WIN_FORM        , 1) )
//#define         DTP_RPT_FORM         ( MAKELONG( DTC_RPT_FORM        , 1) )
#define         DTP_MENU             ( MAKELONG( DTC_MENU            , 1) )
#define         DTP_FONT             ( MAKELONG( DTC_FONT            , 1) )
#define         DTP_EDIT             ( MAKELONG( DTC_EDIT            , 1) )
#define         DTP_PIC_BOX          ( MAKELONG( DTC_PIC_BOX         , 1) )
#define         DTP_SHAPE_BOX        ( MAKELONG( DTC_SHAPE_BOX       , 1) )
#define         DTP_DRAW_PANEL       ( MAKELONG( DTC_DRAW_PANEL      , 1) )
#define         DTP_GROUP_BOX        ( MAKELONG( DTC_GROUP_BOX       , 1) )
#define         DTP_LABEL            ( MAKELONG( DTC_LABEL           , 1) )
#define         DTP_BUTTON           ( MAKELONG( DTC_BUTTON          , 1) )
#define         DTP_CHECK_BOX        ( MAKELONG( DTC_CHECK_BOX       , 1) )
#define         DTP_RADIO_BOX        ( MAKELONG( DTC_RADIO_BOX       , 1) )
#define         DTP_COMBO_BOX        ( MAKELONG( DTC_COMBO_BOX       , 1) )
#define         DTP_LIST_BOX         ( MAKELONG( DTC_LIST_BOX        , 1) )
#define         DTP_CHKLIST_BOX      ( MAKELONG( DTC_CHKLIST_BOX     , 1) )
#define         DTP_HSCROLL_BAR      ( MAKELONG( DTC_HSCROLL_BAR     , 1) )
#define         DTP_VSCROLL_BAR      ( MAKELONG( DTC_VSCROLL_BAR     , 1) )
#define         DTP_PROCESS_BAR      ( MAKELONG( DTC_PROCESS_BAR     , 1) )
#define         DTP_SLIDER_BAR       ( MAKELONG( DTC_SLIDER_BAR      , 1) )
#define         DTP_TAB              ( MAKELONG( DTC_TAB             , 1) )
#define         DTP_ANIMATE          ( MAKELONG( DTC_ANIMATE         , 1) )
#define         DTP_DATE_TIME_PICKER ( MAKELONG( DTC_DATE_TIME_PICKER, 1) )
#define         DTP_MONTH_CALENDAR   ( MAKELONG( DTC_MONTH_CALENDAR  , 1) )
#define         DTP_DRIVER_BOX       ( MAKELONG( DTC_DRIVER_BOX      , 1) )
#define         DTP_DIR_BOX          ( MAKELONG( DTC_DIR_BOX         , 1) )
#define         DTP_FILE_BOX         ( MAKELONG( DTC_FILE_BOX        , 1) )
#define         DTP_COLOR_PICKER     ( MAKELONG( DTC_COLOR_PICKER    , 1) )
#define         DTP_HYPER_LINKER     ( MAKELONG( DTC_HYPER_LINKER    , 1) )
#define         DTP_SPIN             ( MAKELONG( DTC_SPIN            , 1) )
#define         DTP_COMMON_DLG       ( MAKELONG( DTC_COMMON_DLG      , 1) )
#define         DTP_TIMER            ( MAKELONG( DTC_TIMER           , 1) )
#define         DTP_PRINTER          ( MAKELONG( DTC_PRINTER         , 1) )
#define         DTP_FIELD_INF        ( MAKELONG( DTC_FIELD_INF       , 1) )
#define         DTP_HTML_VIEWER      ( MAKELONG( DTC_HTML_VIEWER     , 1) )
#define         DTP_UDP              ( MAKELONG( DTC_UDP             , 1) )
#define         DTP_SOCK_CLIENT      ( MAKELONG( DTC_SOCK_CLIENT     , 1) )
#define         DTP_SOCK_SERVER      ( MAKELONG( DTC_SOCK_SERVER     , 1) )
#define         DTP_SERIAL_PORT      ( MAKELONG( DTC_SERIAL_PORT     , 1) )
#define         DTP_PRINT_INF        ( MAKELONG( DTC_PRINT_INF       , 1) )
#define         DTP_GRID             ( MAKELONG( DTC_GRID            , 1) )
#define         DTP_DATA_SOURCE      ( MAKELONG( DTC_DATA_SOURCE     , 1) )
#define         DTP_NPROVIDER        ( MAKELONG( DTC_NPROVIDER       , 1) )
#define         DTP_DBPROVIDER       ( MAKELONG( DTC_DBPROVIDER      , 1) )
#define         DTP_RGN_BUTTON       ( MAKELONG( DTC_RGN_BUTTON      , 1) )
#define         DTP_ODBC_DB          ( MAKELONG( DTC_ODBC_DB         , 1) )
#define         DTP_ODBCPROVIDER     ( MAKELONG( DTC_ODBCPROVIDER    , 1) )
#define         DTP_COM_OBJECT       ( MAKELONG( DTC_COM_OBJECT      , 1) )
#define         DTP_VARIANT          ( MAKELONG( DTC_VARIANT         , 1) )
#define         DTP_VARIANT_TYPE     ( MAKELONG( DTC_VARIANT_TYPE    , 1) )
#define         DTP_ERROR_CTL        ( MAKELONG( DTC_ERROR_CTL       , 1) )


/////////////

// 系统核心支持库的GUID串：
// guid: { 0xd09f2340, 0x8185, 0x11d3, { 0x96, 0xf6, 0xaa, 0xf8, 0x44, 0xc7, 0xe3, 0x25 } }
#define		LI_KRNL_LIB_GUID_STR	"d09f2340818511d396f6aaf844c7e325"

//   系统核心支持库的文件名（由于不带后缀，所以不为实际的文件名），实际文件
// 名称根据其类型为"krnln.fne"、"krnln.fnl"、"krnln.fnr"等。
#define		LI_KRNL_LIB_FILE_NAME	"krnln"
#define		LI_KRNL_LIB_NAME		"系统核心支持库"	// 系统核心支持库的库名

#define		LI_KRNL_LIB_MAJOR_VER	    4					// 本系统核心支持库的主版本
#define		LI_KRNL_LIB_MINOR_VER	    5					// 本系统核心支持库的次版本
#define		LI_KRNL_LIB_MAJOR_VER_STR	"4"					// 本系统核心支持库的主版本文本
#define		LI_KRNL_LIB_MINOR_VER_STR	"5"					// 本系统核心支持库的次版本文本

// ??? 错误取代命令在核心库中的所处位置索引。
#define     ERR_CMD_INDEX       622
// ??? 错误取代常量在核心库中的所处位置索引。
#define     ERR_CONST_INDEX     330

#endif
