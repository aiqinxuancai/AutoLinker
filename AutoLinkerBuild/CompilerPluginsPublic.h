
#ifndef __COMPILER_PLUSINS_PUBLIC_H__
#define __COMPILER_PLUSINS_PUBLIC_H__

typedef enum
{
    // �޸Ĵ��������������Ϣ
    //   dwParam1: COMPILE_PROCESSOR_PRG_INFO* ָ���µĳ�����Ϣ,������Ч�Ҳ�ΪNULL.
    PCFN_SET_PRG_INFO = 1,

    // �޸Ĵ����������ָ����ֵ�ͳ�����ֵ,�ɹ�����1,ʧ�ܷ���0.
    //   dwParam1: const TCHAR* ָ����޸���ֵ����������
    //   dwParam2: INT �����޸ĵ��ĳ���ֵ
    PCFN_SET_PRG_NUM_CONST,

    // �޸Ĵ����������ָ������ʱ���ͳ�����ֵ,�ɹ�����1,ʧ�ܷ���0.
    //   dwParam1: const TCHAR* ָ����޸�����ʱ���ͳ���������
    //   dwParam2: DATE* ָ�������޸ĵ��ĳ���ֵ,������Ч�Ҳ�ΪNULL.
    PCFN_SET_PRG_DATE_TIME_CONST,

    // �޸Ĵ����������ָ���߼��ͳ�����ֵ,�ɹ�����1,ʧ�ܷ���0.
    //   dwParam1: const TCHAR* ָ����޸��߼��ͳ���������
    //   dwParam2: DWORD �����޸ĵ��ĳ���ֵ,Ϊ1��ʾ��,Ϊ0��ʾ��.
    PCFN_SET_PRG_BOOL_CONST,

    // �޸Ĵ����������ָ���ı��ͳ�����ֵ,�ɹ�����1,ʧ�ܷ���0.
    //   dwParam1: const TCHAR* ָ����޸��ı��ͳ���������
    //   dwParam2: const TCHAR* ָ�������޸ĵ��ĳ����ı�ֵ
    PCFN_SET_PRG_TEXT_CONST,

    // �޸Ĵ����������ָ��ͼƬ�������ೣ����ֵ,�ɹ�����TRUE,ʧ�ܷ���FALSE.
    //   dwParam1: const TCHAR* ָ����޸�ͼƬ�������ೣ��������
    //   dwParam2: BYTE* ����Ϊһ��INT��¼ͼƬ���������ݵĳ���(��Ϊ��,��ֵΪ0),�����Ӧ���ȵ�ͼƬ�������ֽ�����.������Ч�Ҳ�ΪNULL.
    PCFN_SET_PRG_BIN_CONST,

    // ���ش����������ָ����ֵ/�߼����ͳ�����ֵ.
    // ������ֵ����,ֻ�ܷ�������������,�����߼��ͳ���,Ϊ�淵��TRUE,���򷵻�FALSE.
    // ���δ�ҵ���ָ�����Ƶĳ������߸ó�����Ϊ��ֵ/�߼���,��ֱ�ӷ���0.
    //   dwParam1: const TCHAR* ָ����������ֵ/�߼��ͳ���������
    PCFN_GET_PRG_CONST,
}
CALL_BACK_FUNC_NO;

// pCallBackData: COMPILE_PROCESSOR_INFO.m_pCallBackDataֵ
typedef DWORD (WINAPI *FN_COMPILE_CALLBACK) (const void* pCallBackData,
        const DWORD dwFuncNO, const DWORD dwParam1, const DWORD dwParam2);

typedef struct
{
	const TCHAR* m_szPrgName;	// ��������
	const TCHAR* m_szExplain;	// �йر�����������ͱ�ע

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

    const TCHAR* m_szSaveToFileName;  // �����������浽���ļ�ȫ·����,Ϊ�ձ�ʾ��Ҫ�û��ֹ�����.
}
COMPILE_PROCESSOR_PRG_INFO;

typedef struct
{
    FN_COMPILE_CALLBACK m_fnCallBack;  // �ṩ����������ص��������������Ӧ����
    void* m_pCallBackData;  // �������ṩ�����,����ڵ���m_fnCallBackʱ����ԭֵ����.

    //---------------------------------------------------

    BOOL m_blCompileReleaseVersion;  // Ϊ�����Ϊ���뷢���汾,����Ϊ������԰汾.
    BOOL m_blCompileEPackage;  // Ϊ�����Ϊ���ڱ����װ�
    const TCHAR* m_szPrgFileName;  // �����������ļ�����(����·���ͺ�׺),����(�ó�����δ����)��Ϊ��.
    const TCHAR* m_szPurePrgFileName;  // ���������Ĵ����ļ�����(������·���ͺ�׺),����(�ó�����δ����)��Ϊ��.

    //---------------------------------------------------

    DWORD	m_dwSysPrgVersion;  // ϵͳ����İ汾��.
	INT		m_nSysType;			// ϵͳ��������ͺ�
	INT		m_nLanguageVer;		// ϵͳ��������԰汾.
    // �������������ͣ�!!!���к�ֵ���Բ��ɸı䣬���к�ֵ��������㣩��
    #define AT_WIN_WINDOWS_EXE      0
    #define AT_WIN_CONSOLE_EXE      1
    #define AT_WIN_DLL              2
    #define AT_WIN_ECOM             1000
    #define AT_LINUX_CONSOLE_ELF    10000
    #define AT_LINUX_ECOM           11000
    INT m_nAppType;

    COMPILE_PROCESSOR_PRG_INFO m_infPrg;  // ���������������Ϣ
}
COMPILE_PROCESSOR_INFO;

// ����ǰ����Ķ���ӿں���ԭ��. �������ʾ����ɹ�,���ؼٱ�ʾ����ʧ��.
//   fnCallBack: ���������������
//   pInf: �ṩ��ر���ʱ��Ϣ,�ض���ΪNULL.
typedef BOOL (WINAPI *FN_COMPILE_PROCESSOR) (const COMPILE_PROCESSOR_INFO* pInf);

// ����������������
#define COMPILER_PLUGINS_EXPORT_NAME  "CompileProcessor"

#endif
