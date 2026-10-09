#include "BackupRestorePage.h"
#include <CommCtrl.h>
#include <algorithm>
#include <vector>
#include "BackupRestore.h"
#include "AIChatSourceBackup.h"
#include "EideProjectRestore.h"

namespace {
constexpr wchar_t kClass[] = L"AutoLinker.BackupRestore.Settings";
struct Page {
    HWND list{}, status{}, sourceLabel{}, restore{}, refresh{};
    HFONT font{};
    bool restoring = false;
    std::filesystem::path source;
    std::vector<std::filesystem::path> backups;
};
HMODULE Module() {
    HMODULE module{};
    GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
        reinterpret_cast<LPCWSTR>(&Module), &module);
    return module;
}
void Selection(Page* p) {
    const int selected = ListView_GetNextItem(p->list, -1, LVNI_SELECTED);
    EnableWindow(p->restore, !p->restoring && EideProjectRestore::IsSupported() &&
        selected >= 0 && static_cast<size_t>(selected) < p->backups.size() &&
        _wcsicmp(p->source.extension().c_str(), L".e") == 0);
}
void Refresh(Page* p) {
    if (p->restoring) return;
    ListView_DeleteAllItems(p->list);
    p->backups.clear();
    try {
        p->source = BackupRestore::CurrentSource();
        const auto label = p->source.empty() ? L"当前未打开已保存的工程" : p->source.wstring();
        SetWindowTextW(p->sourceLabel, label.c_str());
        if (!p->source.empty() && p->source.is_absolute()) p->backups = AIChatSourceBackup::ListBackups(p->source);
        for (size_t i = 0; i < p->backups.size(); ++i) {
            const auto name = p->backups[i].filename().wstring();
            const auto date = name.substr(0,4) + L"-" + name.substr(4,2) + L"-" + name.substr(6,2) + L" " +
                name.substr(8,2) + L":" + name.substr(10,2) + L":" + name.substr(12,2);
            LVITEMW item{}; item.mask = LVIF_TEXT; item.iItem = static_cast<int>(i); item.pszText = const_cast<wchar_t*>(date.c_str());
            SendMessageW(p->list, LVM_INSERTITEMW, 0, reinterpret_cast<LPARAM>(&item));
            item.iSubItem = 1; item.pszText = const_cast<wchar_t*>(name.c_str());
            SendMessageW(p->list, LVM_SETITEMTEXTW, i, reinterpret_cast<LPARAM>(&item));
            const auto size = std::to_wstring(std::filesystem::file_size(p->backups[i])) + L" 字节";
            item.iSubItem = 2; item.pszText = const_cast<wchar_t*>(size.c_str());
            SendMessageW(p->list, LVM_SETITEMTEXTW, i, reinterpret_cast<LPARAM>(&item));
        }
        const wchar_t* status = !EideProjectRestore::IsSupported() ? L"无法恢复，此功能需要 e5.95 版本。当前 IDE 内部地址校验不匹配。" :
            p->backups.empty() ? L"当前工程暂无自动备份。备份保存在源码同目录的 {文件名}.BAK 中。" :
            _wcsicmp(p->source.extension().c_str(), L".e") != 0 ? L"目前仅支持 .e 工程恢复。" :
            L"选择备份后点击恢复。恢复前会自动备份当前内容；原工程路径和磁盘源码保持不变。";
        SetWindowTextW(p->status, status);
    } catch (...) { SetWindowTextW(p->status, L"读取备份目录失败，请检查目录权限后刷新。" ); }
    Selection(p);
}
void Layout(HWND window, Page* p) {
    RECT r{}; GetClientRect(window, &r);
    const int width = (std::max)(100L, r.right - 48);
    MoveWindow(GetDlgItem(window, 10), 24, 22, width, 30, TRUE);
    MoveWindow(p->sourceLabel, 24, 60, width, 28, TRUE);
    MoveWindow(p->list, 24, 100, width, (std::max)(80L, r.bottom - 270), TRUE);
    const int bottom = (std::max)(190L, r.bottom - 152);
    MoveWindow(p->restore, 24, bottom, 154, 34, TRUE);
    MoveWindow(p->refresh, 192, bottom, 100, 34, TRUE);
    MoveWindow(p->status, 24, bottom + 48, width, 90, TRUE);
    ListView_SetColumnWidth(p->list, 1, (std::max)(200, width - 305));
}
LRESULT CALLBACK Proc(HWND window, UINT msg, WPARAM w, LPARAM l) {
    auto* p = reinterpret_cast<Page*>(GetWindowLongPtrW(window, GWLP_USERDATA));
    if (msg == WM_NCCREATE) { p = new Page; SetWindowLongPtrW(window,GWLP_USERDATA,reinterpret_cast<LONG_PTR>(p)); return TRUE; }
    if (!p) return DefWindowProcW(window,msg,w,l);
    switch (msg) {
    case WM_CREATE: {
        p->font = CreateFontW(-16,0,0,0,FW_NORMAL,FALSE,FALSE,FALSE,DEFAULT_CHARSET,OUT_DEFAULT_PRECIS,CLIP_DEFAULT_PRECIS,CLEARTYPE_QUALITY,DEFAULT_PITCH,L"Microsoft YaHei UI");
        auto control = [&](const wchar_t* cls, const wchar_t* text, DWORD style, int id) {
            HWND h=CreateWindowExW(0,cls,text,WS_CHILD|WS_VISIBLE|style,0,0,0,0,window,reinterpret_cast<HMENU>(static_cast<INT_PTR>(id)),Module(),nullptr);
            SendMessageW(h,WM_SETFONT,reinterpret_cast<WPARAM>(p->font),TRUE); return h;
        };
        control(L"STATIC",L"备份与恢复",SS_LEFT,10);
        p->sourceLabel=control(L"STATIC",L"",SS_PATHELLIPSIS,11);
        p->status=control(L"STATIC",L"",SS_LEFT,12);
        p->list=control(WC_LISTVIEWW,L"",WS_TABSTOP|WS_BORDER|LVS_REPORT|LVS_SINGLESEL|LVS_SHOWSELALWAYS,20);
        ListView_SetExtendedListViewStyle(p->list,LVS_EX_FULLROWSELECT|LVS_EX_DOUBLEBUFFER|LVS_EX_LABELTIP);
        const wchar_t* labels[]={L"备份时间",L"备份文件 / 需求摘要",L"大小"};
        for(int i=0;i<3;++i) { LVCOLUMNW c{};c.mask=LVCF_TEXT|LVCF_WIDTH;c.pszText=const_cast<wchar_t*>(labels[i]);c.cx=i==0?175:i==1?450:125;SendMessageW(p->list,LVM_INSERTCOLUMNW,i,reinterpret_cast<LPARAM>(&c)); }
        p->restore=control(L"BUTTON",L"恢复选中备份",WS_TABSTOP|BS_PUSHBUTTON,21);
        p->refresh=control(L"BUTTON",L"刷新列表",WS_TABSTOP|BS_PUSHBUTTON,22);
        Refresh(p);Layout(window,p);return 0;
    }
    case WM_SIZE: Layout(window,p);return 0;
    case WM_SHOWWINDOW: if(w) Refresh(p);break;
    case WM_NOTIFY: if(reinterpret_cast<NMHDR*>(l)->hwndFrom==p->list) Selection(p);break;
    case WM_COMMAND:
        if(p->restoring) return 0;
        if(LOWORD(w)==22) { Refresh(p);return 0; }
        if(LOWORD(w)==21) {
            const int index=ListView_GetNextItem(p->list,-1,LVNI_SELECTED);
            if(index<0 || static_cast<size_t>(index)>=p->backups.size()) return 0;
            // 复制选中项，刷新不能改变本次恢复目标。源路径也使用列表生成时的值。
            const auto backup=p->backups[index], source=p->source;
            p->restoring=true;Selection(p);EnableWindow(p->refresh,FALSE);
            HWND root=GetAncestor(window,GA_ROOT);EnableWindow(root,FALSE);
            SetWindowTextW(p->status,L"正在保存恢复前安全备份并重新加载工程内容…");UpdateWindow(window);
            std::filesystem::path safety;std::wstring result;
            const bool ok=BackupRestore::Restore(source,backup,safety,result);
            EnableWindow(root,TRUE);p->restoring=false;EnableWindow(p->refresh,TRUE);Refresh(p);
            SetWindowTextW(p->status,result.c_str());
            MessageBoxW(window,result.c_str(),L"备份恢复",MB_OK|(ok?MB_ICONINFORMATION:MB_ICONWARNING));
            return 0;
        }
        break;
    case WM_CTLCOLORSTATIC: SetBkColor(reinterpret_cast<HDC>(w),RGB(255,255,255));return reinterpret_cast<LRESULT>(GetStockObject(WHITE_BRUSH));
    case WM_NCDESTROY: if(p->font) DeleteObject(p->font);delete p;SetWindowLongPtrW(window,GWLP_USERDATA,0);break;
    }
    return DefWindowProcW(window,msg,w,l);
}
}
HWND CreateBackupRestoreSettingsPage(HWND parent) {
    INITCOMMONCONTROLSEX controls{sizeof(controls),ICC_LISTVIEW_CLASSES};InitCommonControlsEx(&controls);
    WNDCLASSEXW c{};c.cbSize=sizeof(c);c.lpfnWndProc=Proc;c.hInstance=Module();c.hCursor=LoadCursor(nullptr,IDC_ARROW);c.hbrBackground=reinterpret_cast<HBRUSH>(GetStockObject(WHITE_BRUSH));c.lpszClassName=kClass;RegisterClassExW(&c);
    return CreateWindowExW(WS_EX_CONTROLPARENT,kClass,L"",WS_CHILD|WS_CLIPCHILDREN|WS_CLIPSIBLINGS,0,0,0,0,parent,nullptr,Module(),nullptr);
}

