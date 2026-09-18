// 独立回归测试：链接真实镜像实现，用桩模拟 IDE 查询源码时同步通知镜像清理。
#include <Windows.h>
#include "../src/WorkspaceMirror.h"
#include "../src/EPackagerIntegration.h"
#include <functional>
#include <fstream>
#include <iostream>
#include <stdexcept>
#include <thread>
#include <atomic>
#include "../src/IdeOutputQueue.h"
#include "../src/ScopedTreeDragGuard.h"
#include "../src/IdeToolScheduler.h"

namespace {
bool sourceChanged = false;
bool sourceAvailable = true;
unsigned sourceNumber = 0;
unsigned resetCount = 0;
unsigned unpackCount = 0;
void Check(bool ok, const char* message)
{
    if (!ok) throw std::runtime_error(message);
}
}

HWND logWindow = nullptr;
std::atomic_uint logCount = 0;
std::atomic_bool failUnpack = false, failMain = false;
void OutputStringToELog(const std::string& text) { IdeOutputQueue::Instance().Push(logWindow, text); }
LRESULT CALLBACK TestWindowProc(HWND window, UINT message, WPARAM w, LPARAM l) {
    if (message == IdeOutputQueue::Instance().Message()) {
        logCount += static_cast<unsigned>(IdeOutputQueue::Instance().Take(window).size());
        return 0;
    }
    return DefWindowProc(window, message, w, l);
}
void Pump() { MSG message{}; while (PeekMessage(&message,nullptr,0,0,PM_REMOVE)) { TranslateMessage(&message); DispatchMessage(&message); } }


namespace EPackagerIntegration {
bool GetCurrentSourcePath(std::filesystem::path& path, std::string& error)
{
    // 与 UpdateCurrentOpenSourceFile -> OnCurrentSourceFilePathChanged 的同步回调一致。
    if (sourceChanged) {
        sourceChanged = false;
        ++sourceNumber;
        WorkspaceMirror::ResetAndCleanup();
        ++resetCount;
    }
    error.clear();
    path.clear();
    if (!sourceAvailable) {
        error = "no current source";
        return false;
    }
    path = std::filesystem::temp_directory_path() / ("mirror-reentry-" + std::to_string(sourceNumber) + ".e");
    return true;
}
std::filesystem::path BuildCurrentProjectSnapshotPathForSource(const std::filesystem::path& path)
{
    return path;
}
bool WriteCurrentProjectSnapshot(const std::filesystem::path&, size_t& bytes, std::string& trace, std::string& error)
{
    bytes = 1;
    trace.clear();
    error.clear();
    return true;
}
void CleanupSnapshotRoot(const std::filesystem::path&) {}
bool GetInstalledTool(std::filesystem::path& path, std::string& error)
{
    path = "test-packager.exe";
    error.clear();
    return true;
}
ProcessRunResult RunProcessAndCapture(const std::filesystem::path&, const std::vector<std::wstring>& args,
    const std::filesystem::path&, unsigned long, const std::function<bool()>& cancelled)
{
    ++unpackCount;
    OutputStringToELog("worker unpack log");
    if ((cancelled && cancelled()) || failUnpack || (failMain && args.size() > 3)) {
        ProcessRunResult failure; failure.error = "simulated failure"; return failure;
    }
    const std::filesystem::path root(args.at(2));
    std::filesystem::create_directories(root / "src");
    std::filesystem::create_directories(root / "project");
    std::ofstream(root / "src" / "page.txt") << "source content";
    std::ofstream(root / "project" / "_meta.json")
        << R"({"sourceFiles":[{"relativePath":"src/page.txt","logicalName":"page"}]})";
    ProcessRunResult result;
    result.ok = true;
    result.elapsedMs = 1;
    return result;
}
}


bool Refresh(WorkspaceMirror::RefreshMode mode, std::string& error, std::string* resultMode = nullptr) {
    auto task = WorkspaceMirror::CaptureRefreshInputOnMainThread(mode,error);
    if (!task) return false;
    std::thread worker([&] { WorkspaceMirror::BuildRefreshCandidate(task); });
    // 故意不泵 UI：后台日志必须不依赖 UI 线程完成。
    worker.join();
    std::string used;
    const bool ok = WorkspaceMirror::CommitRefreshOnMainThread(task,error,used);
    if(resultMode) *resultMode=used;
    return ok;
}
std::string Read(const std::filesystem::path& path) { std::ifstream in(path,std::ios::binary); return {std::istreambuf_iterator<char>(in),{}}; }
int main() {
    try {
        WNDCLASSW wc{}; wc.hInstance=GetModuleHandle(nullptr); wc.lpfnWndProc=TestWindowProc; wc.lpszClassName=L"MirrorRegression";
        RegisterClassW(&wc);
        logWindow=CreateWindowW(wc.lpszClassName,L"",0,0,0,0,0,HWND_MESSAGE,nullptr,wc.hInstance,nullptr);
        Check(logWindow != nullptr,"test window");
        std::string error, mode;
        for(auto refreshMode : {WorkspaceMirror::RefreshMode::Auto,WorkspaceMirror::RefreshMode::MainOnly,WorkspaceMirror::RefreshMode::Full}) {
            sourceChanged=true;
            Check(Refresh(refreshMode,error),error.c_str());
            WorkspaceMirror::FileAccessSnapshot before;
            Check(WorkspaceMirror::GetPreparedFileAccessSnapshot(before,error),"snapshot");
            Check(Refresh(refreshMode,error),error.c_str());
            Check(std::filesystem::exists(before.mirrorRoot/"src/page.txt"),"leased old mirror deleted");
            Check(WorkspaceMirror::EnsureMirrorFresh(error),"prepared access");
        }
        std::cout<<"PASS capture reentry and old-directory leases\n";
        auto stale=WorkspaceMirror::CaptureRefreshInputOnMainThread(WorkspaceMirror::RefreshMode::Full,error);
        WorkspaceMirror::BuildRefreshCandidate(stale);
        WorkspaceMirror::ResetAndCleanup(); // A -> B -> A，即使路径相同也必须拒绝。
        Check(!WorkspaceMirror::CommitRefreshOnMainThread(stale,error,mode)&&error=="workspace_changed","ABA candidate published");
        Check(Refresh(WorkspaceMirror::RefreshMode::Full,error),error.c_str());
        auto cancelled=WorkspaceMirror::CaptureRefreshInputOnMainThread(WorkspaceMirror::RefreshMode::Full,error);
        WorkspaceMirror::CancelRefresh(cancelled); WorkspaceMirror::BuildRefreshCandidate(cancelled);
        Check(!WorkspaceMirror::CommitRefreshOnMainThread(cancelled,error,mode)&&error=="cancelled","cancelled publish");
        std::cout<<"PASS stale and cancelled candidates\n";
        auto beforeGeneration=WorkspaceMirror::GetGeneration();
        failUnpack=true;
        Check(!Refresh(WorkspaceMirror::RefreshMode::Full,error),"failure accepted");
        Check(WorkspaceMirror::GetGeneration()==beforeGeneration,"failed build changed mirror");
        failUnpack=false; failMain=true;
        Check(Refresh(WorkspaceMirror::RefreshMode::MainOnly,error,&mode)&&mode=="full","main-only fallback"); failMain=false;
        std::cout<<"PASS failed candidate isolation and main-only fallback\n";
        WorkspaceMirror::FileAccessSnapshot old;
        Check(WorkspaceMirror::GetPreparedFileAccessSnapshot(old,error),"old snapshot");
        auto update=WorkspaceMirror::CaptureTextUpdateOnMainThread("src/page.txt","updated",error);
        Check(update!=nullptr,"update capture");
        std::thread writer([&] { WorkspaceMirror::BuildRefreshCandidate(update); }); writer.join();
        Check(WorkspaceMirror::CommitRefreshOnMainThread(update,error,mode),error.c_str());
        WorkspaceMirror::FileAccessSnapshot current;
        Check(WorkspaceMirror::GetPreparedFileAccessSnapshot(current,error),"new snapshot");
        Check(Read(old.mirrorRoot/"src/page.txt")=="source content","old reader saw mutation");
        Check(Read(current.mirrorRoot/"src/page.txt").find("updated")!=std::string::npos,"update missing");
        std::cout<<"PASS immutable incremental update\n";
        auto external=std::filesystem::temp_directory_path()/"autolinker-reference-test.ec";
        std::ofstream(external)<<"test";
        auto reference=WorkspaceMirror::CaptureReferenceInputOnMainThread(external.string(),error);
        Check(reference!=nullptr,"reference capture"); WorkspaceMirror::BuildRefreshCandidate(reference);
        Check(WorkspaceMirror::CommitRefreshOnMainThread(reference,error,mode),error.c_str());
        const auto referenceResult=WorkspaceMirror::GetReferenceResult(reference);
        Check(referenceResult.fileCount>0,"reference empty");
        Check(Refresh(WorkspaceMirror::RefreshMode::Full,error),error.c_str());
        Check(WorkspaceMirror::GetPreparedFileAccessSnapshot(current,error),"reference snapshot");
        Check(std::filesystem::exists(current.mirrorRoot/referenceResult.outputDirectoryUtf8),"reference lost");
        std::filesystem::remove(external);
        sourceAvailable=false; sourceChanged=true;
        Check(!Refresh(WorkspaceMirror::RefreshMode::Full,error)&&error=="no current source","closed project");
        Check(std::filesystem::exists(old.mirrorRoot),"close removed leased directory");
        std::cout<<"PASS reference preservation and source close\n";
        Pump();
        OutputStringToELog(std::string(4 * 1024 * 1024 + 1, 'x'));
        const auto oversizedBatch = IdeOutputQueue::Instance().Take(logWindow);
        Check(oversizedBatch.size() == 1 && oversizedBatch.front().find("1 lines omitted") != std::string::npos,
            "oversized log escaped bounded UI queue");
        Pump();
        std::thread logger([]{for(int i=0;i<1200;++i) OutputStringToELog("background line");}); logger.join();
        Pump(); Check(logCount>0,"async logs not drained");
        std::cout<<"PASS background logging without UI wait and bounded overflow\n";
        INITCOMMONCONTROLSEX ic{sizeof(ic),ICC_TREEVIEW_CLASSES}; InitCommonControlsEx(&ic);
        HWND tree=CreateWindowExW(0,WC_TREEVIEWW,L"",WS_CHILD,0,0,300,300,logWindow,nullptr,wc.hInstance,nullptr);
        TVINSERTSTRUCTW ins{}; ins.hParent=TVI_ROOT; ins.hInsertAfter=TVI_LAST; ins.item.mask=TVIF_TEXT;
        wchar_t label[]=L"page"; ins.item.pszText=label;
        auto item=reinterpret_cast<HTREEITEM>(SendMessageW(tree,TVM_INSERTITEMW,0,reinterpret_cast<LPARAM>(&ins)));
        RECT rect{}; *reinterpret_cast<HTREEITEM*>(&rect)=item;
        SendMessageW(tree,TVM_GETITEMRECT,TRUE,reinterpret_cast<LPARAM>(&rect));
        auto style=GetWindowLongPtr(tree,GWL_STYLE);
        { ScopedTreeDragGuard guard(tree);
          DWORD_PTR result=0; auto start=GetTickCount64();
          Check(SendMessageTimeoutW(tree,WM_LBUTTONDOWN,MK_LBUTTON,MAKELPARAM(rect.left+2,rect.top+2),SMTO_BLOCK|SMTO_ABORTIFHUNG,50,&result)!=0,"tree down failed");
          Check(GetTickCount64()-start<1000,"tree down waited for physical input");
          SendMessageW(tree,WM_LBUTTONUP,0,0);
        }
        Check(GetWindowLongPtr(tree,GWL_STYLE)==style,"tree style not restored");
        DestroyWindow(tree);
        std::cout<<"PASS tree fallback no physical mouse release required\n";
        { IdeToolScheduler scheduler;
          scheduler.Start(logWindow,[]{OutputStringToELog("scheduler worker");});
          auto deadline=GetTickCount64()+3000;
          while(!scheduler.WorkerDone()&&GetTickCount64()<deadline) Sleep(1);
          Check(scheduler.WorkerDone(),"scheduler worker blocked on UI");
          scheduler.JoinCompleted(logWindow);
          scheduler.Stop(logWindow);
        }
        WorkspaceMirror::ResetAndCleanup();
        IdeOutputQueue::Instance().Stop(); DestroyWindow(logWindow);
        std::cout<<"PASS all workspace concurrency regressions\n";
        return 0;
    } catch(const std::exception& ex) { std::cerr<<"FAIL "<<ex.what()<<'\n'; return 1; }
}
