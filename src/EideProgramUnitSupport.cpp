#include "EideProgramUnitSupport.h"

#include <Windows.h>
#include <CommCtrl.h>
#include <cstring>
#include <fstream>
#include <string_view>

#include "../thirdparty/json.hpp"
#include "IDEFacade.h"
#include "WorkspaceMirror.h"

namespace e571 {
namespace {
constexpr INT kDeleteSelectedProgramItem = static_cast<INT>(0x82020017u);

std::string_view Trim(std::string_view text)
{
    const auto first = text.find_first_not_of(" \t\r\n");
    if (first == std::string_view::npos) return {};
    return text.substr(first, text.find_last_not_of(" \t\r\n") - first + 1);
}

// 镜像来自当前工程快照。窗口绑定使用显式元数据，类类型使用导出声明，不猜名称或图标。
bool ReadDeletionMetadata(const std::string& filePath, const std::string& expectedName,
    WorkspaceMirror::FileAccessSnapshot& snapshot, std::string& fileType, std::string& trace)
{
    WorkspaceMirror::ProgramItemRef mapped;
    std::string error, relativePath;
    std::filesystem::path sourcePath;
    if (!WorkspaceMirror::ResolveFileToProgramItem(filePath, mapped, error) ||
        !mapped.editable || mapped.fixedTable || mapped.formXml || mapped.pageNameLocal != expectedName ||
        !WorkspaceMirror::GetPreparedFileAccessSnapshot(snapshot, error) ||
        !WorkspaceMirror::ResolvePreparedFilePath(snapshot, mapped.relativePathUtf8, sourcePath, relativePath, error)) {
        trace += "|program_unit_metadata_unavailable|" + error;
        return false;
    }
    std::ifstream input(snapshot.mirrorRoot / L"project" / L"_meta.json", std::ios::binary);
    if (!input.is_open()) input.open(snapshot.mirrorRoot / L"src" / L"_meta.json", std::ios::binary);
    const auto meta = nlohmann::json::parse(input, nullptr, false);
    if (!meta.is_object() || !meta.contains("sourceFiles") || !meta["sourceFiles"].is_array() ||
        !meta.contains("windowBindings") || !meta["windowBindings"].is_array()) {
        trace += "|program_unit_metadata_invalid";
        return false;
    }
    std::string nameUtf8;
    unsigned matches = 0;
    for (const auto& row : meta["sourceFiles"]) {
        if (!row.is_object() || !row.contains("relativePath") || !row["relativePath"].is_string() ||
            !row.contains("logicalName") || !row["logicalName"].is_string()) {
            trace += "|program_unit_metadata_invalid";
            return false;
        }
        if (row["relativePath"] == relativePath) {
            nameUtf8 = row["logicalName"].get<std::string>();
            ++matches;
        }
    }
    if (matches != 1 || nameUtf8.empty()) {
        trace += "|program_unit_metadata_ambiguous";
        return false;
    }
    for (const auto& binding : meta["windowBindings"]) {
        if (!binding.is_object() || !binding.contains("className") || !binding["className"].is_string() ||
            !binding.contains("formName") || !binding["formName"].is_string() ||
            binding["className"].get_ref<const std::string&>().empty() ||
            binding["formName"].get_ref<const std::string&>().empty()) {
            trace += "|window_binding_metadata_invalid";
            return false;
        }
        if (binding["className"] == nameUtf8) {
            trace += "|window_bound_assembly_not_supported";
            return false;
        }
    }
    std::ifstream source(sourcePath, std::ios::binary);
    const std::string_view directive(reinterpret_cast<const char*>(u8".程序集"));
    std::string line;
    while (std::getline(source, line)) {
        auto text = Trim(line);
        if (!text.starts_with(directive) || text.size() <= directive.size() ||
            (text[directive.size()] != ' ' && text[directive.size()] != '\t')) continue;
        text = Trim(text.substr(directive.size()));
        const auto comma = text.find(',');
        if (Trim(text.substr(0, comma)) != nameUtf8) break;
        auto base = comma == std::string_view::npos ? std::string_view{} : text.substr(comma + 1);
        base = Trim(base.substr(0, base.find(',')));
        fileType = base.empty() ? "assembly" : "class";
        trace += "|mirror_program_unit_metadata_ok";
        return true;
    }
    trace += "|program_unit_declaration_invalid";
    return false;
}

struct TreeTarget {
    unsigned int data;
    const std::string& name;
    HWND tree = nullptr;
    HTREEITEM item = nullptr;
    unsigned matches = 0;
};

bool MatchesTarget(HWND tree, HTREEITEM item, const TreeTarget& target)
{
    char name[1024]{};
    TVITEMA info{};
    info.mask = TVIF_TEXT | TVIF_PARAM;
    info.hItem = item;
    info.pszText = name;
    info.cchTextMax = static_cast<int>(sizeof(name));
    return item && SendMessageA(tree, TVM_GETITEMA, 0, reinterpret_cast<LPARAM>(&info)) &&
        static_cast<unsigned int>(info.lParam) == target.data && name == target.name;
}

void FindTarget(HWND tree, HTREEITEM first, TreeTarget& target, unsigned depth = 0)
{
    if (depth > 64) return;
    for (auto item = first; item; item = TreeView_GetNextSibling(tree, item)) {
        if (MatchesTarget(tree, item, target)) {
            target.tree = tree;
            target.item = item;
            ++target.matches;
        }
        FindTarget(tree, TreeView_GetChild(tree, item), target, depth + 1);
    }
}

BOOL CALLBACK FindProgramTree(HWND hwnd, LPARAM param)
{
    char className[64]{};
    GetClassNameA(hwnd, className, static_cast<int>(sizeof(className)));
    if (std::strcmp(className, WC_TREEVIEWA) != 0) return TRUE;
    wchar_t name[64]{};
    TVITEMW root{};
    root.mask = TVIF_TEXT;
    root.hItem = TreeView_GetRoot(hwnd);
    root.pszText = name;
    root.cchTextMax = static_cast<int>(std::size(name));
    if (root.hItem && SendMessageW(hwnd, TVM_GETITEMW, 0, reinterpret_cast<LPARAM>(&root)) &&
        std::wstring_view(name) == L"程序数据") {
        FindTarget(hwnd, root.hItem, *reinterpret_cast<TreeTarget*>(param));
    }
    return TRUE;
}
}

bool DeleteProgramUnitByTreeItemData(unsigned int itemData, const std::string& expectedName,
    const std::string& filePathUtf8, std::string& outFileType, std::string& outTrace)
{
    outFileType.clear();
    outTrace = "program_tree_delete";
    auto& ide = IDEFacade::Instance();
    const HWND mainWindow = ide.GetMainWindow();
    if (!IsWindow(mainWindow) || GetWindowThreadProcessId(mainWindow, nullptr) != GetCurrentThreadId()) {
        outTrace += "|requires_ide_main_thread";
        return false;
    }
    if ((itemData >> 28) != 1 || (itemData & 0xFFFF) != 0 ||
        (itemData & 0x0FFF0000) == 0 || expectedName.empty()) {
        outTrace += "|invalid_program_unit_item";
        return false;
    }
    WorkspaceMirror::FileAccessSnapshot snapshot;
    if (!ReadDeletionMetadata(filePathUtf8, expectedName, snapshot, outFileType, outTrace)) return false;
    const auto contextEpoch = WorkspaceMirror::GetContextEpoch();
    TreeTarget target{itemData, expectedName};
    EnumChildWindows(mainWindow, FindProgramTree, reinterpret_cast<LPARAM>(&target));
    if (target.matches != 1) {
        outTrace += "|program_tree_target_missing_or_ambiguous";
        return false;
    }
    // 不抢焦点、不双击打开编辑器；选中后重新核对，不能将命令发给旧选中项。
    if (!TreeView_SelectItem(target.tree, target.item) ||
        TreeView_GetSelection(target.tree) != target.item || !MatchesTarget(target.tree, target.item, target)) {
        outTrace += "|program_tree_selection_changed";
        return false;
    }
    bool enabled = false;
    if (!ide.RunIsFuncEnabled(kDeleteSelectedProgramItem, enabled) || !enabled) {
        outTrace += "|program_tree_delete_disabled";
        return false;
    }
    if (WorkspaceMirror::GetContextEpoch() != contextEpoch ||
        WorkspaceMirror::GetGeneration() != snapshot.generation ||
        TreeView_GetSelection(target.tree) != target.item || !MatchesTarget(target.tree, target.item, target)) {
        outTrace += "|program_tree_context_changed";
        return false;
    }
    // 功能号由 IDE 分派，插件不读取对象内存，也不调用固定地址。可用性查询不等于版本认证。
    // 返回值只代表命令是否被处理；提交一次后始终走调用方的异步树/镜像验证，禁止回退重删。
    const bool reported = ide.RunFunction(kDeleteSelectedProgramItem);
    outTrace += "|native_delete_selected_program_item|command=0x82020017|reported_ok=" + std::to_string(reported);
    return true;
}
}
