#pragma once

#include <string>

namespace e571 {

// 按程序树标识和镜像路径删除普通程序集或类；使用镜像元数据保护窗口绑定，无固定内部地址。
// 返回值仅表示命令已提交，调用方须重新枚举程序树验证删除结果。
bool DeleteProgramUnitByTreeItemData(unsigned int itemData, const std::string& expectedName,
    const std::string& filePathUtf8, std::string& outFileType, std::string& outTrace);

}
