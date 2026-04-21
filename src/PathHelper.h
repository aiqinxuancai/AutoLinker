#pragma once
#include <filesystem>
#include <optional>
#include <string>
#include <vector>
#include <Windows.h>
// 鑾峰彇褰撳墠涓荤▼搴忔墍鍦ㄧ洰褰曘€?
std::string GetBasePath();
// 鑾峰彇 AutoLinker 宸ヤ綔鐩綍璺緞銆?
std::filesystem::path GetAutoLinkerDirectoryPath();
// 鑾峰彇 AutoLinker 鏃ュ織鐩綍璺緞锛屽涓嶅瓨鍦ㄤ細鑷姩鍒涘缓銆?
std::filesystem::path GetAutoLinkerLogDirectoryPath();
// 鑾峰彇 AutoLinker 鏃ュ織鏂囦欢瀹屾暣璺緞銆?
std::filesystem::path GetAutoLinkerLogFilePath(const std::string& fileName);
// 鎻愬彇涓や釜鐭í绾夸箣闂寸殑鏂囨湰銆?
std::string ExtractBetweenDashes(const std::string& text);
/// <summary>
/// 鏌ユ壘瀛楄妭搴忓垪鍦ㄦ枃浠朵腑鐨勫亸绉汇€?
/// </summary>
/// <param name="filename"></param>
/// <param name="search_bytes"></param>
/// <returns></returns>
std::optional<size_t> FindByteInFile(const std::string& filename, const std::vector<char>& search_bytes);
/// <summary>
/// 鑾峰彇 /out: 瀵瑰簲鐨勮緭鍑烘枃浠跺悕銆?
/// </summary>
/// <param name="s"></param>
/// <returns></returns>
std::string GetLinkerCommandOutFileName(const std::string& s);
// 鑾峰彇鍛戒护琛屼腑鐨?krnln 闈欐€佸簱璺緞銆?
std::string GetLinkerCommandKrnlnFileName(const std::string& s);