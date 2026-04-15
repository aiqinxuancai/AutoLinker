#include <cstdlib>
#include <iostream>
#include <string>

#include "..\\src\\AutoLinkerTestApi.h"

namespace {

std::string JoinArguments(int startIndex, int argc, char* argv[])
{
	std::string text;
	for (int i = startIndex; i < argc; ++i) {
		if (!text.empty()) {
			text.push_back(' ');
		}
		text.append(argv[i]);
	}
	return text;
}

int PrintStringResult(const char* label, int result, const char* text)
{
	if (result >= 0) {
		std::cout << label << ": " << text << std::endl;
		return EXIT_SUCCESS;
	}

	if (result == AUTOLINKER_TEST_STRING_BUFFER_TOO_SMALL) {
		std::cerr << label << " failed: buffer too small" << std::endl;
	}
	else {
		std::cerr << label << " failed: invalid argument" << std::endl;
	}
	return EXIT_FAILURE;
}

int RunSmokeTest()
{
	char buffer[512] = {};
	int compareResult = 0;

	if (!AutoLinkerTest_CompareVersion("1.2.3", "1.2.0", &compareResult)) {
		std::cerr << "version-compare failed" << std::endl;
		return EXIT_FAILURE;
	}

	std::cout << "version-compare: " << compareResult << std::endl;

	int result = AutoLinkerTest_GetVersionText(buffer, static_cast<int>(sizeof(buffer)));
	if (PrintStringResult("version-text", result, buffer) != EXIT_SUCCESS) {
		return EXIT_FAILURE;
	}

	const char* linkerCommand = "link /out:\"D:\\demo\\AutoLinkerTest.exe\" \"D:\\demo\\main.obj\" \"D:\\deps\\static_lib\\krnln_static.lib\"";

	result = AutoLinkerTest_GetLinkerOutFileName(linkerCommand, buffer, static_cast<int>(sizeof(buffer)));
	if (PrintStringResult("linker-out", result, buffer) != EXIT_SUCCESS) {
		return EXIT_FAILURE;
	}

	result = AutoLinkerTest_GetLinkerKrnlnFileName(linkerCommand, buffer, static_cast<int>(sizeof(buffer)));
	if (PrintStringResult("linker-krnln", result, buffer) != EXIT_SUCCESS) {
		return EXIT_FAILURE;
	}

	result = AutoLinkerTest_ExtractBetweenDashes("before - middle - after", buffer, static_cast<int>(sizeof(buffer)));
	return PrintStringResult("extract-between-dashes", result, buffer);
}

int RunVersionCompare(const std::string& left, const std::string& right)
{
	int compareResult = 0;
	if (!AutoLinkerTest_CompareVersion(left.c_str(), right.c_str(), &compareResult)) {
		std::cerr << "version-compare failed" << std::endl;
		return EXIT_FAILURE;
	}

	std::cout << compareResult << std::endl;
	return EXIT_SUCCESS;
}

int RunStringCommand(const std::string& commandName, const std::string& input)
{
	char buffer[524288] = {};
	int result = AUTOLINKER_TEST_STRING_INVALID_ARGUMENT;

	if (commandName == "linker-out") {
		result = AutoLinkerTest_GetLinkerOutFileName(input.c_str(), buffer, static_cast<int>(sizeof(buffer)));
	}
	else if (commandName == "linker-krnln") {
		result = AutoLinkerTest_GetLinkerKrnlnFileName(input.c_str(), buffer, static_cast<int>(sizeof(buffer)));
	}
	else if (commandName == "between-dashes") {
		result = AutoLinkerTest_ExtractBetweenDashes(input.c_str(), buffer, static_cast<int>(sizeof(buffer)));
	}
	else if (commandName == "version-text") {
		result = AutoLinkerTest_GetVersionText(buffer, static_cast<int>(sizeof(buffer)));
	}
	else if (commandName == "module-local-dump") {
		result = AutoLinkerTest_DumpLocalModulePublicInfo(input.c_str(), buffer, static_cast<int>(sizeof(buffer)));
	}

	if (result < 0) {
		return PrintStringResult(commandName.c_str(), result, buffer);
	}

	std::cout << buffer << std::endl;
	return EXIT_SUCCESS;
}

void PrintUsage()
{
	std::cout << "AutoLinkerTest commands:" << std::endl;
	std::cout << "  AutoLinkerTest" << std::endl;
	std::cout << "  AutoLinkerTest version-compare <left> <right>" << std::endl;
	std::cout << "  AutoLinkerTest linker-out <link-command>" << std::endl;
	std::cout << "  AutoLinkerTest linker-krnln <link-command>" << std::endl;
	std::cout << "  AutoLinkerTest between-dashes <text>" << std::endl;
	std::cout << "  AutoLinkerTest version-text" << std::endl;
	std::cout << "  AutoLinkerTest module-local-dump <module-path>" << std::endl;
	std::cout << "  AutoLinkerTest e2txt-generate <input-path> <output-path>" << std::endl;
	std::cout << "  AutoLinkerTest e2txt-restore <input-path> <output-path>" << std::endl;
	std::cout << "  AutoLinkerTest bundle-digest-compare <input.e> <input-dir>" << std::endl;
}

}

int main(int argc, char* argv[])
{
	if (argc == 1) {
		return RunSmokeTest();
	}

	const std::string commandName = argv[1];
	if (commandName == "version-compare") {
		if (argc != 4) {
			PrintUsage();
			return EXIT_FAILURE;
		}
		return RunVersionCompare(argv[2], argv[3]);
	}

	if (commandName == "linker-out" || commandName == "linker-krnln" || commandName == "between-dashes" || commandName == "module-local-dump") {
		if (argc < 3) {
			PrintUsage();
			return EXIT_FAILURE;
		}
		return RunStringCommand(commandName, JoinArguments(2, argc, argv));
	}

	if (commandName == "version-text") {
		if (argc != 2) {
			PrintUsage();
			return EXIT_FAILURE;
		}
		return RunStringCommand(commandName, "");
	}

	if (commandName == "e2txt-generate") {
		if (argc != 4) {
			PrintUsage();
			return EXIT_FAILURE;
		}

		char buffer[524288] = {};
		const int result = AutoLinkerTest_GenerateE2Txt(argv[2], argv[3], buffer, static_cast<int>(sizeof(buffer)));
		if (result < 0) {
			return PrintStringResult(commandName.c_str(), result, buffer);
		}
		std::cout << buffer << std::endl;
		return EXIT_SUCCESS;
	}

	if (commandName == "e2txt-restore") {
		if (argc != 4) {
			PrintUsage();
			return EXIT_FAILURE;
		}

		char buffer[524288] = {};
		const int result = AutoLinkerTest_RestoreE2Txt(argv[2], argv[3], buffer, static_cast<int>(sizeof(buffer)));
		if (result < 0) {
			return PrintStringResult(commandName.c_str(), result, buffer);
		}
		std::cout << buffer << std::endl;
		return EXIT_SUCCESS;
	}

	if (commandName == "bundle-digest-compare") {
		if (argc != 4) {
			PrintUsage();
			return EXIT_FAILURE;
		}

		char buffer[524288] = {};
		const int result = AutoLinkerTest_CompareBundleDigest(argv[2], argv[3], buffer, static_cast<int>(sizeof(buffer)));
		if (result < 0) {
			return PrintStringResult(commandName.c_str(), result, buffer);
		}
		std::cout << buffer << std::endl;
		return EXIT_SUCCESS;
	}

	PrintUsage();
	return EXIT_FAILURE;
}
