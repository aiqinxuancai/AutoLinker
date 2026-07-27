#include "AIChatUserInputRequest.h"

#include <algorithm>
#include <cctype>
#include <initializer_list>
#include <unordered_map>
#include <unordered_set>

#include "..\\thirdparty\\json.hpp"

namespace {
constexpr size_t kMaxQuestionIdBytes = 64;
constexpr size_t kMaxHeaderCharacters = 12;
constexpr size_t kMaxQuestionCharacters = 500;
constexpr size_t kMaxLabelCharacters = 80;
constexpr size_t kMaxDescriptionCharacters = 300;
constexpr size_t kMaxNoteCharacters = 2000;

std::string TrimAsciiCopy(const std::string& text)
{
	size_t begin = 0;
	while (begin < text.size() && std::isspace(static_cast<unsigned char>(text[begin])) != 0) {
		++begin;
	}
	size_t end = text.size();
	while (end > begin && std::isspace(static_cast<unsigned char>(text[end - 1])) != 0) {
		--end;
	}
	return text.substr(begin, end - begin);
}

size_t CountUtf8Characters(const std::string& text)
{
	size_t count = 0;
	for (const unsigned char ch : text) {
		if ((ch & 0xC0) != 0x80) {
			++count;
		}
	}
	return count;
}

bool IsSnakeCaseId(const std::string& text)
{
	if (text.empty() || text.size() > kMaxQuestionIdBytes ||
		text.front() < 'a' || text.front() > 'z') {
		return false;
	}
	for (const unsigned char ch : text) {
		if ((ch < 'a' || ch > 'z') && (ch < '0' || ch > '9') && ch != '_') {
			return false;
		}
	}
	return true;
}

bool HasOnlyKeys(
	const nlohmann::json& value,
	std::initializer_list<const char*> allowedKeys)
{
	if (!value.is_object()) {
		return false;
	}
	for (auto it = value.begin(); it != value.end(); ++it) {
		const bool allowed = std::any_of(
			allowedKeys.begin(),
			allowedKeys.end(),
			[&it](const char* key) { return it.key() == key; });
		if (!allowed) {
			return false;
		}
	}
	return true;
}

bool ReadRequiredString(
	const nlohmann::json& value,
	const char* key,
	size_t maxCharacters,
	std::string& out,
	std::string& outError)
{
	out.clear();
	if (!value.is_object() || !value.contains(key) || !value[key].is_string()) {
		outError = std::string(key) + " must be a string";
		return false;
	}
	out = TrimAsciiCopy(value[key].get<std::string>());
	if (out.empty()) {
		outError = std::string(key) + " must not be empty";
		return false;
	}
	if (CountUtf8Characters(out) > maxCharacters) {
		outError = std::string(key) + " is too long";
		return false;
	}
	return true;
}

const AIChatUserInputQuestion* FindQuestion(
	const std::vector<AIChatUserInputQuestion>& questions,
	const std::string& id)
{
	const auto it = std::find_if(
		questions.begin(),
		questions.end(),
		[&id](const AIChatUserInputQuestion& question) { return question.id == id; });
	return it == questions.end() ? nullptr : &*it;
}
} // namespace

bool ParseAIChatUserInputRequestArguments(
	const std::string& argumentsJsonUtf8,
	std::vector<AIChatUserInputQuestion>& outQuestions,
	std::string& outError)
{
	outQuestions.clear();
	outError.clear();

	nlohmann::json root;
	try {
		root = nlohmann::json::parse(argumentsJsonUtf8);
	}
	catch (const std::exception& ex) {
		outError = std::string("invalid arguments json: ") + ex.what();
		return false;
	}
	if (!root.is_object() || !root.contains("questions") || !root["questions"].is_array()) {
		outError = "questions must be an array";
		return false;
	}
	if (!HasOnlyKeys(root, {"questions"})) {
		outError = "request_user_input contains an unsupported property";
		return false;
	}
	const auto& rows = root["questions"];
	if (rows.empty() || rows.size() > 3) {
		outError = "questions must contain between 1 and 3 items";
		return false;
	}

	std::unordered_set<std::string> questionIds;
	for (const auto& row : rows) {
		if (!HasOnlyKeys(row, {"id", "header", "question", "options"})) {
			outError = "question contains an unsupported property";
			return false;
		}
		AIChatUserInputQuestion question;
		if (!ReadRequiredString(row, "id", kMaxQuestionIdBytes, question.id, outError)) {
			return false;
		}
		if (!IsSnakeCaseId(question.id)) {
			outError = "question id must use snake_case and start with a lowercase letter";
			return false;
		}
		if (!questionIds.insert(question.id).second) {
			outError = "question ids must be unique";
			return false;
		}
		if (!ReadRequiredString(row, "header", kMaxHeaderCharacters, question.headerUtf8, outError) ||
			!ReadRequiredString(row, "question", kMaxQuestionCharacters, question.questionUtf8, outError)) {
			return false;
		}
		if (!row.contains("options") || !row["options"].is_array() ||
			row["options"].size() < 2 || row["options"].size() > 3) {
			outError = "each question must contain between 2 and 3 options";
			return false;
		}

		std::unordered_set<std::string> labels;
		for (const auto& optionValue : row["options"]) {
			if (!HasOnlyKeys(optionValue, {"label", "description"})) {
				outError = "option contains an unsupported property";
				return false;
			}
			AIChatUserInputOption option;
			if (!ReadRequiredString(optionValue, "label", kMaxLabelCharacters, option.labelUtf8, outError) ||
				!ReadRequiredString(optionValue, "description", kMaxDescriptionCharacters, option.descriptionUtf8, outError)) {
				return false;
			}
			if (!labels.insert(option.labelUtf8).second) {
				outError = "option labels must be unique within each question";
				return false;
			}
			question.options.push_back(std::move(option));
		}
		outQuestions.push_back(std::move(question));
	}
	return true;
}

bool BuildAIChatUserInputResponseJson(
	const std::vector<AIChatUserInputQuestion>& questions,
	const std::vector<AIChatUserInputAnswer>& answers,
	std::string& outResponseJsonUtf8,
	std::string& outError)
{
	outResponseJsonUtf8.clear();
	outError.clear();
	if (questions.empty() || answers.size() != questions.size()) {
		outError = "every question must be answered";
		return false;
	}

	std::unordered_map<std::string, const AIChatUserInputAnswer*> answersById;
	for (const auto& answer : answers) {
		if (FindQuestion(questions, answer.questionId) == nullptr) {
			outError = "answer contains an unknown question id";
			return false;
		}
		if (!answersById.emplace(answer.questionId, &answer).second) {
			outError = "answer question ids must be unique";
			return false;
		}
	}

	nlohmann::json response;
	response["answers"] = nlohmann::json::object();
	for (const auto& question : questions) {
		const auto answerIt = answersById.find(question.id);
		if (answerIt == answersById.end()) {
			outError = "every question must be answered";
			return false;
		}
		const AIChatUserInputAnswer& answer = *answerIt->second;
		const std::string label = TrimAsciiCopy(answer.selectedLabelUtf8);
		const std::string note = TrimAsciiCopy(answer.noteUtf8);
		if (CountUtf8Characters(note) > kMaxNoteCharacters) {
			outError = "answer note is too long";
			return false;
		}

		nlohmann::json values = nlohmann::json::array();
		if (!label.empty()) {
			const bool knownLabel = std::any_of(
				question.options.begin(),
				question.options.end(),
				[&label](const AIChatUserInputOption& option) { return option.labelUtf8 == label; });
			if (!knownLabel) {
				outError = "answer contains an unknown option label";
				return false;
			}
			values.push_back(label);
		}
		else if (note.empty()) {
			outError = "other answers require a note";
			return false;
		}
		if (!note.empty()) {
			values.push_back("user_note: " + note);
		}
		response["answers"][question.id] = {{"answers", std::move(values)}};
	}

	outResponseJsonUtf8 = response.dump(-1, ' ', false, nlohmann::json::error_handler_t::replace);
	return true;
}

std::string BuildAIChatUserInputRequestSelfTestJson()
{
	const std::string validArguments = R"json({"questions":[{"id":"target_scope","header":"范围","question":"请选择实现范围？","options":[{"label":"完整支持 (Recommended)","description":"实现完整流程。"},{"label":"最小支持","description":"只实现基础选择。"}]}]})json";
	std::vector<AIChatUserInputQuestion> questions;
	std::string error;
	const bool validParsed = ParseAIChatUserInputRequestArguments(validArguments, questions, error) &&
		questions.size() == 1 && questions.front().options.size() == 2;

	std::string response;
	const bool responseBuilt = validParsed && BuildAIChatUserInputResponseJson(
		questions,
		{{"target_scope", "完整支持 (Recommended)", "保留历史记录"}},
		response,
		error) &&
		response.find("user_note: 保留历史记录") != std::string::npos;

	std::vector<AIChatUserInputQuestion> ignored;
	std::string invalidError;
	const bool duplicateRejected = !ParseAIChatUserInputRequestArguments(
		R"({"questions":[{"id":"same","header":"A","question":"A?","options":[{"label":"1","description":"a"},{"label":"2","description":"b"}]},{"id":"same","header":"B","question":"B?","options":[{"label":"1","description":"a"},{"label":"2","description":"b"}]}]})",
		ignored,
		invalidError);
	const bool additionalPropertyRejected = !ParseAIChatUserInputRequestArguments(
		R"({"questions":[{"id":"scope","header":"A","question":"A?","options":[{"label":"1","description":"a"},{"label":"2","description":"b"}]}],"autoResolutionMs":60000})",
		ignored,
		invalidError);
	const bool unknownOptionRejected = validParsed && !BuildAIChatUserInputResponseJson(
		questions,
		{{"target_scope", "unknown", ""}},
		response,
		invalidError);
	const bool emptyOtherRejected = validParsed && !BuildAIChatUserInputResponseJson(
		questions,
		{{"target_scope", "", ""}},
		response,
		invalidError);

	return nlohmann::json({
		{"name", "plan-user-input-protocol"},
		{"ok", validParsed && responseBuilt && duplicateRejected && additionalPropertyRejected && unknownOptionRejected && emptyOtherRejected},
		{"valid_arguments", validParsed},
		{"response_shape", responseBuilt},
		{"duplicate_id_rejected", duplicateRejected},
		{"additional_property_rejected", additionalPropertyRejected},
		{"unknown_option_rejected", unknownOptionRejected},
		{"empty_other_rejected", emptyOtherRejected}
	}).dump();
}
