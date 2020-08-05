#pragma once

#include <string>
#include <strsafe.h>
#include <comdef.h>
#include "wmi.h"
#include <array>

namespace WHelper
{
	[[nodiscard]]
	int CastSizetToInt(const size_t size) noexcept;

	[[nodiscard]]
	std::string BSTRToString(const BSTR bstr) noexcept;

	[[nodiscard]]
	std::string ConvertWStringToString(const wchar_t* pstr, const size_t wslen);

	[[nodiscard]]
	std::wstring ConvertStringToWstring(const char* str, const size_t strlen) noexcept;

	[[nodiscard]]
	std::string ConvertComErrorToString(const HRESULT result) noexcept;

	[[nodiscard]]
	std::string ConvertWindowsCoreErrorToString(const DWORD codeError) noexcept;

	[[nodiscard]]
	std::string GetWindowsVersion(Wmi* const w);
}