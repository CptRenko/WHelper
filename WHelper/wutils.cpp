#include "wutils.h"

namespace WHelper {

	int CastSizetToInt(const size_t size) noexcept
	{
		if (size > static_cast<size_t>((std::numeric_limits<int>::max)()))
		{
			return -1;
		}

		return static_cast<int>(size);
	}

	std::string BSTRToString(const BSTR bstr) noexcept
	{
		int wslen = SysStringLen(bstr);
		return ConvertWStringToString(bstr, wslen);
	}

	std::string ConvertWStringToString(const wchar_t* pstr, const size_t wslen)
	{
		std::string str;
		const int size = CastSizetToInt(wslen);
		if (size > -1)
		{
			const int bytesLength = WideCharToMultiByte(CP_UTF8, 0, pstr, size, nullptr, 0, nullptr, nullptr);
			str.resize(bytesLength);
			WideCharToMultiByte(CP_UTF8, 0, pstr, size, &str.at(0), bytesLength, nullptr, nullptr);
		}

		return str;
	}

	std::wstring ConvertStringToWstring(const char* str, const size_t strlen) noexcept
	{
		std::wstring wstr;
		const int size = CastSizetToInt(strlen);
		if (size > -1)
		{
			const int bytesLength = MultiByteToWideChar(CP_UTF8, 0, str, size, nullptr, NULL);
			wstr.resize(bytesLength);
			MultiByteToWideChar(CP_UTF8, 0, str, size, &wstr.at(0), bytesLength);
		}

		return wstr;
	}

	std::string ConvertComErrorToString(const HRESULT result) noexcept
	{
		_com_error error(result);
		std::string str = ConvertWStringToString(error.ErrorMessage(), wcslen(error.ErrorMessage()));
		return str;
	}

	std::string ConvertWindowsCoreErrorToString(const DWORD codeError) noexcept
	{
		wchar_t msg[256];

		FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
			nullptr, codeError, LANG_SPANISH, msg, sizeof(msg), nullptr);

		std::string message = ConvertWStringToString(msg, wcslen(msg));
		return message;
	}

	std::string GetWindowsVersion(Wmi* const w)
	{
		std::string version;

		if (w != nullptr)
		{
			IEnumWbemClassObject* enumerador = nullptr;
			IWbemClassObject* data = nullptr;

			ULONG ureturn = 0;
			HRESULT hres;

			bool result = w->Query(&enumerador, "SELECT Caption, OSArchitecture, Version  FROM Win32_OperatingSystem");

			if (result)
			{
				std::array<VARIANT, 3> VData;

				//Set VT_EMPTY
				for (size_t i = 0; i < VData.size(); i++)
				{
					VariantInit(&VData.at(i));
				}

				while (enumerador)
				{
					hres = enumerador->Next(WBEM_INFINITE, 1, &data, &ureturn);

					if (SUCCEEDED(hres))
					{
						if (ureturn == 0)
						{
							break;
						}

						data->Get(L"Caption", 0, &VData.at(0), 0, 0);
						data->Get(L"OSArchitecture", 0, &VData.at(1), 0, 0);
						data->Get(L"Version", 0, &VData.at(2), 0, 0);
						data->Release();

						break;
					}
					else
					{
						continue;
					}
				}

				enumerador->Release();

				if (ureturn > 0)
				{
					std::array<std::string, 3> aData;

					_bstr_t strConcat;

					for (size_t i = 0; i < VData.size(); i++)
					{
						_bstr_t value(VData.at(i).bstrVal, true);
						strConcat += value + " ";

						VariantClear(&VData.at(i));
					}

					version = BSTRToString(strConcat);
				}
			}
		}

		return version;
	}
}