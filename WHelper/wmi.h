#pragma once

#include <comdef.h>
#include <Wbemidl.h>
#include <sstream>

#pragma comment(lib, "wbemuuid.lib")

namespace WHelper
{
	static constexpr int numberOfRequestObjects = 1;

	class Wmi
	{
		private:
			HRESULT _hres = 0;
			IWbemLocator* _iwl = nullptr;
			IWbemServices* _iws = nullptr;

		public:
		
			Wmi() noexcept;

			//TRUE = It's possible access to WMI service.
			[[nodiscard]]
			bool Init();

			//Sync query to CIM repo.
			[[nodiscard]]
			bool Query(IEnumWbemClassObject** enumerador, const char* query);

			//Clean pointers
			void Disconnect();
		};
}