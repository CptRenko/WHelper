#include "wmi.h"

namespace WHelper
{
	Wmi::Wmi() noexcept
	{
		this->_iwl = nullptr;
		this->_iws = nullptr;
		this->_hres = 0;
	}

	bool Wmi::Init()
	{
	
	//If this it's being compiler as .EXE , it's necesary init COM service.
	//On DLL/Dynamic library, it's responsability of the main app.
	#ifndef _WINDLL
		
		//Init COM service.
		this->_hres = CoInitializeEx(nullptr, COINIT_MULTITHREADED);

		if (FAILED(this->_hres))
		{
			return false;
		}

		//Permissions...
		this->_hres = CoInitializeSecurity(
			nullptr, //Not a server.
			-1, 
			NULL, 
			NULL, 
			RPC_C_AUTHN_LEVEL_DEFAULT, //Default level.
			RPC_C_IMP_LEVEL_IMPERSONATE, 
			NULL, //Local access
			EOAC_NONE, 
			NULL 
		);
	
		if (FAILED(this->_hres))
		{
			return false;
		} 
	#endif

		//WMI Pointer
		this->_hres = CoCreateInstance(CLSID_WbemLocator, nullptr, CLSCTX_INPROC_SERVER, IID_IWbemLocator, (LPVOID*)&this->_iwl);
		
		if (FAILED(this->_hres))
		{
			#ifndef _WINDLL
					CoUninitialize();
			#endif
			return false;
		}

		//Start IWbemServices proxy, required for WMI access.
		this->_hres = this->_iwl->ConnectServer(
			_bstr_t(L"ROOT\\CIMV2"), 
			nullptr, 
			nullptr, 
			nullptr, //Local
			NULL,
			0, 
			0, 
			&this->_iws //Puntero
		);

		if (FAILED(this->_hres))
		{
			this->_iwl->Release();
	#ifndef _WINDLL
			CoUninitialize();
	#endif
			return false;
		}

		//Grant permissions to IWbemServices
		this->_hres = CoSetProxyBlanket(
			this->_iws, 
			RPC_C_AUTHN_WINNT, //Microsoft NT Lan Manager (NTLM) SSP
			RPC_C_AUTHZ_NONE, //MNLMSSP
			nullptr, 
			RPC_C_AUTHN_LEVEL_CALL, 
			RPC_C_IMP_LEVEL_IMPERSONATE,
			NULL,
			EOAC_NONE
		);

		if (FAILED(this->_hres))
		{
			this->_iws->Release();
			this->_iwl->Release();
	#ifndef _WINDLL
			CoUninitialize();
	#endif
			return false;
		}

		return true;
	}

	bool Wmi::Query(IEnumWbemClassObject** enumerador, const char* query)
	{
		if(this->_iws != nullptr && enumerador != nullptr)
		{
			this->_hres = this->_iws->ExecQuery(
				bstr_t("WQL"), //Windows Management Instrumentation Query Language (WQL)
				bstr_t(query),
				WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, 
				nullptr,
				enumerador
			);

			if (FAILED(this->_hres))
			{
				return false;
			}

			return true;
		}
		else
		{
			return false;
		}
	}

	void Wmi::Disconnect()
	{
		if (this->_iws != nullptr)
		{
			this->_iws->Release();
		}

		if (this->_iwl != nullptr)
		{
			this->_iwl->Release();
		}

	//Si la librería levanto COM, es deber de nosotros eliminar la instancia.
	#ifndef _WINDLL
		CoUninitialize();
	#endif

	}
 }