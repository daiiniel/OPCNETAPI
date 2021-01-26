#pragma once

#include <Windows.h>

#include "ComUtils.h"


public ref class ComBase
{

protected:

	IUnknown* server;
	
	bool isRemote;

	System::String^ m_pHost;
	System::String^ m_pUser;
	System::String^ m_pPassword;
	System::String^ m_pDomain;

	ComBase();

internal:

	HRESULT GetInterface(
		IID iid,
		void** ppOut);

	HRESULT SetProxyBlanket(
		IUnknown* ppIn);

	ComBase(
		ComBase^ ppParent,
		IUnknown* server);

	~ComBase();

	HRESULT Connect(
		System::Guid ppClsid,
		System::Guid ppIID);

	HRESULT Connect(
		System::Guid ppClsid,
		System::String^ inHost,
		System::String^ inUser,
		System::String^ inPassword,
		System::String^ inDomain,
		System::Guid ppIID);

	void Release();

};

