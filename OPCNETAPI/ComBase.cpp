#include "stdafx.h"
#include "ComBase.h"


ComBase::ComBase()
{
	server = NULL;

	m_pHost		= System::String::Empty;
	m_pUser		= System::String::Empty;
	m_pPassword	= System::String::Empty;
	m_pDomain		= System::String::Empty;
}

HRESULT ComBase::GetInterface(
	IID iid,
	void** ppOut)
{
	HRESULT hr = E_FAIL;

	if (server == NULL)
		return E_NOT_SET;
	
	if (isRemote)
	{
		LPWSTR pHost = (LPWSTR)System::Runtime::InteropServices::Marshal::StringToCoTaskMemUni(m_pHost).ToPointer();
		char* pUser = (char*)System::Runtime::InteropServices::Marshal::StringToCoTaskMemAnsi(m_pUser).ToPointer();
		char* pPassword = (char*)System::Runtime::InteropServices::Marshal::StringToCoTaskMemAnsi(m_pPassword).ToPointer();
		char* pDomain = (char*)System::Runtime::InteropServices::Marshal::StringToCoTaskMemAnsi(m_pDomain).ToPointer();

		hr = QueryInterfaceEx(
			server,
			pHost,
			pUser,
			pPassword,
			pDomain,
			iid,
			ppOut);

		System::Runtime::InteropServices::Marshal::ZeroFreeCoTaskMemUnicode((System::IntPtr) pHost);
		System::Runtime::InteropServices::Marshal::ZeroFreeCoTaskMemAnsi((System::IntPtr) pUser);
		System::Runtime::InteropServices::Marshal::ZeroFreeCoTaskMemAnsi((System::IntPtr) pPassword);
		System::Runtime::InteropServices::Marshal::ZeroFreeCoTaskMemAnsi((System::IntPtr) pDomain);
	}
	else
	{
		hr = QueryInterface(
			server,
			iid,
			ppOut);
	}

	return hr;
}

HRESULT ComBase::SetProxyBlanket(
	IUnknown* ppIn)
{
	if (server == NULL)
		return E_NOT_SET;

	if (!this->isRemote)
		return S_OK;

	LPWSTR pHost = (LPWSTR)System::Runtime::InteropServices::Marshal::StringToCoTaskMemUni(m_pHost).ToPointer();
	char* pUser = (char*)System::Runtime::InteropServices::Marshal::StringToCoTaskMemAnsi(m_pUser).ToPointer();
	char* pPassword = (char*)System::Runtime::InteropServices::Marshal::StringToCoTaskMemAnsi(m_pPassword).ToPointer();
	char* pDomain = (char*)System::Runtime::InteropServices::Marshal::StringToCoTaskMemAnsi(m_pDomain).ToPointer();

	HRESULT hr = CoSetProxyBlanket(
		ppIn,
		pHost,
		pUser,
		pPassword,
		pDomain);

	System::Runtime::InteropServices::Marshal::ZeroFreeCoTaskMemUnicode((System::IntPtr) pHost);
	System::Runtime::InteropServices::Marshal::ZeroFreeCoTaskMemAnsi((System::IntPtr) pUser);
	System::Runtime::InteropServices::Marshal::ZeroFreeCoTaskMemAnsi((System::IntPtr) pPassword);
	System::Runtime::InteropServices::Marshal::ZeroFreeCoTaskMemAnsi((System::IntPtr) pDomain);

	return hr;
}

ComBase::ComBase(
	ComBase^ ppParent,
	IUnknown* server)
{
	m_pHost = ppParent->m_pHost;
	m_pUser = ppParent->m_pUser;
	m_pPassword = ppParent->m_pPassword;
	m_pDomain = ppParent->m_pDomain;

	this->isRemote = ppParent->isRemote;

	this->server = server;
}

ComBase::~ComBase()
{
	this->Release();
}

HRESULT ComBase::Connect(System::Guid ppClsid, System::Guid ppIID)
{
	if (server != NULL)
	{
		this->Release();
	}		

	isRemote = false;

	const CLSID clsid = GuidToCLSID(ppClsid);
	const IID iid = (const IID)GuidToCLSID(ppIID);

	void* instance = static_cast<void*>(server);

	HRESULT hr = CreateInstance(
		clsid,
		iid,
		&instance);

	server = (IUnknown*)instance;

	return hr;
}

HRESULT ComBase::Connect(System::Guid ppClsid, System::String ^ inHost, System::String ^ inUser, System::String ^ inPassword, System::String ^ inDomain, System::Guid ppIID)
{
	if (server != NULL)
	{
		this->Release();
	}

	isRemote = true;

	m_pHost = inHost;
	m_pUser = inUser;
	m_pPassword = inPassword;
	m_pDomain = inDomain;

	LPWSTR pHost = (LPWSTR)System::Runtime::InteropServices::Marshal::StringToCoTaskMemUni(m_pHost).ToPointer();
	char* pUser = (char*)System::Runtime::InteropServices::Marshal::StringToCoTaskMemAnsi(m_pUser).ToPointer();
	char* pPassword = (char*)System::Runtime::InteropServices::Marshal::StringToCoTaskMemAnsi(m_pPassword).ToPointer();
	char* pDomain = (char*)System::Runtime::InteropServices::Marshal::StringToCoTaskMemAnsi(m_pDomain).ToPointer();

	const CLSID clsid = GuidToCLSID(ppClsid);
	const IID iid = (const IID)GuidToCLSID(ppIID);

	void* instance = static_cast<void*>(server);

	HRESULT hr = CreateInstanceEx(
		clsid,
		pHost,
		pUser,
		pPassword,
		pDomain,
		iid,
		&instance);

	server = (IUnknown*)instance;

	System::Runtime::InteropServices::Marshal::ZeroFreeCoTaskMemUnicode((System::IntPtr) pHost);
	System::Runtime::InteropServices::Marshal::ZeroFreeCoTaskMemAnsi((System::IntPtr) pUser);
	System::Runtime::InteropServices::Marshal::ZeroFreeCoTaskMemAnsi((System::IntPtr) pPassword);
	System::Runtime::InteropServices::Marshal::ZeroFreeCoTaskMemAnsi((System::IntPtr) pDomain);

	return hr;
}

void ComBase::Release()
{
	HRESULT hr;

	m_pHost		= nullptr;
	m_pUser		= nullptr;
	m_pPassword	= nullptr;
	m_pDomain		= nullptr;

	if (server != NULL)
	{
		hr = server->Release();
		server = NULL;
	}
}
