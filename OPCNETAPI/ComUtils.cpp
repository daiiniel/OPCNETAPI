#include "stdafx.h"
#include "ComUtils.h"

HRESULT CreateInstance(
	const CLSID ppClsid, 
	const IID iid, 
	void ** ppUnknown)
{
	HRESULT hr = CoInitializeEx(
		NULL,
		COINIT_MULTITHREADED);

	/*

	This function can return the standard return values E_INVALIDARG, E_OUTOFMEMORY, and E_UNEXPECTED, as well as the following values.

	S_OK					The COM library was initialized successfully on this thread.
	S_FALSE					The COM library is already initialized on this thread.
	RPC_E_CHANGED_MODE		A previous call to CoInitializeEx specified the concurrency model for this thread as multithread apartment (MTA). This could also indicate that a change from neutral-threaded apartment to single-threaded apartment has occurred.

	*/

	if (hr != S_OK && hr != S_FALSE && hr != RPC_E_CHANGED_MODE)
	{
		throw gcnew System::ComponentModel::Win32Exception(hr);
	}

	return CoCreateInstance(
		ppClsid,
		NULL,
		CLSCTX_LOCAL_SERVER,
		iid,
		(void**)ppUnknown);
}

HRESULT CreateInstanceEx(
	const CLSID ppClsid, 
	const LPWSTR pHost, 
	const char * pUser, 
	const char * pPassword, 
	const char * pDomain, 
	const IID iid, 
	void ** ppUnknown)
{
	HRESULT hr;

	hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);

	/*

	This function can return the standard return values E_INVALIDARG, E_OUTOFMEMORY, and E_UNEXPECTED, as well as the following values.

	S_OK					The COM library was initialized successfully on this thread.
	S_FALSE					The COM library is already initialized on this thread.
	RPC_E_CHANGED_MODE		A previous call to CoInitializeEx specified the concurrency model for this thread as multithread apartment (MTA). This could also indicate that a change from neutral-threaded apartment to single-threaded apartment has occurred.

	*/

	if (hr != S_OK && hr != S_FALSE && hr != RPC_E_CHANGED_MODE)
	{
		throw gcnew System::ComponentModel::Win32Exception(hr);
	}

	COAUTHIDENTITY id;

	id.User = (USHORT*)pUser;
	id.UserLength = (pUser != NULL) ? (DWORD)strlen(pUser) : 0UL;

	id.Domain = (USHORT*)pDomain;
	id.DomainLength = (pDomain != NULL) ? (DWORD)strlen(pDomain) : 0UL;

	id.Password = (USHORT*)pPassword;
	id.PasswordLength = (pPassword != NULL) ? (DWORD)strlen(pPassword) : 0UL;

	id.Flags = SEC_WINNT_AUTH_IDENTITY_ANSI;

	COAUTHINFO coin;

	coin.dwAuthnSvc = RPC_C_AUTHN_WINNT;
	coin.dwAuthzSvc = RPC_C_AUTHN_DEFAULT;
	coin.pwszServerPrincName = NULL;
	coin.dwAuthnLevel = RPC_C_AUTHN_LEVEL_CALL;
	coin.dwImpersonationLevel = RPC_C_IMP_LEVEL_IMPERSONATE;
	coin.pAuthIdentityData = &id;
	coin.dwCapabilities = EOAC_NONE;

	COSERVERINFO sin;

	sin.dwReserved1 = 0;
	sin.dwReserved2 = 0;
	sin.pwszName = pHost;
	sin.pAuthInfo = &coin;

	tagMULTI_QI res;

	res.pIID = &iid;
	res.pItf = NULL;

	hr = CoCreateInstanceEx(
		ppClsid,
		NULL,
		CLSCTX_REMOTE_SERVER,
		&sin,
		1,
		&res);

	if (hr != S_OK)
	{
		if (res.pItf != NULL)
		{
			res.pItf->Release();
		}

		return hr;
	}

	*ppUnknown = res.pItf;

	HRESULT hrProxy = CoSetProxyBlanket(
		(IUnknown*)*ppUnknown,
		pHost,
		pUser,
		pPassword,
		pDomain);

	return res.hr;
}

HRESULT QueryInterface(
	IUnknown * ppInUnknown, 
	const IID iid, 
	void ** ppOutUnknown)
{
	*ppOutUnknown = NULL;

	if (ppInUnknown == NULL)
		return E_NOT_SET;

	HRESULT hr = ppInUnknown->QueryInterface(
		iid,
		(void**)ppOutUnknown);

	if (hr != S_OK)
	{
		if (ppOutUnknown != NULL && *ppOutUnknown != NULL)
		{
			((IUnknown*)(*ppOutUnknown))->Release();
			(*ppOutUnknown) = NULL;
		}

		return hr;
	}

	return S_OK;
}

HRESULT QueryInterfaceEx(
	IUnknown * ppInUnknown, 
	const LPWSTR pHost, 
	const char * pUser, 
	const char * pPassword, 
	const char * pDomain, 
	const IID iid, 
	void ** ppOutUnknown)
{
	HRESULT hr = QueryInterface(
		ppInUnknown,
		iid,
		ppOutUnknown);

	if (hr != S_OK)
		return hr;

	HRESULT hrProxy = CoSetProxyBlanket(
		(IUnknown*)*ppOutUnknown,
		pHost,
		pUser,
		pPassword,
		pDomain);
	
	return S_OK;
}

HRESULT CoSetProxyBlanket(
	IUnknown* ppInUnknown,
	const LPWSTR pHost,
	const char * pUser,
	const char * pPassword,
	const char * pDomain)
{
	SEC_WINNT_AUTH_IDENTITY pAuthInfo;

	pAuthInfo.User = (USHORT*)pUser;
	pAuthInfo.UserLength = (pUser != NULL) ? (DWORD)strlen(pUser) : 0UL;

	pAuthInfo.Domain = (USHORT*)pDomain;
	pAuthInfo.DomainLength = (pDomain != NULL) ? (DWORD)strlen(pDomain) : 0UL;

	pAuthInfo.Password = (USHORT*)pPassword;
	pAuthInfo.PasswordLength = (pPassword != NULL) ? (DWORD)strlen(pPassword) : 0UL;

	pAuthInfo.Flags = SEC_WINNT_AUTH_IDENTITY_ANSI;

	return CoSetProxyBlanket(
		ppInUnknown,
		RPC_C_AUTHN_WINNT,
		RPC_C_AUTHN_DEFAULT,
		NULL,
		RPC_C_AUTHN_LEVEL_CALL,
		RPC_C_IMP_LEVEL_IMPERSONATE,
		&pAuthInfo,
		EOAC_NONE);
}

CLSID GuidToCLSID(System::Guid guid)
{
	CLSID res;

	array<unsigned char>^ guidData = guid.ToByteArray();

	res.Data1 = 0UL;
	res.Data1 = guidData[3] << 8 * 3 | guidData[2] << 8 * 2 | guidData[1] << 8 | guidData[0];

	res.Data2 = 0;
	res.Data2 = guidData[5] << 8 | guidData[4];

	res.Data3 = 0;
	res.Data3 = guidData[7] << 8 | guidData[6];

	res.Data4[0] = guidData[8];
	res.Data4[1] = guidData[9];
	res.Data4[2] = guidData[10];
	res.Data4[3] = guidData[11];
	res.Data4[4] = guidData[12];
	res.Data4[5] = guidData[13];
	res.Data4[6] = guidData[14];
	res.Data4[7] = guidData[15];

	return res;
}

System::Guid IIDToGuid(const IID iid)
{
	System::Guid guid(
		iid.Data1,
		iid.Data2,
		iid.Data3,
		iid.Data4[0],
		iid.Data4[1],
		iid.Data4[2],
		iid.Data4[3],
		iid.Data4[4],
		iid.Data4[5],
		iid.Data4[6],
		iid.Data4[7]);

	return guid;
}

System::Object^ VariantToObject(VARIANT& val)
{
	return System::Runtime::InteropServices::Marshal::GetObjectForNativeVariant((System::IntPtr) &val);	
}

VARIANT& ObjectToVariant(System::Object^ val)
{
	VARIANT variant;
	VariantInit(&variant);

	System::Runtime::InteropServices::Marshal::GetNativeVariantForObject(val, (System::IntPtr) &variant);

	return variant;
}