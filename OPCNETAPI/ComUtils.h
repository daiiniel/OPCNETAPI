#pragma once

#include <Windows.h>

HRESULT CreateInstance(
	const CLSID ppClsid,
	const IID iid,
	void ** ppUnknown);

HRESULT CreateInstanceEx(
	const CLSID ppClsid,
	const LPWSTR host,
	const char * user,
	const char * password,
	const char * domain,
	const IID iid,
	void ** ppUnknown);

HRESULT QueryInterface(
	IUnknown* ppInUnknown,
	const IID iid,
	void ** ppOutUnknown);

HRESULT QueryInterfaceEx(
	IUnknown* ppInUnknown,
	const LPWSTR host,
	const char * user,
	const char * password,
	const char * domain,
	const IID iid,
	void ** ppOutUnknown);

HRESULT CoSetProxyBlanket(
	IUnknown* ppInUnknown,
	const LPWSTR host,
	const char * user,
	const char * password,
	const char * domain);

CLSID GuidToCLSID(System::Guid guid);
System::Guid IIDToGuid(const IID iid);

System::Object^ VariantToObject(VARIANT& val);
VARIANT& ObjectToVariant(System::Object^ val);