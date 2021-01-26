#include "stdafx.h"
#include "Enumerator.h"


using namespace OPC::Enum;

IOPCServerList2 * OPC::Enum::Enumerator::GetServer()
{
	return (IOPCServerList2*) server;
}

OPC::Enum::Enumerator::Enumerator() : ComBase()
{
	HRESULT hr = ComBase::Connect(
		IIDToGuid(__uuidof(OpcServerList)),
		IIDToGuid(__uuidof(IOPCServerList2)));

	if (hr != S_OK)
		throw gcnew System::ComponentModel::Win32Exception(hr);
}


OPC::Enum::Enumerator::Enumerator(
	System::String^ host,
	System::Net::NetworkCredential^ credentials) :ComBase()
{
	HRESULT hr = ComBase::Connect(
		IIDToGuid(__uuidof(OpcServerList)),
		host,
		credentials->UserName,
		credentials->Password,
		credentials->Domain,
		IIDToGuid(__uuidof(IOPCServerList2)));

	if (hr != S_OK)
		throw gcnew System::ComponentModel::Win32Exception(hr);
}

OPC::Enum::Enumerator::~Enumerator()
{
	this->Release();
}

System::Collections::Generic::IList<Server>^ OPC::Enum::Enumerator::GetAvailableServers(... array<Specification^>^ specifications)
{
	size_t count = specifications->Length;

	IOPCServerList2* serverList = GetServer();

	if (serverList == NULL)
	{
		throw gcnew System::InvalidOperationException("Enumerator instance is not connected to any OPC Enum service.");
	}

	if (count == 0)
		return nullptr;

	IID* specs = (IID*) CoTaskMemAlloc(count * sizeof(IID));

	for (size_t ii = 0; ii < count; ii++)
	{
		specs[ii] = GuidToCLSID(specifications[ii]->GUID);
	}

	IOPCEnumGUID* availableServers = nullptr;

	HRESULT hr = serverList->EnumClassesOfCategories(
		count,
		specs,
		0,
		nullptr,
		&availableServers);

	if (hr != S_OK && hr != S_FALSE)
	{
		goto end;
	}

	this->SetProxyBlanket(availableServers);

	ULONG maxGuidsToFetch = 100;
	GUID* guids = (GUID*)CoTaskMemAlloc(maxGuidsToFetch * sizeof(GUID));
	ULONG guidsFetched = 0;

	hr = availableServers->Next(
		maxGuidsToFetch,
		guids,
		&guidsFetched);

	if (hr != S_OK && hr != S_FALSE)
	{
		goto end;
	}

	System::Collections::Generic::IList<Server>^ servers = gcnew System::Collections::Generic::List<Server>(guidsFetched);

	while (guidsFetched > 0)
	{
		for (ULONG ii = 0; ii < guidsFetched; ii++)
		{
			LPOLESTR ProgId;
			LPOLESTR UserType;
			LPOLESTR VerProgId;

			HRESULT hrDet = GetServer()->GetClassDetails(
				guids[ii],
				&ProgId,
				&UserType,
				&VerProgId);

			Server serverToAdd;

			serverToAdd.m_Guid = IIDToGuid(guids[ii]);
			serverToAdd.m_ProgId = System::Runtime::InteropServices::Marshal::PtrToStringUni((System::IntPtr)ProgId);
			serverToAdd.m_UserType = System::Runtime::InteropServices::Marshal::PtrToStringUni((System::IntPtr)UserType);
			serverToAdd.m_VerProgId = System::Runtime::InteropServices::Marshal::PtrToStringUni((System::IntPtr)VerProgId);

			servers->Add(serverToAdd);

			System::Runtime::InteropServices::Marshal::ZeroFreeCoTaskMemUnicode((System::IntPtr) ProgId);
			System::Runtime::InteropServices::Marshal::ZeroFreeCoTaskMemUnicode((System::IntPtr) UserType);
			System::Runtime::InteropServices::Marshal::ZeroFreeCoTaskMemUnicode((System::IntPtr) VerProgId);
		}

		hr = availableServers->Next(
			maxGuidsToFetch,
			guids,
			&guidsFetched);
	}

end:

	
	if (guids != nullptr)
	{
		CoTaskMemFree(guids);
		guids = nullptr;
	}

	if (availableServers != nullptr)
		hr = availableServers->Release();

	if(specs != nullptr)
		CoTaskMemFree(specs);

	if (hr != S_OK && hr != S_FALSE)
		throw gcnew System::ComponentModel::Win32Exception(hr);

	return servers;
}

System::Guid OPC::Enum::Enumerator::GetCLSIDFromProgID(System::String^ progID)
{
	IOPCServerList2* serverList = GetServer();

	if (serverList == nullptr)
	{
		throw gcnew System::InvalidOperationException("Enumerator instance is not connected to any OPC Enum service.");
	}

	LPCOLESTR p = (LPCOLESTR)System::Runtime::InteropServices::Marshal::StringToCoTaskMemUni(progID).ToPointer();

	CLSID clsid;

	HRESULT hr = serverList->CLSIDFromProgID(
		p,
		&clsid);

	System::Runtime::InteropServices::Marshal::ZeroFreeCoTaskMemUnicode((System::IntPtr) (LPWSTR)p);

	if (hr != S_OK)
		throw gcnew System::ComponentModel::Win32Exception(hr);

	return IIDToGuid(clsid);
}

void OPC::Enum::Enumerator::Release() 
{
	ComBase::Release();
}

Server OPC::Enum::Enumerator::FromProgId(System::String^ progId)
{
	Server s;

	s.m_ProgId = progId;

	OPC::Enum::Enumerator^ enumerator = nullptr;

	try
	{
		enumerator = gcnew OPC::Enum::Enumerator();

		s.m_Guid = enumerator->GetCLSIDFromProgID(s.m_ProgId);
	}
	finally
	{
		if (enumerator != nullptr)
			enumerator->Release();
	}

	return s;
}