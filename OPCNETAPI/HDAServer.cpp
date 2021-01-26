#include "stdafx.h"
#include "HDAServer.h"

IOPCHDA_Server* OPC::HDA::HDAServer::GetServer()
{
	if (server == NULL)
	{
		throw gcnew System::InvalidOperationException("HDA Client instance is not connected to any OPC HDA server.");
	}

	return (IOPCHDA_Server*)server;
}

IOPCHDA_SyncRead * OPC::HDA::HDAServer::GetSyncRead()
{
	if (m_IOPCSyncRead != NULL)
		return m_IOPCSyncRead;

	IOPCHDA_Server* client = GetServer();	

	IOPCHDA_SyncRead* syncRead = NULL;

	HRESULT hr = GetInterface(
		__uuidof(IOPCHDA_SyncRead),
		(void**)&syncRead);

	if (hr != S_OK)
	{
		if (syncRead != NULL)
		{
			syncRead->Release();
			syncRead = NULL;
		}

		throw gcnew System::ComponentModel::Win32Exception(hr);
	}

	m_IOPCSyncRead = syncRead;

	return m_IOPCSyncRead;
}

IOPCHDA_Browser * OPC::HDA::HDAServer::GetBrowser()
{
	if (m_IOPCBrowser != NULL)
		return m_IOPCBrowser;

	IOPCHDA_Server* client = GetServer();

	IOPCHDA_Browser* browser = NULL;
	HRESULT* results = NULL;

	unsigned long attrs[1] = { OPCHDA_NODE_NAME };
	tagOPCHDA_OPERATORCODES codes[1] = { tagOPCHDA_OPERATORCODES::OPCHDA_EQUAL };
	VARIANT filter[1];

	HRESULT hr = client->CreateBrowse(
		0,
		attrs,
		codes,
		filter,
		&browser,
		&results);

	VariantClear(&filter[0]);

	if (hr != S_OK && hr != S_FALSE)
	{
		if (results != NULL)
		{
			CoTaskMemFree(results);
			results = NULL;
		}

		if (browser != NULL)
		{
			browser->Release();
			browser = NULL;
		}

		return nullptr;
	}

	m_IOPCBrowser = browser;

	return m_IOPCBrowser;
}

IOPCHDA_AsyncRead* OPC::HDA::HDAServer::GetAsyncRead()
{
	if (m_IOPCAsyncRead != nullptr)
	{
		return m_IOPCAsyncRead;
	}

	HRESULT hr;
	IOPCHDA_AsyncRead* asyncRead = nullptr;

	IOPCHDA_Server* server = this->GetServer();

	hr = GetInterface(
		__uuidof(IOPCHDA_AsyncRead),
		(void**) &asyncRead);

	if (hr != S_OK)
	{
		if (asyncRead != nullptr)
		{
			asyncRead->Release();
			asyncRead = nullptr;
		}

		throw gcnew System::ComponentModel::Win32Exception(hr);
	}

	m_IOPCAsyncRead = asyncRead;

	return asyncRead;
}

OPC::HDA::HDAServer::HDAServer() : ComBase(), cCounter(1), m_IOPCBrowser(nullptr), m_IOPCSyncRead(nullptr), m_IOPCAsyncRead(nullptr)
{
	
}

void OPC::HDA::HDAServer::Connect(Server s)
{
	HRESULT hr = ComBase::Connect(s.GUID, IIDToGuid(__uuidof(IOPCHDA_Server)));

	if (hr != S_OK)
		throw gcnew System::ComponentModel::Win32Exception(hr);
}

void OPC::HDA::HDAServer::Connect(Server s, System::String ^ host, System::Net::NetworkCredential ^ credentials)
{
	HRESULT hr = ComBase::Connect(
		s.GUID,
		host,
		credentials->UserName,
		credentials->Password,
		credentials->Domain,
		IIDToGuid(__uuidof(IOPCHDA_Server)));

	if (hr != S_OK)
		throw gcnew System::ComponentModel::Win32Exception(hr);
}

void OPC::HDA::HDAServer::Connect(URL url, System::Net::NetworkCredential ^ credentials)
{
	HDAServer::Connect(
		url.ToServer(credentials),
		url.Host, credentials);
}

delegate void LocalServerConnectHandler(Server s);

void OPC::HDA::HDAServer::ConnectAsync(
	Server s,
	[System::Runtime::InteropServices::Optional] System::Nullable<int> timeout)
{
	if (!timeout.HasValue)
	{
		timeout = 15000;
	}

	LocalServerConnectHandler^ handler = gcnew LocalServerConnectHandler(this, &HDAServer::Connect);

	System::IAsyncResult^ result = handler->BeginInvoke(s, nullptr, nullptr);

	if (!result->AsyncWaitHandle->WaitOne(timeout.Value, true))
	{
		throw gcnew System::TimeoutException();
	}

	if (!result->IsCompleted)
		throw gcnew System::Exception("Could not connect to the server");

	try
	{
		return handler->EndInvoke(result);
	}
	catch (System::Exception^)
	{
		throw;
	}
}

delegate void RemoteServerConnectHandler(
	Server s,
	System::String^ host,
	System::Net::NetworkCredential^credentials);

void OPC::HDA::HDAServer::ConnectAsync(
	Server s,
	System::String^ host,
	System::Net::NetworkCredential^credentials,
	[System::Runtime::InteropServices::Optional] System::Nullable<int> timeout)
{
	if (!timeout.HasValue)
	{
		timeout = 15000;
	}

	RemoteServerConnectHandler^ handler = gcnew RemoteServerConnectHandler(this, &HDAServer::Connect);

	System::IAsyncResult^ result = handler->BeginInvoke(s, host, credentials, nullptr, nullptr);

	if (!result->AsyncWaitHandle->WaitOne(timeout.Value, true))
	{
		throw gcnew System::TimeoutException();
	}

	if (!result->IsCompleted)
		throw gcnew System::Exception("Could not connect to the server");

	try
	{
		return handler->EndInvoke(result);
	}
	catch (System::Exception^)
	{
		throw;
	}
}

delegate void RemoteURLConnectHandler(
	URL url,
	System::Net::NetworkCredential^ credentials);

void OPC::HDA::HDAServer::ConnectAsync(
	URL url,
	System::Net::NetworkCredential^ credentials,
	[System::Runtime::InteropServices::Optional] System::Nullable<int> timeout)
{
	if (!timeout.HasValue)
	{
		timeout = 15000;
	}

	RemoteURLConnectHandler^ handler = gcnew RemoteURLConnectHandler(this, &HDAServer::Connect);

	System::IAsyncResult^ result = handler->BeginInvoke(url, credentials, nullptr, nullptr);

	if (!result->AsyncWaitHandle->WaitOne(timeout.Value, true))
	{
		throw gcnew System::TimeoutException();
	}

	if (!result->IsCompleted)
		throw gcnew System::Exception("Could not connect to the server");

	try
	{
		return handler->EndInvoke(result);
	}
	catch (System::Exception^)
	{
		throw;
	}
}


OPC::HDA::HDAServer::~HDAServer()
{
	this->Release();
}

void OPC::HDA::HDAServer::ValidateItemIDs(System::Collections::Generic::List<System::Tuple<System::String^, bool>^> ^ itemsToValidate)
{
	HRESULT hr;

	IOPCHDA_Server* server = this->GetServer();

	DWORD count = itemsToValidate->Count;

	if (!count)
	{
		return;
	}

	LPWSTR* items = (LPWSTR*) CoTaskMemAlloc(count * sizeof(LPWSTR));
	HRESULT* results = nullptr;

	for (DWORD ii = 0; ii < count; ii++)
	{
		items[ii] = (LPWSTR) System::Runtime::InteropServices::Marshal::StringToHGlobalUni(itemsToValidate[ii]->Item1).ToPointer();
	}

	hr = server->ValidateItemIDs(
		count,
		items,
		&results);

	if (hr != S_OK && hr != S_FALSE)
	{
		goto end;
	}

	hr = S_OK;

	for (DWORD ii = 0; ii < count; ii++)
	{
		bool validated = false;

		switch (results[ii])
		{
		case S_OK:
			validated = true;

			break;
		case OPC_E_BADRIGHTS:
		case OPC_E_UNKNOWNITEMID:
		case OPC_E_INVALIDITEMID:
			validated = false;

			break;
		default:
			break;
		}

		itemsToValidate[ii] = gcnew System::Tuple<System::String^, bool>(itemsToValidate[ii]->Item1, validated);

		System::Runtime::InteropServices::Marshal::FreeHGlobal((System::IntPtr) items[ii]);
	}

end:

	if (items != nullptr)
	{
		CoTaskMemFree(items);
		items = nullptr;
	}

	if (results != nullptr)
	{
		CoTaskMemFree(results);
		results = nullptr;
	}

	if (hr != S_OK)
	{
		throw gcnew System::ComponentModel::Win32Exception(hr);
	}
}

HistorianStatus^ OPC::HDA::HDAServer::GetHistorianStatus()
{
	HRESULT hr;

	HistorianStatus^ res = nullptr;

	IOPCHDA_Server* server = this->GetServer();

	OPCHDA_SERVERSTATUS status;

	FILETIME * currentTime = nullptr;
	FILETIME * startTime = nullptr;

	WORD major;
	WORD minor;
	WORD buildNumber;
	DWORD maxReturnValues;

	LPWSTR statusString;
	LPWSTR vendorInfo;

	hr = server->GetHistorianStatus(
		&status,
		&currentTime,
		&startTime,
		&major,
		&minor,
		&buildNumber,
		&maxReturnValues,
		&statusString,
		&vendorInfo);

	res = HistorianStatus::Create(
		currentTime,
		startTime,
		major,
		minor,
		buildNumber,
		maxReturnValues,
		statusString,
		vendorInfo);

end:

	if (currentTime != nullptr)
	{
		CoTaskMemFree(currentTime);
		currentTime = nullptr;
	}

	if (startTime != nullptr)
	{
		CoTaskMemFree(startTime);
		startTime = nullptr;
	}

	if (hr != S_OK)
	{
		throw gcnew System::ComponentModel::Win32Exception(hr);
	}

	return res;
}

System::Collections::Generic::IList<ItemDefinition^>^ OPC::HDA::HDAServer::Browse(System::String^ nodeToBrowse, BrowseType browseType)
{
	IOPCHDA_Browser* browser = GetBrowser();

	HRESULT hr;

	if (!System::String::IsNullOrWhiteSpace(nodeToBrowse))
	{
		LPWSTR itemToBrowse = (LPWSTR)System::Runtime::InteropServices::Marshal::StringToCoTaskMemUni(nodeToBrowse).ToPointer();

		hr = browser->ChangeBrowsePosition(
			OPCHDA_BROWSEDIRECTION::OPCHDA_BROWSE_DIRECT,
			itemToBrowse);

		System::Runtime::InteropServices::Marshal::ZeroFreeCoTaskMemUnicode((System::IntPtr)itemToBrowse);
	}

	IEnumString* enumerator = NULL;

	tagOPCHDA_BROWSETYPE t;

	switch (browseType)
	{
	case BrowseType::Flat:
		t = tagOPCHDA_BROWSETYPE::OPCHDA_FLAT;
		break;
	case BrowseType::Branch:
		t = tagOPCHDA_BROWSETYPE::OPCHDA_BRANCH;
		break;
	case BrowseType::Leaf:
		t = tagOPCHDA_BROWSETYPE::OPCHDA_LEAF;
		break;
	case BrowseType::Items:
		t = tagOPCHDA_BROWSETYPE::OPCHDA_ITEMS;
		break;
	default:
		break;
	}

	hr = browser->GetEnum(
		t,
		&enumerator);

	if (hr != S_OK && hr != S_FALSE)
	{
		if (enumerator != NULL)
		{
			enumerator->Release();
			enumerator = NULL;
		}

		return nullptr;
	}
	
	System::Collections::Generic::IList<ItemDefinition^>^ definitions = gcnew System::Collections::Generic::List<ItemDefinition^>();

	ULONG itemsToFetch = 100;
	LPOLESTR* items = (LPOLESTR*) CoTaskMemAlloc(itemsToFetch * sizeof(LPOLESTR));
	ULONG itemsFetched = 0;

	hr = enumerator->Next(itemsToFetch, items, &itemsFetched);

	while (itemsFetched > 0)
	{
		for (ULONG ii = 0; ii < itemsFetched; ii++)
		{
			LPOLESTR it = items[ii];
			LPWSTR id = NULL;

			browser->GetItemID(it, &id);

			ItemDefinition^ def = gcnew ItemDefinition(System::Runtime::InteropServices::Marshal::PtrToStringUni((System::IntPtr)id));

			def->SetItemID(System::Runtime::InteropServices::Marshal::PtrToStringUni((System::IntPtr)it));

			System::Runtime::InteropServices::Marshal::ZeroFreeCoTaskMemUnicode((System::IntPtr)it);
			System::Runtime::InteropServices::Marshal::ZeroFreeCoTaskMemUnicode((System::IntPtr)id);

			definitions->Add(def);
		}

		hr = enumerator->Next(itemsToFetch, items, &itemsFetched);
	}

	if (items != NULL)
	{
		CoTaskMemFree(items);
		items = NULL;
	}

	if (enumerator != NULL)
	{
		enumerator->Release();
		enumerator = NULL;
	}

	return definitions;


error:

	if (items != NULL)
	{
		CoTaskMemFree(items);
		items = NULL;
	}

	if (enumerator != NULL)
	{
		enumerator->Release();
		enumerator = NULL;
	}
}

delegate void definitionsHandler(System::Collections::Generic::IList<ItemDefinition^>^ definitions);

void OPC::HDA::HDAServer::GetHandlesAsync(System::Collections::Generic::IList<ItemDefinition^>^ definitions)
{
	definitionsHandler^ handler = gcnew definitionsHandler(this, &OPC::HDA::HDAServer::GetHandles);

	System::IAsyncResult^ result = handler->BeginInvoke(definitions, nullptr, nullptr);

	if (!result->AsyncWaitHandle->WaitOne(10000, true))
	{
		throw gcnew System::TimeoutException();
	}

	if (!result->IsCompleted)
		throw gcnew System::Exception("Could not add the items.");

	try
	{
		return handler->EndInvoke(result);
	}
	catch (System::Exception^)
	{
		throw;
	}
}

void OPC::HDA::HDAServer::GetHandles(System::Collections::Generic::IList<ItemDefinition^>^ definitions)
{
	IOPCHDA_Server* client = GetServer();

	if (client == NULL)
	{
		throw gcnew System::InvalidOperationException("HDA Client instance is not connected to any OPC HDA server.");
	}
	
	DWORD count = (DWORD) definitions->Count;

	if (count == 0)
		return;

	LPWSTR* itemIDs = (LPWSTR*)CoTaskMemAlloc(count * sizeof(LPWSTR));
	DWORD* clientHandles = (DWORD*)CoTaskMemAlloc(count * sizeof(DWORD));

	for (DWORD ii = 0; ii < count; ii++)
	{
		System::String^ id = !System::String::IsNullOrWhiteSpace(definitions[ii]->ItemID) ? definitions[ii]->ItemID : definitions[ii]->Name;

		itemIDs[ii] = (LPWSTR) System::Runtime::InteropServices::Marshal::StringToCoTaskMemUni(id).ToPointer();
		clientHandles[ii] = (DWORD) System::Threading::Interlocked::Increment((int%)cCounter);
	}

	DWORD* serverHandles = NULL;
	HRESULT* results = NULL;

	HRESULT hr = client->GetItemHandles(
		count,
		itemIDs,
		clientHandles,
		&serverHandles,
		&results);

	if ((hr == S_OK || hr == S_FALSE) && serverHandles != NULL && results != NULL)
	{
		for (DWORD ii = 0; ii < count; ii++)
		{
			if (results[ii] == S_OK)
			{
				definitions[ii]->SetClientHandle(clientHandles[ii]);
				definitions[ii]->SetServerHandle(serverHandles[ii]);
			}

			definitions[ii]->SetResult((Results)results[ii]);

			System::Runtime::InteropServices::Marshal::ZeroFreeCoTaskMemUnicode((System::IntPtr)itemIDs[ii]);
		}
	}

	if (itemIDs != NULL)
	{
		CoTaskMemFree(itemIDs);
		itemIDs = NULL;
	}

	if (clientHandles != NULL)
	{
		CoTaskMemFree(clientHandles);
		clientHandles = NULL;
	}

	if (serverHandles != NULL)
	{
		CoTaskMemFree(serverHandles);
		serverHandles = NULL;
	}

	if (results != NULL)
	{
		CoTaskMemFree(results);
		results = NULL;
	}
}

void OPC::HDA::HDAServer::RemoveHandlesAsync(System::Collections::Generic::IList<ItemDefinition^>^ definitions)
{
	definitionsHandler^ handler = gcnew definitionsHandler(this, &OPC::HDA::HDAServer::RemoveHandles);

	System::IAsyncResult^ result = handler->BeginInvoke(definitions, nullptr, nullptr);

	if (!result->AsyncWaitHandle->WaitOne(10000, true))
	{
		throw gcnew System::TimeoutException();
	}

	if (!result->IsCompleted)
		throw gcnew System::Exception("Could not add the items.");

	try
	{
		return handler->EndInvoke(result);
	}
	catch (System::Exception^)
	{
		throw;
	}
}

void OPC::HDA::HDAServer::RemoveHandles(System::Collections::Generic::IList<ItemDefinition^>^ definitions)
{
	IOPCHDA_Server* client = GetServer();

	if (client == NULL)
	{
		throw gcnew System::InvalidOperationException("HDA Client instance is not connected to any OPC HDA server.");
	}

	DWORD count = (DWORD) definitions->Count;

	if (count == 0)
		return;

	DWORD* serverHandles = (DWORD*)CoTaskMemAlloc(count * sizeof(DWORD));
	HRESULT* results = NULL;

	HRESULT hr = client->ReleaseItemHandles(
		count,
		serverHandles,
		&results);

	for (DWORD ii = 0; ii < count; ii++)
	{
		definitions[ii]->SetClientHandle(0);
		definitions[ii]->SetServerHandle(0);
	}

	if (results != NULL)
	{
		CoTaskMemFree(results);
		results = NULL;
	}

	if (serverHandles != NULL)
	{
		CoTaskMemFree(serverHandles);
		serverHandles = NULL;
	}
}

System::Collections::Generic::IList<Aggregate^>^ OPC::HDA::HDAServer::GetAggregates()
{
	IOPCHDA_Server* server = this->GetServer();

	System::Collections::Generic::List<Aggregate^>^ aggregates = nullptr;

	DWORD count;
	HRESULT hr;

	DWORD* ids = nullptr;
	LPWSTR* names = nullptr;
	LPWSTR* descriptions = nullptr;

	hr = server->GetAggregates(
		&count, 
		&ids,
		&names, 
		&descriptions);

	if (hr != S_OK && hr != S_FALSE)
	{
		goto end;
	}

	hr = S_OK;

	aggregates = gcnew System::Collections::Generic::List<Aggregate^>(count);

	for (DWORD ii = 0; ii < count; ii++)
	{
		Aggregate^ a = Aggregate::Create(ids[ii], names[ii], descriptions[ii]);

		aggregates->Add(a);

		CoTaskMemFree(names[ii]);
		CoTaskMemFree(descriptions[ii]);
	}

end:

	if (ids != nullptr)
	{
		CoTaskMemFree(ids);
		ids = nullptr;
	}

	if (names != nullptr)
	{
		CoTaskMemFree(names);
		names = nullptr;
	}

	if (descriptions != nullptr)
	{
		CoTaskMemFree(descriptions);
		descriptions = nullptr;
	}

	if (hr != S_OK)
	{
		throw gcnew System::ComponentModel::Win32Exception(hr);
	}

	return aggregates;
}

HRESULT OPC::HDA::HDAServer::processResults(
	System::Collections::Generic::IList<ItemDefinition^>^ definitions, 
	System::Collections::Generic::IList<HDAReadResult^>^ res, 
	OPCHDA_ITEM* values, 
	HRESULT* errors, 
	size_t count)
{
	for (DWORD ii = 0; ii < count; ii++)
	{
		DWORD valuesCount = values[ii].dwCount;

		HDAReadResult^ r = gcnew HDAReadResult(definitions[ii], valuesCount);

		r->SetResult((Results)errors[ii]);

		for (DWORD iii = 0; iii < valuesCount; iii++)
		{
			ReadResult rr;

			rr.SetValue(VariantToObject(values[ii].pvDataValues[iii]));
			rr.SetQuality((Qualities)(byte)(values[ii].pdwQualities[iii]));
			rr.SetTimeStamp(FileTimeToDateTime(values[ii].pftTimeStamps[iii]));

			r->Values->Add(rr);

			VariantClear(&values[ii].pvDataValues[iii]);
		}

		CoTaskMemFree(values[ii].pvDataValues);
		CoTaskMemFree(values[ii].pdwQualities);
		CoTaskMemFree(values[ii].pftTimeStamps);

		res->Add(r);
	}

	return S_OK;
}

System::Collections::Generic::IList<HDAReadResult^>^ OPC::HDA::HDAServer::ReadRaw(
	System::DateTime% from, 
	System::DateTime% to, 
	System::Collections::Generic::IList<ItemDefinition^>^ definitions, 
	System::UInt32 maxItems, bool includeBounds)
{
	IOPCHDA_SyncRead* syncRead = GetSyncRead();

	HRESULT hr;

	from = from.ToUniversalTime();
	to = to.ToUniversalTime();

	DWORD count = (DWORD)definitions->Count;

	if (count == 0)
		return nullptr;

	OPCHDA_TIME oFrom = DatetimeToOPCTime(from);
	OPCHDA_TIME oTo = DatetimeToOPCTime(to);

	DWORD* serverHandles = (DWORD*)CoTaskMemAlloc(count* sizeof(DWORD));

	for (DWORD ii = 0; ii < count; ii++)
	{
		serverHandles[ii] = definitions[ii]->ServerHandle;
	}

	OPCHDA_ITEM* values = NULL;
	HRESULT* errors = NULL;

	hr = syncRead->ReadRaw(
		&oFrom,
		&oTo,
		maxItems,
		includeBounds,
		count,
		serverHandles,
		&values,
		&errors);

	from = OPCTimeToDatetime(oFrom);
	to = OPCTimeToDatetime(oTo);

	System::Collections::Generic::IList<HDAReadResult^>^ res = gcnew System::Collections::Generic::List<HDAReadResult^>(count);

	if ((hr != S_OK && hr != S_FALSE) || values == NULL || errors == NULL)
	{
		goto end;
	}

	hr = processResults(
		definitions,
		res,
		values,
		errors,
		count);

end:

	if (errors != NULL)
	{
		CoTaskMemFree(values);
		values = NULL;
	}

	if (values != NULL)
	{
		CoTaskMemFree(values);
		values = NULL;
	}

	if (errors != NULL)
	{
		CoTaskMemFree(errors);
		errors = NULL;
	}

	if (serverHandles != NULL)
	{
		CoTaskMemFree(serverHandles);
		serverHandles = NULL;
	}

	if(hr != S_OK)
		throw gcnew System::ComponentModel::Win32Exception(hr);

	return res;
}

System::Collections::Generic::IList<HDAReadResult^>^ OPC::HDA::HDAServer::ReadProcessed(
	System::DateTime% from,
	System::DateTime% to,
	System::TimeSpan interval,
	System::UInt32 aggregate,
	System::Collections::Generic::IList<ItemDefinition^>^ definitions)
{
	HRESULT hr;
	System::Collections::Generic::IList<HDAReadResult^>^ hdaReadResults = nullptr;

	IOPCHDA_SyncRead* server = this->GetSyncRead();

	from = from.ToUniversalTime();
	to = to.ToUniversalTime();

	OPCHDA_TIME oFrom = DatetimeToOPCTime(from);
	OPCHDA_TIME oTo = DatetimeToOPCTime(to);

	FILETIME fInterval = TimeSpanToFileTime(interval);

	DWORD nItems = (DWORD)definitions->Count;

	if (!nItems)
		return nullptr;

	OPCHANDLE* handles = (OPCHANDLE*) CoTaskMemAlloc(nItems * sizeof(OPCHANDLE));
	DWORD*  aggregates = (DWORD*)CoTaskMemAlloc(nItems * sizeof(DWORD));

	for (size_t ii = 0; ii < nItems; ii++)
	{
		handles[ii] = definitions[ii]->ServerHandle;
		aggregates[ii] = aggregate;
	}

	OPCHDA_ITEM* values = nullptr;
	HRESULT* results = nullptr;

	hr = server->ReadProcessed(
		&oFrom,
		&oTo,
		fInterval,
		nItems, 
		handles, 
		aggregates,
		&values,
		&results);

	if ((hr != S_OK && hr != S_FALSE) || values == NULL || results == NULL)
	{
		goto end;
	}

	hdaReadResults = gcnew System::Collections::Generic::List<HDAReadResult^>(nItems);

	hr = processResults(
		definitions, hdaReadResults, values, results, nItems);

end:

	if (values != nullptr)
	{
		CoTaskMemFree(values);
		values = nullptr;
	}

	if (results != nullptr)
	{
		CoTaskMemFree(results);
		values = nullptr;
	}

	if (handles != nullptr)
	{
		CoTaskMemFree(handles);
		handles = nullptr;
	}

	if (aggregates != nullptr)
	{
		CoTaskMemFree(aggregates);
		aggregates = nullptr;
	}

	if (hr != S_OK)
	{
		throw gcnew System::ComponentModel::Win32Exception(hr);
	}

	return hdaReadResults;
}

System::Collections::Generic::IList<HDAReadResult^>^ OPC::HDA::HDAServer::ReadProcessed(
	System::DateTime% from,
	System::DateTime% to,
	System::TimeSpan interval,
	System::Collections::Generic::IList<AggregateItem^>^ definitions)
{
	HRESULT hr;
	System::Collections::Generic::IList<HDAReadResult^>^ hdaReadResults = nullptr;

	IOPCHDA_SyncRead* server = this->GetSyncRead();

	from = from.ToUniversalTime();
	to = to.ToUniversalTime();

	OPCHDA_TIME oFrom = DatetimeToOPCTime(from);
	OPCHDA_TIME oTo = DatetimeToOPCTime(to);

	FILETIME fInterval = TimeSpanToFileTime(interval);

	DWORD nItems = (DWORD)definitions->Count;

	if (!nItems)
		return nullptr;

	OPCHANDLE* handles = (OPCHANDLE*)CoTaskMemAlloc(nItems * sizeof(OPCHANDLE));
	DWORD*  aggregates = (DWORD*)CoTaskMemAlloc(nItems * sizeof(DWORD));

	for (size_t ii = 0; ii < nItems; ii++)
	{
		handles[ii] = definitions[ii]->Item->ServerHandle;
		aggregates[ii] = definitions[ii]->Aggregate;
	}

	OPCHDA_ITEM* values = nullptr;
	HRESULT* results = nullptr;

	hr = server->ReadProcessed(
		&oFrom,
		&oTo,
		fInterval,
		nItems,
		handles,
		aggregates,
		&values,
		&results);

	if ((hr != S_OK && hr != S_FALSE) || values == NULL || results == NULL)
	{
		goto end;
	}

	hdaReadResults = gcnew System::Collections::Generic::List<HDAReadResult^>(nItems);

	for (DWORD ii = 0; ii < nItems; ii++)
	{
		DWORD valuesCount = values[ii].dwCount;

		HDAReadResult^ r = gcnew HDAReadResult(definitions[ii]->Item, valuesCount);

		r->SetResult((Results)results[ii]);

		for (DWORD iii = 0; iii < valuesCount; iii++)
		{
			ReadResult rr;

			rr.SetValue(VariantToObject(values[ii].pvDataValues[iii]));
			rr.SetQuality((Qualities)(byte)(values[ii].pdwQualities[iii]));
			rr.SetTimeStamp(FileTimeToDateTime(values[ii].pftTimeStamps[iii]));

			r->Values->Add(rr);

			VariantClear(&values[ii].pvDataValues[iii]);
		}

		CoTaskMemFree(values[ii].pvDataValues);
		CoTaskMemFree(values[ii].pdwQualities);
		CoTaskMemFree(values[ii].pftTimeStamps);

		hdaReadResults->Add(r);
	}

end:

	if (values != nullptr)
	{
		CoTaskMemFree(values);
		values = nullptr;
	}

	if (results != nullptr)
	{
		CoTaskMemFree(results);
		values = nullptr;
	}

	if (handles != nullptr)
	{
		CoTaskMemFree(handles);
		handles = nullptr;
	}

	if (aggregates != nullptr)
	{
		CoTaskMemFree(aggregates);
		aggregates = nullptr;
	}

	if (hr != S_OK)
	{
		throw gcnew System::ComponentModel::Win32Exception(hr);
	}

	return hdaReadResults;
}

System::Collections::Generic::IList<HDAReadResult^>^ OPC::HDA::HDAServer::ReadAtTime(
	System::Collections::Generic::IList<System::DateTime>^ timeStamps,
	System::Collections::Generic::IList<ItemDefinition^>^definitions)
{
	System::Collections::Generic::IList<HDAReadResult^>^ res = nullptr;

	HRESULT hr;

	IOPCHDA_SyncRead* server = this->GetSyncRead();

	DWORD countTimeStamps = timeStamps->Count;
	DWORD countItems = definitions->Count;

	FILETIME* ftimeStamps = (FILETIME*) CoTaskMemAlloc(countTimeStamps * sizeof(FILETIME));
	OPCHANDLE* hItems = (OPCHANDLE*)CoTaskMemAlloc(countTimeStamps * sizeof(OPCHANDLE));

	for (DWORD ii = 0; ii < countTimeStamps; ii++)
	{
		ftimeStamps[ii] = DateTimeToFileTime(timeStamps[ii]);
	}

	for (DWORD ii = 0; ii < countItems; ii++)
	{
		hItems[ii] = definitions[ii]->ServerHandle;
	}

	OPCHDA_ITEM* values = nullptr;
	HRESULT* results = nullptr;

	hr = server->ReadAtTime(
		countTimeStamps,
		ftimeStamps,
		countItems,
		hItems,
		&values,
		&results);

	if ((hr != S_OK && hr != S_FALSE) || values == nullptr || results == nullptr)
	{
		goto end;
	}

	res = gcnew System::Collections::Generic::List<HDAReadResult^>(countItems);

	hr = processResults(definitions, res, values, results, countItems);



end:

	if (results != nullptr)
	{
		CoTaskMemFree(results);
		results = nullptr;
	}

	if (values != nullptr)
	{
		CoTaskMemFree(values);
		values == nullptr;
	}

	if (ftimeStamps != nullptr)
	{
		CoTaskMemFree(ftimeStamps);
		ftimeStamps = nullptr;
	}

	if (hItems != nullptr)
	{
		CoTaskMemFree(hItems);
		hItems = nullptr;
	}

	if(hr != S_OK)
	{
		throw gcnew System::ComponentModel::Win32Exception(hr);
	}

	return res;
}

System::Collections::Generic::IList<HDAResult^>^ OPC::HDA::HDAServer::AdviseRaw(
	System::UInt32 transactionId,
	System::DateTime startTime,
	System::TimeSpan interval,
	System::Collections::Generic::IList<ItemDefinition^>^ definitions,
	[System::Runtime::InteropServices::OutAttribute] System::UInt32% cancelId)
{
	HRESULT hr;

	IOPCHDA_AsyncRead* asyncRead = this->GetAsyncRead();

	System::Collections::Generic::IList<HDAResult^>^ res = nullptr;

	DWORD dwTransactionId = transactionId;
	OPCHDA_TIME htStartTime = DatetimeToOPCTime(startTime);
	FILETIME ftUpdateInterval = TimeSpanToFileTime(interval);
	DWORD count = definitions->Count;

	if (count == 0)
		goto end;

	OPCHANDLE* handles = (OPCHANDLE*)CoTaskMemAlloc(count * sizeof(OPCHANDLE));

	for (int ii = 0; ii < count; ii++)
	{
		handles[ii] = definitions[ii]->ServerHandle;
	}

	DWORD pCancelId;
	HRESULT* results = nullptr;

	hr = asyncRead->AdviseRaw(
		dwTransactionId,
		&htStartTime,
		ftUpdateInterval,
		count,
		handles,
		&pCancelId,
		&results);

	if (hr != S_OK && hr != S_FALSE)
	{
		goto end;
	}

	res = gcnew System::Collections::Generic::List<HDAResult^>(count);

	for (DWORD ii = 0; ii < count; ii++)
	{
		HDAResult^ result = gcnew HDAResult(definitions[ii]);

		result->SetResult((OPC::Results) results[ii]);
	}

end:

	if (handles != nullptr)
	{
		CoTaskMemFree(handles);
		handles = nullptr;
	}

	if (results != nullptr)
	{
		CoTaskMemFree(results);
		results = nullptr;
	}

	if (hr != S_OK)
	{
		throw gcnew System::ComponentModel::Win32Exception(hr);
	}

	return res;
}

void OPC::HDA::HDAServer::Release()
{
	if (m_IOPCBrowser != nullptr)
	{
		m_IOPCBrowser->Release();
		m_IOPCBrowser = nullptr;
	}

	if (m_IOPCSyncRead != nullptr)
	{
		m_IOPCSyncRead->Release();
		m_IOPCSyncRead = nullptr;
	}

	if (m_IOPCAsyncRead != nullptr)
	{
		m_IOPCAsyncRead->Release();
		m_IOPCAsyncRead = nullptr;
	}

	ComBase::Release();
}