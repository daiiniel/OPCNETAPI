#include "stdafx.h"
#include "OPCShutdown.h"


HRESULT OPC::DA::DAServer::OnServerShutDownFunc(LPCWSTR szReason)
{
	System::String^ reason = System::Runtime::InteropServices::Marshal::PtrToStringUni((System::IntPtr)(LPWSTR)szReason);

	OnServerShutDown(
		gcnew ServerShutdownEventArgs(
			reason));
	
	return S_OK;
}

IOPCServer* OPC::DA::DAServer::GetServer()
{
	if (server == NULL)
	{
		throw gcnew System::InvalidOperationException("DA Client instance is not connected to any OPC DA server.");
	}

	return (IOPCServer*)ComBase::server;
}

IOPCBrowseServerAddressSpace* DAServer::GetBrowser()
{
	if (m_IOPCAddressSpaceBrowser != NULL)
	{
		return m_IOPCAddressSpaceBrowser;
	}

	IOPCServer* client = GetServer();

	HRESULT hr = E_FAIL;

	IOPCBrowseServerAddressSpace* browser = NULL;

	hr = ComBase::GetInterface(
		__uuidof(IOPCBrowseServerAddressSpace),
		(void**)&browser);

	if (hr != S_OK || browser == NULL)
	{
		if (browser != NULL)
		{
			browser->Release();
			browser = NULL;
		}

		throw gcnew System::InvalidOperationException("Could not get browser from the remote OPC Server.");
	}

	m_IOPCAddressSpaceBrowser = browser;

	return m_IOPCAddressSpaceBrowser;
}

OPC::DA::DAServer::DAServer() : ComBase(), cCounter(1), m_shutdown(nullptr), m_IOPCAddressSpaceBrowser(nullptr)
{
	handle = System::Runtime::InteropServices::GCHandle::Alloc(this);
}

OPC::DA::DAServer::~DAServer()
{
	this->Release();
}


void OPC::DA::DAServer::OnServerShutDown(ServerShutdownEventArgs^ args)
{
	if (m_serverShutDownHandler != nullptr)
	{
		m_serverShutDownHandler(this, args);
	}
}

void OPC::DA::DAServer::Connect(Server s)
{
	HRESULT hr = ComBase::Connect(s.GUID, IIDToGuid(__uuidof(IOPCServer)));

	if (hr != S_OK)
		throw gcnew System::ComponentModel::Win32Exception(hr);

	this->Init();
}

void OPC::DA::DAServer::Connect(Server s, System::String ^ host, System::Net::NetworkCredential ^ credentials)
{
	HRESULT hr = ComBase::Connect(
		s.GUID,
		host,
		credentials->UserName,
		credentials->Password,
		credentials->Domain,
		IIDToGuid(__uuidof(IOPCServer)));

	if (hr != S_OK)
		throw gcnew System::ComponentModel::Win32Exception(hr);

	this->Init();
}

void OPC::DA::DAServer::Connect(URL url, System::Net::NetworkCredential ^ credentials)
{
	DAServer::Connect(
		url.ToServer(credentials),
		url.Host, 
		credentials);
}

void OPC::DA::DAServer::Init()
{
	HRESULT hr;

	this->m_shutdown = new OPCShutdown();

	IOPCServer* client = GetServer();

	IConnectionPointContainer* container = NULL;

	hr = this->GetInterface(
		__uuidof(IConnectionPointContainer),
		(void**)&container);

	if (hr != S_OK)
	{
		if (container != NULL)
		{
			container->Release();
			container = NULL;
		}

		return;
	}

	IConnectionPoint* connectionPoint = NULL;

	hr = container->FindConnectionPoint(
		IID_IOPCShutdown,
		&connectionPoint);

	if (hr != S_OK)
	{
		if (connectionPoint != NULL)
		{
			connectionPoint->Release();
			connectionPoint = NULL;
		}

		if (container != NULL)
		{
			container->Release();
			container = NULL;
		}

		return;
	}

	hr = ComBase::SetProxyBlanket(
		connectionPoint);

	hr = m_shutdown->Advise(
		connectionPoint,
		System::Runtime::InteropServices::GCHandle::ToIntPtr(handle));

	if (connectionPoint != NULL)
	{
		connectionPoint->Release();
		connectionPoint = NULL;
	}

	if (container != NULL)
	{
		container->Release();
		container = NULL;
	}

	subscriptions = gcnew System::Collections::Generic::Dictionary<System::UInt32, Subscription^>();
}

delegate void LocalServerConnectHandler(Server s);

void OPC::DA::DAServer::ConnectAsync(
	Server s,
	[System::Runtime::InteropServices::Optional] System::Nullable<int> timeout)
{
	if (!timeout.HasValue)
	{
		timeout = 15000;
	}

	LocalServerConnectHandler^ handler = gcnew LocalServerConnectHandler(this, &DAServer::Connect);

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

void OPC::DA::DAServer::ConnectAsync(
	Server s,
	System::String^ host,
	System::Net::NetworkCredential^credentials,
	[System::Runtime::InteropServices::Optional] System::Nullable<int> timeout)
{
	if (!timeout.HasValue)
	{
		timeout = 15000;
	}

	RemoteServerConnectHandler^ handler = gcnew RemoteServerConnectHandler(this, &DAServer::Connect);

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

void OPC::DA::DAServer::ConnectAsync(
	URL url,
	System::Net::NetworkCredential^ credentials,
	[System::Runtime::InteropServices::Optional] System::Nullable<int> timeout)
{
	if (!timeout.HasValue)
	{
		timeout = 15000;
	}

	RemoteURLConnectHandler^ handler = gcnew RemoteURLConnectHandler(this, &DAServer::Connect);

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

System::Collections::Generic::IList<ItemDefinition^>^ OPC::DA::DAServer::Browse(System::String^ nodeToBrowse, BrowseType browseType)
{
	HRESULT hr = S_OK;
	IOPCBrowseServerAddressSpace* browser = GetBrowser();

	if (!System::String::IsNullOrWhiteSpace(nodeToBrowse))
	{
		LPWSTR lNodeToBrowse = (LPWSTR)System::Runtime::InteropServices::Marshal::StringToCoTaskMemUni(nodeToBrowse).ToPointer();

		hr = browser->ChangeBrowsePosition(
			OPCBROWSEDIRECTION::OPC_BROWSE_TO,
			lNodeToBrowse);

		System::Runtime::InteropServices::Marshal::ZeroFreeGlobalAllocUnicode((System::IntPtr) lNodeToBrowse);
	}

	if (hr != S_OK)
	{
		throw gcnew System::ComponentModel::Win32Exception(
			hr,
			System::String::Format(
				"Given node {0} does not exist within the current address space and cannot be browsed.",
				nodeToBrowse));
	}

	OPCBROWSETYPE oBrowseType;

	switch (browseType)
	{
	case BrowseType::Flat:
		oBrowseType = OPCBROWSETYPE::OPC_FLAT;
		break;
	case BrowseType::Branch:
		oBrowseType = OPCBROWSETYPE::OPC_BRANCH;
		break;
	case BrowseType::Leaf:
		oBrowseType = OPCBROWSETYPE::OPC_LEAF;
		break;
	case BrowseType::Items:
	default:

		if (browser != NULL)
		{
			browser->Release();
			browser = NULL;
		}

		throw gcnew System::InvalidOperationException("OPC Browse type has to be either Flat, Branch or Leaf.");
	}

	IEnumString* enumerator = NULL;

	hr = browser->BrowseOPCItemIDs(
		oBrowseType,
		L"",
		VT_EMPTY,
		OPC_BROWSE_FILTER_ALL,
		&enumerator);

	if ((hr != S_OK && hr != S_FALSE) || enumerator == NULL)
	{
		goto error;
	}

	System::Collections::Generic::List<ItemDefinition^>^ definitions = gcnew System::Collections::Generic::List<ItemDefinition^>();

	ULONG itemsToFetch = 100;
	LPOLESTR* items = (LPOLESTR*)CoTaskMemAlloc(itemsToFetch * sizeof(LPOLESTR));
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

			def->SetName(System::Runtime::InteropServices::Marshal::PtrToStringUni((System::IntPtr)it));

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

	throw gcnew System::InvalidOperationException("Could not get items from the address space.", gcnew System::ComponentModel::Win32Exception(hr));
}

delegate System::Collections::Generic::IList<ItemDefinition^>^ BrowseHandler(
	System::String^ nodeToBrowse,
	BrowseType browseType);

System::Collections::Generic::IList<ItemDefinition^>^ OPC::DA::DAServer::BrowseAsync(
	System::String^ nodeToBrowse, 
	BrowseType browseType, 
	[System::Runtime::InteropServices::Optional] System::Nullable<int> timeout)
{
	if (!timeout.HasValue)
	{
		timeout = 15000;
	}

	BrowseHandler^ handler = gcnew BrowseHandler(this, &DAServer::Browse);

	System::IAsyncResult^ result = handler->BeginInvoke(nodeToBrowse, browseType, nullptr, nullptr);

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

Subscription^ OPC::DA::DAServer::AddSubscription(System::String^ sName, System::UInt32 updateRate, float deadBand, bool activated)
{
	IOPCServer* client = GetServer();

	if (client == NULL)
	{
		throw gcnew System::InvalidOperationException("DA Client instance is not connected to any OPC DA server.");
	}

	HRESULT hr = E_FAIL;

	LPWSTR szName = (LPWSTR)System::Runtime::InteropServices::Marshal::StringToCoTaskMemUni(sName).ToPointer();

	System::UInt32 cHandle = System::Threading::Interlocked::Increment((int%)cCounter);
	DWORD sHandle = 0;

	DWORD revisedUpdateRate = 0;

	IOPCItemMgt* itemMgt = NULL;

	hr = client->AddGroup(
		szName,
		activated,
		updateRate,
		cHandle,
		NULL,
		&deadBand,
		NULL,
		&sHandle,
		&revisedUpdateRate,
		IID_IOPCItemMgt,
		(LPUNKNOWN*)&itemMgt);

	if (hr != S_OK)
	{
		if (itemMgt != NULL)
		{
			itemMgt->Release();
			itemMgt = NULL;
		}

		throw gcnew System::ComponentModel::Win32Exception(hr);
	}

	ComBase::SetProxyBlanket(
		itemMgt);

	System::Runtime::InteropServices::Marshal::ZeroFreeCoTaskMemUnicode((System::IntPtr)szName);

	Subscription^ subscription = gcnew Subscription(sName, cHandle, sHandle, revisedUpdateRate, deadBand, this, itemMgt);

	subscriptions->Add(subscription->ClientHandle, subscription);

	return subscription;
}

void OPC::DA::DAServer::RemoveSubscription(Subscription^ subscriptionToRemove)
{
	IOPCServer* client = GetServer();

	if (client == NULL)
	{
		throw gcnew System::InvalidOperationException("DA Client instance is not connected to any OPC DA server.");
	}

	HRESULT hr = E_FAIL;

	hr = client->RemoveGroup(
		subscriptionToRemove->ServerHandle, 
		true);

	if (hr != S_OK)
		throw gcnew System::ComponentModel::Win32Exception(hr);

	subscriptions->Remove(subscriptionToRemove->ClientHandle);

	subscriptionToRemove->Release();
	subscriptionToRemove = nullptr;
}

delegate Subscription^ AddSubscriptionHandler(System::String^ sName, System::UInt32 updateRate, float deadBand, bool activated);

Subscription^ OPC::DA::DAServer::AddSubscriptionAsync(
	System::String^ sName, 
	System::UInt32 updateRate, 
	float deadBand, 
	bool activated, 
	[System::Runtime::InteropServices::Optional] System::Nullable<int> timeout)
{
	if (!timeout.HasValue)
	{
		timeout = 15000;
	}

	AddSubscriptionHandler^ handler = gcnew AddSubscriptionHandler(this, &DAServer::AddSubscription);

	System::IAsyncResult^ result = handler->BeginInvoke(sName, updateRate, deadBand, activated, nullptr, nullptr);

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

delegate void RemoveSubscriptionHandler(Subscription^ subscriptionToRemove);

void OPC::DA::DAServer::RemoveSubscriptionAsync(Subscription^ subscriptionToRemove, [System::Runtime::InteropServices::Optional] System::Nullable<int> timeout)
{
	if (!timeout.HasValue)
	{
		timeout = 15000;
	}

	RemoveSubscriptionHandler^ handler = gcnew RemoveSubscriptionHandler(this, &DAServer::RemoveSubscription);

	System::IAsyncResult^ result = handler->BeginInvoke(subscriptionToRemove, nullptr, nullptr);

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

void OPC::DA::DAServer::Release()
{	
	array<System::UInt32>^ keys = gcnew array<System::UInt32>(this->subscriptions->Keys->Count);

	this->subscriptions->Keys->CopyTo(keys, 0);

	System::Collections::IEnumerator^ e = keys->GetEnumerator();

	while (e->MoveNext())
	{
		System::UInt32 key = (System::UInt32) e->Current;

		Subscription^ s = this->subscriptions[key];

		this->RemoveSubscription(s);
	}	

	if (this->m_IOPCAddressSpaceBrowser != NULL)
	{
		this->m_IOPCAddressSpaceBrowser->Release();
		m_IOPCAddressSpaceBrowser = NULL;
	}

	if (this->m_shutdown != NULL)
	{
		this->m_shutdown->Release();
		delete this->m_shutdown;
		m_shutdown = NULL;
	}

	if(handle.IsAllocated)
		handle.Free();

	ComBase::Release();
}