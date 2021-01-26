#include "stdafx.h"
#include "Subscription.h"


HRESULT Subscription::OnDataChangedFunc(
	DWORD dwTransid, 
	OPCHANDLE hGroup, 
	HRESULT hrMasterquality, 
	HRESULT hrMastererror, 
	DWORD dwCount, 
	OPCHANDLE * phClientItems, 
	VARIANT * pvValues, 
	WORD * pwQualities, 
	FILETIME * pftTimeStamps, 
	HRESULT * pErrors)
{
	if (DataChangedHandler != nullptr)
	{
		DataChangedEventArgs^ e = gcnew DataChangedEventArgs();

		e->TransactionId = dwTransid;
		e->Results = gcnew System::Collections::Generic::List<DAReadResult>();

		for (DWORD ii = 0; ii < dwCount; ii++)
		{
			ItemDefinition^ definition;
			DAReadResult r;

			this->items->TryGetValue(phClientItems[ii], definition);

			r.SetName(definition->Name);
			r.SetItemID(definition->ItemID);
			r.SetResult((Results)pErrors[ii]);

			r.SetClientHandle(definition->ClientHandle);
			r.SetServerHandle(definition->ServerHandle);

			if (pErrors[ii] == S_OK)
			{
				ReadResult rr;

				rr.SetValue(VariantToObject(pvValues[ii]));
				rr.SetQuality((Qualities)pwQualities[ii]);
				rr.SetTimeStamp(FileTimeToDateTime(pftTimeStamps[ii]));

				VariantClear(&pvValues[ii]);

				r.SetValue(rr);
			}

			e->Results->Add(r);
		}

		for each(System::Delegate^ del in DataChangedHandler->GetInvocationList())
		{
			del->DynamicInvoke(
				this,
				e);
		}
	}

	return S_OK;
}

HRESULT Subscription::OnReadCompletedFunc(
	DWORD dwTransid, 
	OPCHANDLE hGroup, 
	HRESULT hrMasterquality, 
	HRESULT hrMastererror, 
	DWORD dwCount, 
	OPCHANDLE * phClientItems, 
	VARIANT * pvValues, 
	WORD * pwQualities, 
	FILETIME * pftTimeStamps, 
	HRESULT * pErrors)
{
	if (ReadCompletedHandler != nullptr)
	{
		DataChangedEventArgs^ e = gcnew DataChangedEventArgs();

		e->TransactionId = dwTransid;
		e->Results = gcnew System::Collections::Generic::List<DAReadResult>();

		for (DWORD ii = 0; ii < dwCount; ii++)
		{
			ItemDefinition^ definition;
			DAReadResult r;

			this->items->TryGetValue(phClientItems[ii], definition);

			r.SetName(definition->Name);
			r.SetItemID(definition->ItemID);
			r.SetResult((Results)pErrors[ii]);

			r.SetClientHandle(definition->ClientHandle);
			r.SetServerHandle(definition->ServerHandle);

			if (pErrors[ii] == S_OK)
			{
				ReadResult rr;

				rr.SetValue(VariantToObject(pvValues[ii]));
				rr.SetQuality((Qualities)pwQualities[ii]);
				rr.SetTimeStamp(FileTimeToDateTime(pftTimeStamps[ii]));

				VariantClear(&pvValues[ii]);

				r.SetValue(rr);
			}

			e->Results->Add(r);
		}

		for each(System::Delegate^ del in ReadCompletedHandler->GetInvocationList())
		{
			del->DynamicInvoke(
				this,
				e);
		}
	}

	return S_OK;
}

HRESULT Subscription::OnWriteCompletedFunc(
	DWORD dwTransid, 
	OPCHANDLE hGroup, 
	HRESULT hrMastererr, 
	DWORD dwCount, 
	OPCHANDLE * pClienthandles, 
	HRESULT * pErrors)
{
	if (WriteCompletedHandler != nullptr)
	{
		for each(System::Delegate^ del in WriteCompletedHandler->GetInvocationList())
		{
			WriteCompletedEventArgs^ e = gcnew WriteCompletedEventArgs();

			e->TransactionId = dwTransid;

			del->DynamicInvoke(
				this,
				e);
		}
	}

	return S_OK;
}

HRESULT OPC::DA::Subscription::OnCancelCompleted(
	DWORD dwTransid, 
	OPCHANDLE hGroup)
{
	if (CancelCompletedHandler != nullptr)
	{
		for each(System::Delegate^ del in CancelCompletedHandler->GetInvocationList())
		{
			CancelCompletedEventArgs^ e = gcnew CancelCompletedEventArgs();

			e->TransactionId = dwTransid;

			del->DynamicInvoke(
				this,
				e);
		}
	}

	return S_OK;
}

OPC::DA::Subscription::Subscription(
	System::String^ name, 
	System::UInt32 cHandle, 
	System::UInt32 sHandle,
	System::UInt32 uRate, 
	float dBand,
	ComBase^ pParent, 
	IUnknown* ppIn) : ComBase(
		pParent,
		ppIn), cCounter(1)
{
	SetName(name);
	SetClientHandle(cHandle);
	SetServerHandle(sHandle);

	IOPCItemMgt* client = GetServer();

	m_IOPCSyncIO = NULL;
	m_IOPCASyncIO = NULL;

	if (client == NULL)
	{
		throw gcnew System::InvalidOperationException("DA Client instance is not connected to any OPC DA server.");
	}

	HRESULT hr = E_FAIL;

	IConnectionPointContainer* container = NULL;

	hr = this->GetInterface(
		__uuidof(IConnectionPointContainer),
		(void**)&container);

	if (hr != S_OK)
	{
		if (container != NULL)
		{
			container->Release();
		}
	}

	IConnectionPoint* connectionPoint = NULL;

	hr = container->FindConnectionPoint(
		IID_IOPCDataCallback,
		&connectionPoint);

	if (hr != S_OK)
	{
		if (connectionPoint != NULL)
		{
			connectionPoint->Release();
		}

		if (container != NULL)
		{
			container->Release();
		}
		
		throw gcnew System::ComponentModel::Win32Exception(hr);
	}

	hr = ComBase::SetProxyBlanket(
		connectionPoint);

	dataCallback = new DADataCallback();

	handle = System::Runtime::InteropServices::GCHandle::Alloc(this);
	
	HRESULT cHr = dataCallback->Advise(
		System::Runtime::InteropServices::GCHandle::ToIntPtr(this->handle),
		connectionPoint);

	if (connectionPoint != NULL)
	{
		connectionPoint->Release();
	}

	if (container != NULL)
	{
		container->Release();
	}

	items = gcnew System::Collections::Generic::Dictionary<System::UInt32, ItemDefinition^>();
}

IOPCItemMgt* Subscription::GetServer()
{
	if (server == NULL)
	{
		throw gcnew System::InvalidOperationException("DA Client instance is not connected to any OPC DA server.");
	}

	return (IOPCItemMgt*)server;
}
IOPCSyncIO2* Subscription::GetSyncIO()
{
	if (m_IOPCSyncIO != NULL)
	{
		return m_IOPCSyncIO;
	}

	IOPCItemMgt* client = GetServer();	

	HRESULT hr = E_FAIL;

	IOPCSyncIO2* syncIO = NULL;

	hr = ComBase::GetInterface(
		__uuidof(IOPCSyncIO2),
		(void**)&syncIO);

	if (hr != S_OK)
	{
		if (syncIO != NULL)
		{
			syncIO->Release();
		}

		throw gcnew System::ComponentModel::Win32Exception(hr);
	}

	m_IOPCSyncIO = syncIO;

	return m_IOPCSyncIO;
}
IOPCAsyncIO2* Subscription::GetAsyncIO()
{
	if (m_IOPCASyncIO != NULL)
	{
		return m_IOPCASyncIO;
	}

	IOPCItemMgt* client = GetServer();

	HRESULT hr = E_FAIL;

	IOPCAsyncIO2* asyncIO = NULL;

	hr = ComBase::GetInterface(
		__uuidof(IOPCAsyncIO2),
		(void**)&asyncIO);

	if (hr != S_OK)
	{
		if (asyncIO != NULL)
		{
			asyncIO->Release();
		}

		throw gcnew System::ComponentModel::Win32Exception(hr);
	}

	m_IOPCASyncIO = asyncIO;

	return m_IOPCASyncIO;
}
IOPCGroupStateMgt* Subscription::GetGroupStateMgt()
{
	if (m_IOPCGroupStateMgt != NULL)
	{
		return m_IOPCGroupStateMgt;
	}

	IOPCItemMgt* client = GetServer();

	HRESULT hr = E_FAIL;

	IOPCGroupStateMgt* groupStateMgt = NULL;

	hr = ComBase::GetInterface(
		IID_IOPCGroupStateMgt2,
		(void**)&groupStateMgt);

	if (hr != S_OK)
	{
		if (groupStateMgt != NULL)
		{
			groupStateMgt->Release();
		}

		throw gcnew System::ComponentModel::Win32Exception(hr);
	}

	m_IOPCGroupStateMgt = groupStateMgt;

	return m_IOPCGroupStateMgt;
}

void Subscription::ChangeState(DWORD* updateRate, bool* status, float* deadband)
{		
	HRESULT hr = E_FAIL;

	IOPCGroupStateMgt* statMgt = GetGroupStateMgt();

	DWORD revisedUpdateRate = 0;

	hr = statMgt->SetState(
		(updateRate != NULL && *updateRate != 0) ? updateRate : NULL,
		&revisedUpdateRate,
		(BOOL*)status,
		NULL,
		(deadband != NULL) ? deadband : NULL,
		NULL,
		NULL);

	if (hr != S_OK && hr != S_FALSE)
	{
		throw gcnew System::ComponentModel::Win32Exception(hr);
	}

	SetUpdateRate(revisedUpdateRate);

	if (deadband != NULL)
		SetDeadBand(*deadband);

	if (status != NULL)
		active = *status;
}

delegate void definitionsHandler(
	System::Collections::Generic::IList<ItemDefinition^>^ definitionsToAdd);

void OPC::DA::Subscription::AddItemsAsync(
	System::Collections::Generic::IList<ItemDefinition^>^ definitionsToAdd, 
	[System::Runtime::InteropServices::Optional] System::Nullable<int> timeout)
{
	if (!timeout.HasValue)
		timeout = 15000;

	definitionsHandler^handler = gcnew definitionsHandler(this, &OPC::DA::Subscription::AddItems);

	System::IAsyncResult^ result = handler->BeginInvoke(definitionsToAdd, nullptr, nullptr);

	if (!result->AsyncWaitHandle->WaitOne(timeout.Value, true))
	{
		throw gcnew System::TimeoutException();
	}

	if (!result->IsCompleted)
		throw gcnew System::Exception("Could not add items.");

	try
	{
		handler->EndInvoke(result);
	}
	catch (System::Exception^)
	{
		throw;
	}
}

void Subscription::AddItems(
	System::Collections::Generic::IList<ItemDefinition^>^ definitionsToAdd)
{
	IOPCItemMgt* client = GetServer();

	if (client == NULL)
	{
		throw gcnew System::InvalidOperationException("DA Client instance is not connected to any OPC DA server.");
	}

	HRESULT hr = E_FAIL;

	DWORD count = definitionsToAdd->Count;

	OPCITEMDEF* itemsToAdd = (OPCITEMDEF*)CoTaskMemAlloc(sizeof(OPCITEMDEF) * count);

	for (DWORD ii = 0; ii < count; ii++)
	{
		itemsToAdd[ii].szItemID = (LPWSTR)System::Runtime::InteropServices::Marshal::StringToCoTaskMemUni(definitionsToAdd[ii]->ItemID).ToPointer();
		itemsToAdd[ii].bActive = true;

		itemsToAdd[ii].szAccessPath = L"";
		itemsToAdd[ii].vtRequestedDataType = 0;
		itemsToAdd[ii].hClient = System::Threading::Interlocked::Increment((int%)cCounter);
		itemsToAdd[ii].pBlob = NULL;
		itemsToAdd[ii].dwBlobSize = 0;
		itemsToAdd[ii].wReserved = 0;
	}

	OPCITEMRESULT* results = NULL;
	HRESULT* errors = NULL;

	hr = client->AddItems(
		count,
		itemsToAdd,
		&results,
		&errors);

	if (hr != S_OK && hr != S_FALSE)
	{
		if (itemsToAdd != NULL)
		{
			for (DWORD ii = 0; ii < count; ii++)
			{
				System::Runtime::InteropServices::Marshal::ZeroFreeCoTaskMemUnicode((System::IntPtr)itemsToAdd[ii].szItemID);
			}

			CoTaskMemFree(itemsToAdd);
		}

		if (results != NULL)
		{
			CoTaskMemFree(results);
		}

		if (errors != NULL)
		{
			CoTaskMemFree(errors);
		}

		throw gcnew System::ComponentModel::Win32Exception(hr);
	}

	for (DWORD ii = 0; ii < count; ii++)
	{
		System::Runtime::InteropServices::Marshal::ZeroFreeCoTaskMemUnicode((System::IntPtr)itemsToAdd[ii].szItemID);
				
		definitionsToAdd[ii]->SetResult((Results)errors[ii]);

		if (errors[ii] == S_OK)
		{
			definitionsToAdd[ii]->SetClientHandle(itemsToAdd[ii].hClient);
			definitionsToAdd[ii]->SetServerHandle(results[ii].hServer);

			items->Add(definitionsToAdd[ii]->ClientHandle, definitionsToAdd[ii]);
		}
	}

	if (itemsToAdd != NULL)
	{
		CoTaskMemFree(itemsToAdd);
	}

	if (results != NULL)
	{
		CoTaskMemFree(results);
	}

	if (errors != NULL)
	{
		CoTaskMemFree(errors);
	}
}

void Subscription::RemoveItemsAsync(
	System::Collections::Generic::IList<ItemDefinition^>^ definitionsToRemove, 
	[System::Runtime::InteropServices::Optional] System::Nullable<int> timeout)
{
	if (!timeout.HasValue)
		timeout = 15000;

	definitionsHandler^handler = gcnew definitionsHandler(this, &OPC::DA::Subscription::RemoveItems);

	System::IAsyncResult^ result = handler->BeginInvoke(definitionsToRemove, nullptr, nullptr);

	if (!result->AsyncWaitHandle->WaitOne(timeout.Value, true))
	{
		throw gcnew System::TimeoutException();
	}

	if (!result->IsCompleted)
		throw gcnew System::Exception("Could not add items.");

	try
	{
		handler->EndInvoke(result);
	}
	catch (System::Exception^)
	{
		throw;
	}
}

void Subscription::RemoveItems(
	System::Collections::Generic::IList<ItemDefinition^>^ definitionsToRemove)
{
	IOPCItemMgt* client = GetServer();

	if (client == NULL)
	{
		throw gcnew System::InvalidOperationException("DA Client instance is not connected to any OPC DA server.");
	}

	HRESULT hr = E_FAIL;

	DWORD count = definitionsToRemove->Count;

	DWORD* itemsToRemove = (DWORD*)CoTaskMemAlloc(sizeof(DWORD) * count);

	for (DWORD ii = 0; ii < count; ii++)
	{
		itemsToRemove[ii] = definitionsToRemove[ii]->ServerHandle;
	}

	HRESULT* errors = NULL;

	hr = client->RemoveItems(
		count,
		itemsToRemove,
		&errors);

	if (hr != S_OK && hr != S_FALSE)
	{
		if (errors != NULL)
		{
			CoTaskMemFree(errors);
		}

		if (itemsToRemove != NULL)
		{
			CoTaskMemFree(itemsToRemove);
		}

		throw gcnew System::ComponentModel::Win32Exception(hr);
	}

	for (DWORD ii = 0; ii < count; ii++)
	{
		if (errors[ii] == S_OK)
		{
			items->Remove(definitionsToRemove[ii]->ClientHandle);
		}

		definitionsToRemove[ii]->SetClientHandle(0);
		definitionsToRemove[ii]->SetServerHandle(0);
		definitionsToRemove[ii]->SetResult((Results)errors[ii]);
	}

	if (errors != NULL)
	{
		CoTaskMemFree(errors);
	}

	if (itemsToRemove != NULL)
	{
		CoTaskMemFree(itemsToRemove);
	}
}

Subscription::~Subscription()
{
	if (this->items->Count > 0)
	{	
		System::Collections::Generic::IList<ItemDefinition^>^ itemsToRemove = gcnew System::Collections::Generic::List<ItemDefinition^>(this->items->Count);

		for each(ItemDefinition^ def in this->items->Values)
		{
			itemsToRemove->Add(def);
		}

		this->RemoveItems(itemsToRemove);
	}


	if (dataCallback != NULL)
	{
		dataCallback->Release();
		delete dataCallback;

		dataCallback = NULL;
	}
				
	this->Release();
}

System::Collections::Generic::IList<DAReadResult>^ Subscription::Read(
	System::Collections::Generic::IList<ItemDefinition^>^ definitionsToRead,
	DataSources source)
{
	IOPCSyncIO2* syncIO = GetSyncIO();

	DWORD count = definitionsToRead->Count;

	if (count == 0)
		return gcnew System::Collections::Generic::List<DAReadResult>();

	DWORD* phServer = (DWORD*)CoTaskMemAlloc(count * sizeof(DWORD));

	for (DWORD ii = 0; ii < count; ii++)
	{
		phServer[ii] = definitionsToRead[ii]->ServerHandle;
	}

	OPCITEMSTATE* states = NULL;
	HRESULT* errors = NULL;

	OPCDATASOURCE OPCSource;

	switch (source)
	{
	case OPC::DA::DataSources::Device:
		OPCSource = OPC_DS_DEVICE;
		break;
	case OPC::DA::DataSources::Cache:
		OPCSource = OPC_DS_CACHE;
		break;
	default:
		break;
	}

	HRESULT hr = syncIO->Read(
		OPCSource,
		count,
		phServer,
		&states,
		&errors);

	if (hr != S_OK && hr != S_FALSE)
	{
		goto error;
	}

	System::Collections::Generic::IList<DAReadResult>^ results = gcnew System::Collections::Generic::List<DAReadResult>();

	for (DWORD ii = 0; ii < count; ii++)
	{
		DAReadResult r;

		r.SetName(definitionsToRead[ii]->Name);
		r.SetItemID(definitionsToRead[ii]->ItemID);
		r.SetResult((Results)errors[ii]);

		r.SetClientHandle(definitionsToRead[ii]->ClientHandle);
		r.SetServerHandle(definitionsToRead[ii]->ServerHandle);

		if (errors[ii] == S_OK)
		{
			ReadResult rr;

			rr.SetValue(VariantToObject(states[ii].vDataValue));
			rr.SetQuality((Qualities)states[ii].wQuality);
			rr.SetTimeStamp(FileTimeToDateTime(states[ii].ftTimeStamp));

			VariantClear(&states[ii].vDataValue);

			r.SetValue(rr);
		}

		results->Add(r);
	}

	if (errors != NULL)
	{
		CoTaskMemFree(errors);
		errors = NULL;
	}

	if (states != NULL)
	{
		CoTaskMemFree(states);
		states = NULL;
	}

	if (phServer != NULL)
	{
		CoTaskMemFree(phServer);
		phServer = NULL;
	}

	return results;

error:

	if (errors != NULL)
	{
		CoTaskMemFree(errors);
		errors = NULL;
	}

	if (states != NULL)
	{
		CoTaskMemFree(states);
		states = NULL;
	}

	if (phServer != NULL)
	{
		CoTaskMemFree(phServer);
		phServer = NULL;
	}

	throw gcnew System::ComponentModel::Win32Exception(hr);
}

System::Collections::Generic::IList<DAResult^>^ Subscription::ReadAsync(
	System::Collections::Generic::IList<ItemDefinition^>^ definitionsToRead,
	DataSources source,
	System::UInt32 transactionId,
	[System::Runtime::InteropServices::Out] System::UInt32% cancelId)
{
	IOPCAsyncIO2* asyncIO = GetAsyncIO();

	HRESULT hr;

	DWORD count = definitionsToRead->Count;
	OPCHANDLE* handles = (OPCHANDLE*)CoTaskMemAlloc(count * sizeof(OPCHANDLE));
	DWORD dCancelId;
	HRESULT* results = NULL;

	System::Collections::Generic::List<DAResult^>^ res = gcnew System::Collections::Generic::List<DAResult^>(count);

	for (DWORD ii = 0; ii < count; ii++)
	{
		handles[ii] = definitionsToRead[ii]->ServerHandle;
		res->Insert(ii, (DAResult^) definitionsToRead[ii]);
	}

	hr = asyncIO->Read(
		count,
		handles,
		(DWORD)transactionId,
		&dCancelId,
		&results);

	cancelId = dCancelId;

	if (hr != S_OK && hr != S_FALSE)
		goto error;

	for (DWORD ii = 0; ii < count; ii++)
	{
		res[ii]->SetResult((Results) results[ii]);
	}

	if (handles != NULL)
	{
		CoTaskMemFree(handles);
		handles = NULL;
	}

	if (results != NULL)
	{
		CoTaskMemFree(results);
		handles = NULL;
	}

	return res;

error:

	if (handles != NULL)
	{
		CoTaskMemFree(handles);
		handles = NULL;
	}

	if (results != NULL)
	{
		CoTaskMemFree(results);
		handles = NULL;
	}

	throw gcnew System::ComponentModel::Win32Exception(hr);
}


System::Collections::Generic::IList<DAResult^>^ Subscription::Write(
	System::Collections::Generic::IList<ItemValueDefinition^>^ pairItemValuesToWrite)
{
	HRESULT hr;

	IOPCSyncIO2* syncIO = GetSyncIO();

	DWORD count = pairItemValuesToWrite->Count;
	OPCHANDLE* handles = (OPCHANDLE*) CoTaskMemAlloc(count * sizeof(OPCHANDLE));
	VARIANT* values = (VARIANT*)CoTaskMemAlloc(count * sizeof(VARIANT));

	HRESULT* results = nullptr;

	System::Collections::Generic::List<DAResult^>^ res = gcnew System::Collections::Generic::List<DAResult^>(count);

	for (DWORD ii = 0; ii < count; ii++)
	{
		handles[ii] = pairItemValuesToWrite[ii]->ServerHandle;
		values[ii] = ObjectToVariant(pairItemValuesToWrite[ii]->Value);

		res->Insert(ii, (DAResult^) pairItemValuesToWrite[ii]);
	}		

	hr = syncIO->Write(
		count,
		handles,
		values,
		&results);

	if (hr != S_OK && hr != S_FALSE)
		goto error;

	for (DWORD ii = 0; ii < count; ii++)
	{
		VariantClear(&values[ii]);

		res[ii]->SetResult((Results) results[ii]);
	}

	if (handles != NULL)
	{
		CoTaskMemFree(handles);
		handles = NULL;
	}

	if (values == NULL)
	{		
		CoTaskMemFree(values);
		values = NULL;
	}

	return res;

error:
	
	if (handles != NULL)
	{
		CoTaskMemFree(handles);
		handles = NULL;
	}

	if (values == NULL)
	{
		for (DWORD ii = 0; ii < count; ii++)
		{
			VariantClear(&values[ii]);
		}

		CoTaskMemFree(values);
		values = NULL;
	}

	throw gcnew System::ComponentModel::Win32Exception(hr);
}

System::Collections::Generic::IList<DAResult^>^ Subscription::WriteAsync(
	System::Collections::Generic::IList<ItemValueDefinition^>^ pairItemValuesToWrite,
	System::UInt32 transactionId, 
	[System::Runtime::InteropServices::Out] System::UInt32% cancelId)
{
	IOPCAsyncIO2* asyncIO = GetAsyncIO();

	HRESULT hr;

	DWORD count = pairItemValuesToWrite->Count;
	OPCHANDLE* handles = (OPCHANDLE*)CoTaskMemAlloc(count * sizeof(OPCHANDLE));
	VARIANT* values = (VARIANT*)CoTaskMemAlloc(count * sizeof(VARIANT));

	System::Collections::Generic::List<DAResult^>^ res = gcnew System::Collections::Generic::List<DAResult^>(count);

	DWORD dCancelId;

	HRESULT* results = nullptr;

	for (DWORD ii = 0; ii < count; ii++)
	{
		handles[ii] = pairItemValuesToWrite[ii]->ServerHandle;
		values[ii] = ObjectToVariant(pairItemValuesToWrite[ii]->Value);

		res->Insert(ii, (DAResult^)pairItemValuesToWrite[ii]);
	}

	hr = asyncIO->Write(
		count,
		handles,
		values,
		(DWORD)  transactionId,
		(DWORD*) &dCancelId,
		&results);

	cancelId = dCancelId;

	if (hr != S_OK && hr != S_FALSE)
		goto error;

	for (DWORD ii = 0; ii < count; ii++)
	{
		VariantClear(&values[ii]);

		res[ii]->SetResult((Results)results[ii]);
	}

	if (handles != NULL)
	{
		CoTaskMemFree(handles);
		handles = NULL;
	}

	if (values == NULL)
	{
		CoTaskMemFree(values);
		values = NULL;
	}

	if (results != NULL)
	{
		CoTaskMemFree(results);
		handles = NULL;
	}

	return res;

error:

	if (handles != NULL)
	{
		CoTaskMemFree(handles);
		handles = NULL;
	}

	if (values == NULL)
	{
		for (DWORD ii = 0; ii < count; ii++)
		{
			VariantClear(&values[ii]);
		}

		CoTaskMemFree(values);
		values = NULL;
	}

	if (results != NULL)
	{
		CoTaskMemFree(results);
		handles = NULL;
	}

	throw gcnew System::ComponentModel::Win32Exception(hr);
}

void Subscription::Refresh(
	DataSources source,
	System::UInt32 transactionId,
	[System::Runtime::InteropServices::OutAttribute] System::UInt32% cancelId)
{
	IOPCAsyncIO2* asyncIO = GetAsyncIO();

	HRESULT hr;

	OPCDATASOURCE OPCDataSource;
	DWORD dCancelId;
	
	switch (source)
	{
	case OPC::DA::DataSources::Device:
		OPCDataSource = OPC_DS_DEVICE;
		break;
	case OPC::DA::DataSources::Cache:
		OPCDataSource = OPC_DS_CACHE;
		break;
	default:
		break;
	}

	hr = asyncIO->Refresh2(
		OPCDataSource,
		(DWORD)transactionId,
		(DWORD*) &dCancelId);

	cancelId = dCancelId;

	if (hr != S_OK && hr != S_FALSE)
		throw gcnew System::ComponentModel::Win32Exception(hr);
}

void Subscription::Cancel(System::UInt32 cancelId)
{
	IOPCAsyncIO2* asyncIO = GetAsyncIO();

	HRESULT hr;

	hr = asyncIO->Cancel2(
		(DWORD)cancelId);
	
	if (hr != S_OK && hr != S_FALSE)
	{
		throw gcnew System::ComponentModel::Win32Exception(hr);
	}
}

void Subscription::Activate()
{
	HRESULT hr;

	IOPCAsyncIO2* asyncIO = GetAsyncIO();

	hr = asyncIO->SetEnable(true);

	if (hr != S_OK && hr != S_FALSE)
	{
		throw gcnew System::ComponentModel::Win32Exception(hr);
	}
}
void Subscription::DeActivate()
{
	HRESULT hr;

	IOPCAsyncIO2* asyncIO = GetAsyncIO();

	hr = asyncIO->SetEnable(false);

	if (hr != S_OK && hr != S_FALSE)
	{
		throw gcnew System::ComponentModel::Win32Exception(hr);
	}
}

void Subscription::ChangeState(System::UInt32 updateRate, bool status, float deadband)
{
	DWORD dRate = updateRate;

	ChangeState(&dRate, &status, &deadband);
}

void Subscription::Release()
{	
	if (dataCallback != NULL)
	{
		dataCallback->Release();

		delete dataCallback;
		dataCallback = NULL;
	}

	if (m_IOPCASyncIO != NULL)
	{
		m_IOPCASyncIO->Release();
		m_IOPCASyncIO == NULL;
	}

	if (m_IOPCGroupStateMgt != NULL)
	{
		m_IOPCGroupStateMgt->Release();
		m_IOPCGroupStateMgt = NULL;
	}

	if (m_IOPCSyncIO != NULL)
	{
		m_IOPCSyncIO->Release();
		m_IOPCSyncIO = NULL;
	}

	if (this->handle.IsAllocated)
		this->handle.Free();

	ComBase::Release();
}