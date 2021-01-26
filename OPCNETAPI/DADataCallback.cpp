#include "Stdafx.h"
#include "DADataCallback.h"

DADataCallback::DADataCallback() : refCounter(0), cookie(0)
{
	this->refCounter = 0;
	this->cookie = 0;

	this->subscriptionPtr = nullptr;
}


DADataCallback::~DADataCallback()
{
	this->Release();
}

HRESULT DADataCallback::Advise(
	System::IntPtr ptr,
	IConnectionPoint * cPoint)
{
	HRESULT hr;

	if (cPoint == NULL)
		return E_NOT_SET;
	
	this->subscriptionPtr = ptr.ToPointer();

	hr = cPoint->Advise(
		(IUnknown*)this,
		&this->cookie);

	return hr;
}

HRESULT DADataCallback::OnDataChange(
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
	if (!this->subscriptionPtr)
		return E_NOT_SET;

	HRESULT hr = ptrToSubscription(subscriptionPtr)->OnDataChangedFunc(
		dwTransid,
		hGroup,
		hrMasterquality,
		hrMastererror,
		dwCount,
		phClientItems,
		pvValues,
		pwQualities,
		pftTimeStamps,
		pErrors);

	return hr;
}

HRESULT DADataCallback::OnReadComplete(DWORD dwTransid, OPCHANDLE hGroup, HRESULT hrMasterquality, HRESULT hrMastererror, DWORD dwCount, OPCHANDLE * phClientItems, VARIANT * pvValues, WORD * pwQualities, FILETIME * pftTimeStamps, HRESULT * pErrors)
{
	if (!this->subscriptionPtr)
		return E_NOT_SET;

	HRESULT hr = ptrToSubscription(subscriptionPtr)->OnReadCompletedFunc(
		dwTransid,
		hGroup,
		hrMasterquality,
		hrMastererror,
		dwCount,
		phClientItems,
		pvValues,
		pwQualities,
		pftTimeStamps,
		pErrors);	

	return hr;
}

HRESULT DADataCallback::OnWriteComplete(DWORD dwTransid, OPCHANDLE hGroup, HRESULT hrMastererr, DWORD dwCount, OPCHANDLE * pClienthandles, HRESULT * pErrors)
{
	if (!this->subscriptionPtr)
		return E_NOT_SET;
	
	HRESULT hr = ptrToSubscription(subscriptionPtr)->OnWriteCompletedFunc(
		dwTransid,
		hGroup,
		dwCount,
		dwCount,
		pClienthandles,
		pErrors);
		
	return hr;
}

HRESULT DADataCallback::OnCancelComplete(DWORD dwTransid, OPCHANDLE hGroup)
{
	if (!subscriptionPtr)
		return E_NOT_SET;

	HRESULT hr = ptrToSubscription(subscriptionPtr)->OnCancelCompleted(
		dwTransid,
		hGroup);

	return hr;
}

HRESULT DADataCallback::QueryInterface(REFIID riid, void ** ppvObject)
{
	if (riid == IID_IUnknown)
		*ppvObject = (IUnknown*)this;
	else if (riid == IID_IOPCDataCallback)
		*ppvObject = (IOPCDataCallback*)this;
	else
	{
		*ppvObject = NULL;
		return E_NOINTERFACE;
	}

	this->AddRef();

	return S_OK;
}

ULONG DADataCallback::AddRef(void)
{
	return ++this->refCounter;
}

ULONG DADataCallback::Release(void)
{
	ULONG currentCount = --this->refCounter;

	if (currentCount == 0)
	{	
		this->cookie = NULL;
	}

	return currentCount;
}
