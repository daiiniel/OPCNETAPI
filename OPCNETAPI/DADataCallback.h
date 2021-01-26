#pragma once

#include "opcda.h"
#include "opccomn.h"

#include "Subscription.h"

class DADataCallback : IOPCDataCallback
{
	DWORD refCounter;
	DWORD cookie;

	void* subscriptionPtr;
	
public:

	DADataCallback();
	~DADataCallback();

	HRESULT Advise(
		System::IntPtr ptr,
		IConnectionPoint* cPoint);

	virtual HRESULT __stdcall OnDataChange(
		DWORD dwTransid,
		OPCHANDLE hGroup,
		HRESULT hrMasterquality,
		HRESULT hrMastererror,
		DWORD dwCount,
		OPCHANDLE *phClientItems,
		VARIANT *pvValues,
		WORD *pwQualities,
		FILETIME *pftTimeStamps,
		HRESULT *pErrors) override;

	virtual HRESULT __stdcall OnReadComplete(
		DWORD dwTransid,
		OPCHANDLE hGroup,
		HRESULT hrMasterquality,
		HRESULT hrMastererror,
		DWORD dwCount,
		OPCHANDLE *phClientItems,
		VARIANT *pvValues,
		WORD *pwQualities,
		FILETIME *pftTimeStamps,
		HRESULT *pErrors) override;

	virtual HRESULT __stdcall OnWriteComplete(
		DWORD dwTransid,
		OPCHANDLE hGroup,
		HRESULT hrMastererr,
		DWORD dwCount,
		OPCHANDLE *pClienthandles,
		HRESULT *pErrors) override;

	virtual HRESULT __stdcall OnCancelComplete(
		DWORD dwTransid,
		OPCHANDLE hGroup) override;

	virtual HRESULT __stdcall QueryInterface(
		REFIID riid,
		_COM_Outptr_ void __RPC_FAR *__RPC_FAR *ppvObject) override;

	virtual ULONG __stdcall AddRef(void) override;

	virtual ULONG __stdcall Release(void) override;
};

