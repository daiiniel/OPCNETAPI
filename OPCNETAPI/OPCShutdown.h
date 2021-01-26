#pragma once

#include "opccomn.h"

#include "DAServer.h"


class OPCShutdown : public IOPCShutdown
{
	DWORD refCounter;

	DWORD cookie;

	void* serverPtr;

public:

	OPCShutdown()
	{
		this->refCounter = 0;
		this->cookie = 0;

		this->serverPtr = NULL;
	}

	~OPCShutdown()
	{
		this->Release();
	}

	HRESULT Advise(IConnectionPoint* cPoint, System::IntPtr ptr)
	{
		if (cPoint == NULL)
			return E_NOT_SET;

		this->serverPtr = ptr.ToPointer();

		return cPoint->Advise(
			(IUnknown*) this,
			&this->cookie);
	}

	virtual HRESULT __stdcall ShutdownRequest(LPCWSTR szReason) override
	{
		if (serverPtr != NULL)
		{
			OPC::DA::DAServer^ server = (DAServer^)System::Runtime::InteropServices::GCHandle::FromIntPtr(System::IntPtr(serverPtr)).Target;

			server->OnServerShutDownFunc(szReason);

			return S_OK;
		}

		return E_NOT_SET;
	}

	virtual HRESULT __stdcall QueryInterface(REFIID riid, _COM_Outptr_ void __RPC_FAR *__RPC_FAR *ppvObject) override
	{
		if (riid == IID_IUnknown)
			*ppvObject = (IUnknown*)this;
		else if (riid == IID_IOPCShutdown)
			*ppvObject = (IOPCShutdown*)this;
		else
		{
			*ppvObject = NULL;
			return E_NOINTERFACE;
		}

		this->AddRef();

		return S_OK;
	}

	virtual ULONG __stdcall AddRef(void) override
	{
		return ++this->refCounter;
	}

	virtual ULONG __stdcall Release(void) override
	{
		ULONG currentCount = --this->refCounter;

		if (currentCount == 0)
		{
			this->cookie = 0;
			this->serverPtr = NULL;
		}

		return currentCount;
	}
};