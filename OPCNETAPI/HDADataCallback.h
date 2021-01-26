#pragma once

#include "opchda.h"

typedef HRESULT(*DataChangedCallback)(
	/* [in] */ DWORD dwTransactionID,
	/* [in] */ HRESULT hrStatus,
	/* [in] */ DWORD dwNumItems,
	/* [size_is][in] */ OPCHDA_ITEM *pItemValues,
	/* [size_is][in] */ HRESULT *phrErrors);

typedef HRESULT(*ReadCompletedCallback)(
	/* [in] */ DWORD dwTransactionID,
	/* [in] */ HRESULT hrStatus,
	/* [in] */ DWORD dwNumItems,
	/* [size_is][in] */ OPCHDA_ITEM *pItemValues,
	/* [size_is][in] */ HRESULT *phrErrors);

typedef HRESULT(*ReadModifiedCompletedCallback)(
	/* [in] */ DWORD dwTransactionID,
	/* [in] */ HRESULT hrStatus,
	/* [in] */ DWORD dwNumItems,
	/* [size_is][in] */ OPCHDA_MODIFIEDITEM *pItemValues,
	/* [size_is][in] */ HRESULT *phrErrors);

typedef HRESULT (*ReadAttributeCompletedCallback)(
	/* [in] */ DWORD dwTransactionID,
	/* [in] */ HRESULT hrStatus,
	/* [in] */ OPCHANDLE hClient,
	/* [in] */ DWORD dwNumItems,
	/* [size_is][in] */ OPCHDA_ATTRIBUTE *pAttributeValues,
	/* [size_is][in] */ HRESULT *phrErrors);

typedef HRESULT(*ReadAnnotationsCallback)(
	/* [in] */ DWORD dwTransactionID,
	/* [in] */ HRESULT hrStatus,
	/* [in] */ DWORD dwNumItems,
	/* [size_is][in] */ OPCHDA_ANNOTATION *pAnnotationValues,
	/* [size_is][in] */ HRESULT *phrErrors);

typedef HRESULT (*InsertAnnotationsCallback)(
	/* [in] */ DWORD dwTransactionID,
	/* [in] */ HRESULT hrStatus,
	/* [in] */ DWORD dwCount,
	/* [size_is][in] */ OPCHANDLE *phClients,
	/* [size_is][in] */ HRESULT *phrErrors);

typedef HRESULT (*PlaybackCallback)(
	/* [in] */ DWORD dwTransactionID,
	/* [in] */ HRESULT hrStatus,
	/* [in] */ DWORD dwNumItems,
	/* [size_is][in] */ OPCHDA_ITEM **ppItemValues,
	/* [size_is][in] */ HRESULT *phrErrors);

typedef HRESULT (*UpdateCompletedCallback)(
	/* [in] */ DWORD dwTransactionID,
	/* [in] */ HRESULT hrStatus,
	/* [in] */ DWORD dwCount,
	/* [size_is][in] */ OPCHANDLE *phClients,
	/* [size_is][in] */ HRESULT *phrErrors);

typedef HRESULT (*CancelCompletedCallback)(
	/* [in] */ DWORD dwCancelID);

class HDADataCallback : public IOPCHDA_DataCallback
{
	DWORD m_counter;

	DataChangedCallback m_dataChanged;
	ReadCompletedCallback m_readCompleted;
	ReadModifiedCompletedCallback m_readModified;
	ReadAttributeCompletedCallback m_readAttributeCompleted;
	ReadAnnotationsCallback m_readAnnotations;
	InsertAnnotationsCallback m_insertAnnotations;
	PlaybackCallback m_playCallback;
	UpdateCompletedCallback m_updateCompletedCallback;
	CancelCompletedCallback m_cancelCompletedCallback;


public:

	HDADataCallback() : m_counter(0)
	{

	}

#pragma region DataCallback

	virtual HRESULT STDMETHODCALLTYPE OnDataChange(
		/* [in] */ DWORD dwTransactionID,
		/* [in] */ HRESULT hrStatus,
		/* [in] */ DWORD dwNumItems,
		/* [size_is][in] */ OPCHDA_ITEM *pItemValues,
		/* [size_is][in] */ HRESULT *phrErrors)
	{
		if (m_dataChanged != nullptr)
		{
			return m_dataChanged(
				dwTransactionID,
				hrStatus,
				dwNumItems,
				pItemValues,
				phrErrors);
		}

		return E_NOT_SET;
	};

	virtual HRESULT STDMETHODCALLTYPE OnReadComplete(
		/* [in] */ DWORD dwTransactionID,
		/* [in] */ HRESULT hrStatus,
		/* [in] */ DWORD dwNumItems,
		/* [size_is][in] */ OPCHDA_ITEM *pItemValues,
		/* [size_is][in] */ HRESULT *phrErrors)
	{
		if (m_readCompleted != nullptr)
		{
			return m_readCompleted(
				dwTransactionID,
				hrStatus,
				dwNumItems,
				pItemValues,
				phrErrors);
		}

		return E_NOT_SET;
	};

	virtual HRESULT STDMETHODCALLTYPE OnReadModifiedComplete(
		/* [in] */ DWORD dwTransactionID,
		/* [in] */ HRESULT hrStatus,
		/* [in] */ DWORD dwNumItems,
		/* [size_is][in] */ OPCHDA_MODIFIEDITEM *pItemValues,
		/* [size_is][in] */ HRESULT *phrErrors)
	{
		if (m_readModified != nullptr)
		{
			return m_readModified(
				dwTransactionID,
				hrStatus,
				dwNumItems,
				pItemValues,
				phrErrors);
		}

		return E_NOT_SET;
	}

	virtual HRESULT STDMETHODCALLTYPE OnReadAttributeComplete(
		/* [in] */ DWORD dwTransactionID,
		/* [in] */ HRESULT hrStatus,
		/* [in] */ OPCHANDLE hClient,
		/* [in] */ DWORD dwNumItems,
		/* [size_is][in] */ OPCHDA_ATTRIBUTE *pAttributeValues,
		/* [size_is][in] */ HRESULT *phrErrors)
	{
		if (m_readAttributeCompleted != nullptr)
		{
			return m_readAttributeCompleted(
				dwTransactionID,
				hrStatus,
				hClient,
				dwNumItems,
				pAttributeValues,
				phrErrors);
		}

		return E_NOT_SET;
	}

	virtual HRESULT STDMETHODCALLTYPE OnReadAnnotations(
		/* [in] */ DWORD dwTransactionID,
		/* [in] */ HRESULT hrStatus,
		/* [in] */ DWORD dwNumItems,
		/* [size_is][in] */ OPCHDA_ANNOTATION *pAnnotationValues,
		/* [size_is][in] */ HRESULT *phrErrors)
	{
		if (m_readAnnotations != nullptr)
		{
			return m_readAnnotations(
				dwTransactionID,
				hrStatus,
				dwNumItems,
				pAnnotationValues,
				phrErrors);
		}

		return E_NOT_SET;
	}

	virtual HRESULT STDMETHODCALLTYPE OnInsertAnnotations(
		/* [in] */ DWORD dwTransactionID,
		/* [in] */ HRESULT hrStatus,
		/* [in] */ DWORD dwCount,
		/* [size_is][in] */ OPCHANDLE *phClients,
		/* [size_is][in] */ HRESULT *phrErrors)
	{
		if (m_insertAnnotations != nullptr)
		{
			return m_insertAnnotations(
				dwTransactionID,
				hrStatus,
				dwCount,
				phClients,
				phrErrors);
		}

		return E_NOT_SET;
	}

	virtual HRESULT STDMETHODCALLTYPE OnPlayback(
		/* [in] */ DWORD dwTransactionID,
		/* [in] */ HRESULT hrStatus,
		/* [in] */ DWORD dwNumItems,
		/* [size_is][in] */ OPCHDA_ITEM **ppItemValues,
		/* [size_is][in] */ HRESULT *phrErrors)
	{
		if (m_playCallback != nullptr)
		{
			return m_playCallback(
				dwTransactionID,
				hrStatus,
				dwNumItems,
				ppItemValues,
				phrErrors);
		}

		return E_NOT_SET;
	}

	virtual HRESULT STDMETHODCALLTYPE OnUpdateComplete(
		/* [in] */ DWORD dwTransactionID,
		/* [in] */ HRESULT hrStatus,
		/* [in] */ DWORD dwCount,
		/* [size_is][in] */ OPCHANDLE *phClients,
		/* [size_is][in] */ HRESULT *phrErrors) 
	{
		if (m_updateCompletedCallback)
		{
			return m_updateCompletedCallback(
				dwTransactionID,
				hrStatus,
				dwCount,
				phClients,
				phrErrors);
		}

		return E_NOT_SET;
	}

	virtual HRESULT STDMETHODCALLTYPE OnCancelComplete(
		/* [in] */ DWORD dwCancelID)
	{
		if (m_cancelCompletedCallback)
		{
			return m_cancelCompletedCallback(dwCancelID);
		}

		return E_NOT_SET;
	};

#pragma endregion

#pragma region IUnknown

	virtual HRESULT STDMETHODCALLTYPE QueryInterface(
		/* [in] */ REFIID riid,
		/* [iid_is][out] */ _COM_Outptr_ void __RPC_FAR *__RPC_FAR *ppvObject)
	{
		HRESULT hr;

		if (riid == IID_IUnknown)
		{
			hr = S_OK;
			*ppvObject = this;
		}
		else if (riid == IID_IOPCHDA_DataCallback)
		{
			hr = S_OK;
			*ppvObject = this;
		}
		else
		{
			hr = E_NOINTERFACE;
		}

		if (hr == S_OK)
		{
			this->AddRef();
		}

		return hr;
	};

	virtual ULONG STDMETHODCALLTYPE AddRef(void)
	{
		return ++m_counter;
	};

	virtual ULONG STDMETHODCALLTYPE Release(void) = 0;

#pragma endregion

};