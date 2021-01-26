#pragma once

#include "DADataCallback.h"
#include "ItemDefinition.h"
#include "ItemValueDefinition.h"
#include "ComBase.h"
#include "ComUtils.h"
#include "DAReadResult.h"
#include "DAResult.h"
#include "DataSources.h"
#include "Conversions.h"

using namespace OPC;
using namespace OPC::DA;

class DADataCallback;

namespace OPC
{
	namespace DA
	{
		public value struct DataChangedEventArgs
		{
		public:
			System::UInt32 TransactionId;
			System::Collections::Generic::List<DAReadResult>^ Results;
		};

		public delegate void DataChangedEventHandler(System::Object^ sender, DataChangedEventArgs e);

		public value struct ReadCompletedEventArgs
		{
		public:
			System::UInt32 TransactionId;
			System::Collections::Generic::List<DAReadResult>^ Results;
		};

		public delegate void ReadCompletedEventHandler(System::Object^ sender, ReadCompletedEventArgs e);

		public value struct WriteCompletedEventArgs
		{
		public:
			System::UInt32 TransactionId;
		};

		public delegate void WriteCompletedEventHandler(System::Object^ sender, WriteCompletedEventArgs e);

		public value struct CancelCompletedEventArgs
		{
		public:
			System::UInt32 TransactionId;
		};

		public delegate void CancelCompletedEventHandler(System::Object^ sender, CancelCompletedEventArgs e);

		public ref class Subscription : public ComBase
		{
		private:

			IOPCSyncIO2* m_IOPCSyncIO;
			IOPCAsyncIO2* m_IOPCASyncIO;
			IOPCItemMgt* m_IOPCItemMgt;
			IOPCGroupStateMgt* m_IOPCGroupStateMgt;

		internal:

			HRESULT OnDataChangedFunc(
				DWORD dwTransid,
				OPCHANDLE hGroup,
				HRESULT hrMasterquality,
				HRESULT hrMastererror,
				DWORD dwCount,
				OPCHANDLE *phClientItems,
				VARIANT *pvValues,
				WORD *pwQualities,
				FILETIME *pftTimeStamps,
				HRESULT *pErrors);

			HRESULT OnReadCompletedFunc(
				DWORD dwTransid,
				OPCHANDLE hGroup,
				HRESULT hrMasterquality,
				HRESULT hrMastererror,
				DWORD dwCount,
				OPCHANDLE *phClientItems,
				VARIANT *pvValues,
				WORD *pwQualities,
				FILETIME *pftTimeStamps,
				HRESULT *pErrors);

			HRESULT OnWriteCompletedFunc(
				DWORD dwTransid,
				OPCHANDLE hGroup,
				HRESULT hrMastererr,
				DWORD dwCount,
				OPCHANDLE *pClienthandles,
				HRESULT *pErrors);

			HRESULT OnCancelCompleted(
				DWORD dwTransid,
				OPCHANDLE hGroup);

			DADataCallback* dataCallback;

		protected:

			System::UInt32 cCounter;

			bool active;

			System::String^ m_name;

			System::UInt32 m_clientHandle;
			System::UInt32 m_serverHandle;

			System::UInt32 m_UpdateRate;
			float m_DeadBand;

			System::Collections::Generic::Dictionary<System::UInt32, ItemDefinition^>^ items;

			System::Runtime::InteropServices::GCHandle handle;

			DataChangedEventHandler^ DataChangedHandler;
			ReadCompletedEventHandler^ ReadCompletedHandler;
			WriteCompletedEventHandler^ WriteCompletedHandler;
			CancelCompletedEventHandler^ CancelCompletedHandler;

			IOPCItemMgt* GetServer();
			IOPCSyncIO2* GetSyncIO();
			IOPCAsyncIO2* GetAsyncIO();
			IOPCGroupStateMgt* GetGroupStateMgt();

			void ChangeState(DWORD* updateRate, bool* status, float* deadband);

		internal:

			void SetName(System::String^ v) { m_name = v; }

			void SetClientHandle(System::UInt32 v) { m_clientHandle = v; }
			void SetServerHandle(System::UInt32 v) { m_serverHandle = v; }

			void SetUpdateRate(System::UInt32 v) { m_UpdateRate = v; }
			void SetDeadBand(float v) { m_DeadBand = v; }

		public:

			event DataChangedEventHandler^ DataChanged
			{
				void add(DataChangedEventHandler^ p) { DataChangedHandler += p; }
				void remove(DataChangedEventHandler^ p) { DataChangedHandler -= p; }
			}
			event ReadCompletedEventHandler^ ReadCompleted
			{
				void add(ReadCompletedEventHandler^ p) { ReadCompletedHandler += p; }
				void remove(ReadCompletedEventHandler^ p) { ReadCompletedHandler -= p; }
			}
			event WriteCompletedEventHandler^ WriteCompleted
			{
				void add(WriteCompletedEventHandler^ p) { WriteCompletedHandler += p; }
				void remove(WriteCompletedEventHandler^ p) { WriteCompletedHandler -= p; }
			}
			event CancelCompletedEventHandler^ CancelCompleted
			{
				void add(CancelCompletedEventHandler^ p) { CancelCompletedHandler += p; }
				void remove(CancelCompletedEventHandler^ p) { CancelCompletedHandler -= p; }
			}

			Subscription(
				System::String^ name,
				System::UInt32 cHandle,
				System::UInt32 sHandle,
				System::UInt32 uRate,
				float dBand,
				ComBase^ pParent,
				IUnknown* ppIn);

			~Subscription();

			property System::String^ Name
			{
				System::String^ get()
				{
					return m_name;
				}
			}
			property System::UInt32 ClientHandle
			{
				System::UInt32 get()
				{
					return m_clientHandle;
				}
			}
			property System::UInt32 ServerHandle
			{
				System::UInt32 get()
				{
					return m_serverHandle;
				}
			}
			property System::UInt32 UpdateRate
			{
				System::UInt32 get()
				{
					return m_UpdateRate;
				}
			}
			property float DeadBand
			{
				float get()
				{
					return m_DeadBand;
				}
			}

			void AddItemsAsync(
				System::Collections::Generic::IList<ItemDefinition^>^ definitionsToAdd,
				[System::Runtime::InteropServices::Optional] System::Nullable<int> timeout);
			void AddItems(
				System::Collections::Generic::IList<ItemDefinition^>^ definitionsToAdd);

			void RemoveItemsAsync(
				System::Collections::Generic::IList<ItemDefinition^>^ definitionsToRemove,
				[System::Runtime::InteropServices::Optional] System::Nullable<int> timeout);
			void RemoveItems(
				System::Collections::Generic::IList<ItemDefinition^>^ definitionsToRemove);

			System::Collections::Generic::IList<DAReadResult>^ Read(System::Collections::Generic::IList<ItemDefinition^>^ definitionsToRead, DataSources source);
			System::Collections::Generic::IList<DAResult^>^ ReadAsync(
				System::Collections::Generic::IList<ItemDefinition^>^
				definitionsToRead,
				DataSources source,
				System::UInt32 transactionId,
				[System::Runtime::InteropServices::Out] System::UInt32% cancelId);

			System::Collections::Generic::IList<DAResult^>^ Write(System::Collections::Generic::IList<ItemValueDefinition^>^ pairItemValuesToWrite);
			System::Collections::Generic::IList<DAResult^>^ WriteAsync(
				System::Collections::Generic::IList<ItemValueDefinition^>^ pairItemValuesToWrite,
				System::UInt32 transactionId,
				[System::Runtime::InteropServices::Out] System::UInt32% cancelId);

			void Refresh(
				DataSources source,
				System::UInt32 transactionId,
				[System::Runtime::InteropServices::OutAttribute] System::UInt32% cancelId);

			void Cancel(System::UInt32);

			void Activate();
			void DeActivate();

			void ChangeState(System::UInt32 updateRate, bool status, float deadband);

			void Release();

		};

		inline OPC::DA::Subscription^ ptrToSubscription(void* ptr)
		{
			System::Runtime::InteropServices::GCHandle h = System::Runtime::InteropServices::GCHandle::FromIntPtr(System::IntPtr(ptr));

			OPC::DA::Subscription^ s = (OPC::DA::Subscription^) h.Target;

			return s;
		}
	}
}
