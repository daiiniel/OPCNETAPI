#pragma once

#include "Result.h"
#include "URL.h"
#include "BrowseType.h"
#include "Server.h"
#include "ComBase.h"
#include "ComUtils.h"
#include "ItemDefinition.h"
#include "HDAReadResult.h"
#include "Conversions.h"
#include "Aggregate.h"
#include "HistorianStatus.h"
#include "HDAResult.h"
#include "AggregateItem.h"

#include <opchda.h>
#include <opcerror.h>

using namespace OPC;
using namespace OPC::HDA;

namespace OPC
{
	namespace HDA
	{		
		public ref class HDAServer : public ComBase
		{

			IOPCHDA_SyncRead* m_IOPCSyncRead;
			IOPCHDA_Browser* m_IOPCBrowser;
			IOPCHDA_AsyncRead* m_IOPCAsyncRead;

		private:

			HRESULT processResults(
				System::Collections::Generic::IList<ItemDefinition^>^ definitions,
				System::Collections::Generic::IList<HDAReadResult^>^ res,
				OPCHDA_ITEM* values,
				HRESULT* errors,
				size_t count);

		protected:

			System::UInt32 cCounter;

			IOPCHDA_Server* GetServer();
			IOPCHDA_SyncRead* GetSyncRead();
			IOPCHDA_Browser* GetBrowser();
			IOPCHDA_AsyncRead* GetAsyncRead();

		public:

			HDAServer();
			
			void Connect(Server s);
			void Connect(
				Server s,
				System::String^ host,
				System::Net::NetworkCredential^credentials);
			void Connect(
				URL url,
				System::Net::NetworkCredential^ credentials);

			void ConnectAsync(
				Server s,
				[System::Runtime::InteropServices::Optional] System::Nullable<int> timeout);
			void ConnectAsync(
				Server s,
				System::String^ host,
				System::Net::NetworkCredential^credentials,
				[System::Runtime::InteropServices::Optional] System::Nullable<int> timeout);
			void ConnectAsync(
				URL url,
				System::Net::NetworkCredential^ credentials,
				[System::Runtime::InteropServices::Optional] System::Nullable<int> timeout);

			~HDAServer();

			HistorianStatus^ GetHistorianStatus();

			void ValidateItemIDs(System::Collections::Generic::List<System::Tuple<System::String^, bool>^> ^ itemsToValidate);

			System::Collections::Generic::IList<ItemDefinition^>^ Browse(
				System::String^ nodeToBrowse, 
				BrowseType browseType);

			void GetHandlesAsync(System::Collections::Generic::IList<ItemDefinition^>^ definitions);
			void GetHandles(System::Collections::Generic::IList<ItemDefinition^>^ definitions);

			void RemoveHandlesAsync(System::Collections::Generic::IList<ItemDefinition^>^ dedinitions);
			void RemoveHandles(System::Collections::Generic::IList<ItemDefinition^>^ definitions);

			System::Collections::Generic::IList<Aggregate^>^ GetAggregates();

#pragma region Sync

			System::Collections::Generic::IList<HDAReadResult^>^ ReadRaw(
				System::DateTime% from, 
				System::DateTime% to, 
				System::Collections::Generic::IList<ItemDefinition^>^ definitions, 
				System::UInt32 maxItems, 
				bool includeBounds);

			System::Collections::Generic::IList<HDAReadResult^>^ ReadProcessed(
				System::DateTime% from,
				System::DateTime% to,
				System::TimeSpan interval,
				System::UInt32 aggregate,
				System::Collections::Generic::IList<ItemDefinition^>^ definitions);

			System::Collections::Generic::IList<HDAReadResult^>^ ReadProcessed(
				System::DateTime% from,
				System::DateTime% to,
				System::TimeSpan interval,
				System::Collections::Generic::IList<AggregateItem^>^ definitions);

			System::Collections::Generic::IList<HDAReadResult^>^ ReadAtTime(
				System::Collections::Generic::IList<System::DateTime>^ timeStamps,
				System::Collections::Generic::IList<ItemDefinition^>^definitions);

#pragma endregion

#pragma region Async

			System::Collections::Generic::IList<HDAResult^>^ AdviseRaw(
				System::UInt32 transactionId,
				System::DateTime startTime,
				System::TimeSpan interval,
				System::Collections::Generic::IList<ItemDefinition^>^ definitions,
				[System::Runtime::InteropServices::OutAttribute] System::UInt32% cancelId);

#pragma endregion

			void Release();
		};
	}
}
