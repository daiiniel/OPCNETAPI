#pragma once

#include "opccomn.h"
#include "opcda.h"

#include "URL.h"
#include "Subscription.h"
#include "BrowseType.h"
#include "ItemDefinition.h"
#include "ComBase.h"
#include "ComUtils.h"
#include "Server.h"

class OPCShutdown;

using namespace OPC;
using namespace OPC::DA;

namespace OPC
{
	namespace DA
	{

		public ref class ServerShutdownEventArgs : System::EventArgs
		{
		public:

			System::String^ Reason;

			ServerShutdownEventArgs(System::String^ reason) { Reason = reason; }
		};

		public ref class DAServer : public ComBase
		{

		private:

			IOPCBrowseServerAddressSpace* m_IOPCAddressSpaceBrowser;

			System::Runtime::InteropServices::GCHandle handle;

			void Init();

		protected:

			OPCShutdown* m_shutdown;

			System::UInt32 cCounter;

			IOPCServer* GetServer();
			IOPCBrowseServerAddressSpace* GetBrowser();

			System::Collections::Generic::Dictionary<System::UInt32, Subscription^>^ subscriptions;


			System::EventHandler<ServerShutdownEventArgs^>^ m_serverShutDownHandler;

			void OnServerShutDown(ServerShutdownEventArgs^ args);

		internal:

			HRESULT OnServerShutDownFunc(LPCWSTR szReason);

		public:

			event System::EventHandler<ServerShutdownEventArgs^>^ ServerShutDown
			{
				void add(System::EventHandler<ServerShutdownEventArgs^>^ v) { m_serverShutDownHandler += v; }
				void remove(System::EventHandler<ServerShutdownEventArgs^>^ v) { m_serverShutDownHandler -= v; }
			}

			DAServer();
			~DAServer();			
			
			void Connect(Server s);

			void Connect(
				Server s,
				System::String^ host,
				System::Net::NetworkCredential^ credentials);
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

			System::Collections::Generic::IList<ItemDefinition^>^ Browse(System::String^ nodeToBrowse, BrowseType browseType);
			System::Collections::Generic::IList<ItemDefinition^>^ BrowseAsync(System::String^ nodeToBrowse, BrowseType browseType, [System::Runtime::InteropServices::Optional] System::Nullable<int> timeout);
			
			Subscription^ AddSubscription(System::String^ sName, System::UInt32 updateRate, float deadBand, bool activated);
			void RemoveSubscription(Subscription^ subscriptionToRemove);

			Subscription^ AddSubscriptionAsync(System::String^ sName, System::UInt32 updateRate, float deadBand, bool activated, [System::Runtime::InteropServices::Optional] System::Nullable<int> timeout);
			void RemoveSubscriptionAsync(Subscription^ subscriptionToRemove, [System::Runtime::InteropServices::Optional] System::Nullable<int> timeout);

			void Release();

		};
		
		inline DAServer^ ptrToDAServer(void* ptr)
		{
			return (DAServer^) System::Runtime::InteropServices::GCHandle::FromIntPtr(System::IntPtr(ptr)).Target;
		}
	}
}
