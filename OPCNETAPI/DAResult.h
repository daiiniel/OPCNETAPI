#pragma once

#include "Result.h"
#include "ItemDefinition.h"
#include "Qualities.h"

namespace OPC
{
	namespace DA
	{
		public ref class DAResult
		{

		protected:

			System::String^ m_Name;
			System::String^ m_ItemID;

			System::UInt32 m_ClientHandle;
			System::UInt32 m_ServerHandle;

			Results m_Result;

		internal:

			DAResult(ItemDefinition^ i)
			{
				this->SetName(i->Name);
				this->SetItemID(i->ItemID);
				this->SetClientHandle(i->ClientHandle);
				this->SetServerHandle(i->ServerHandle);
				this->SetResult(Results::OPC_E_NOINFO);
			}

			void SetName(System::String^ name)
			{
				m_Name = name;
			}
			void SetItemID(System::String^ itemID)
			{
				m_ItemID = itemID;
			}
			void SetClientHandle(System::UInt32 cHandle)
			{
				m_ClientHandle = cHandle;
			}
			void SetServerHandle(System::UInt32 sHandle)
			{
				m_ServerHandle = sHandle;
			}

			void SetResult(Results r)
			{
				m_Result = r;
			}

		public:

			property System::String^ Name
			{
				System::String^ get()
				{
					return m_Name;
				}
			}
			property System::String^ ItemID
			{
				System::String^ get()
				{
					return m_ItemID;
				}
			}
			property System::UInt32 ClientHandle
			{
				System::UInt32 get()
				{
					return m_ClientHandle;
				}
			}
			property System::UInt32 ServerHandle
			{
				System::UInt32 get()
				{
					return m_ServerHandle;
				}
			}
			property Results Result
			{
				Results get()
				{
					return m_Result;
				}
			}

			ItemDefinition^ ToItemDefinition()
			{
				ItemDefinition^ definition = gcnew ItemDefinition(ItemID);

				definition->SetName(Name);
				definition->SetServerHandle(ServerHandle);
				definition->SetClientHandle(ClientHandle);

				return definition;
			}

			static operator DAResult^(ItemDefinition^ i)
			{
				DAResult^ r = gcnew DAResult(i);

				return r;
			}
		};
	}
}