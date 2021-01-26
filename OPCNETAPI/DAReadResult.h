#pragma once

#include "Result.h"
#include "ReadResult.h"
#include "Qualities.h"
#include "ItemDefinition.h"

namespace OPC
{
	namespace DA
	{
		public value class DAReadResult
		{

		protected:

			System::String^ m_Name;
			System::String^ m_ItemID;

			System::UInt32 m_ClientHandle;
			System::UInt32 m_ServerHandle;

			Results m_Result;

			ReadResult m_Value;

		internal:

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

			void SetValue(ReadResult r)
			{
				m_Value = r;
			}

			DAReadResult(ItemDefinition^ i)
			{
				SetName(i->Name);
				SetItemID(i->ItemID);
				SetClientHandle(i->ClientHandle);
				SetServerHandle(i->ServerHandle);
				SetResult(Results::OPC_E_NOINFO);
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
			property ReadResult Value
			{
				ReadResult get()
				{
					return m_Value;
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

			static operator DAReadResult(ItemDefinition^ i)
			{
				DAReadResult r;

				r.SetName(i->Name);
				r.SetItemID(i->ItemID);
				r.SetClientHandle(i->ClientHandle);
				r.SetServerHandle(i->ServerHandle);

				r.SetResult(OPC::Results::False);

				return r;
			}
		};
	}
}