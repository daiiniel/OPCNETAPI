#pragma once

#include "Result.h"
#include "ItemDefinition.h"
#include "ReadResult.h"

namespace OPC
{
	namespace HDA
	{
		public ref class HDAReadResult
		{

		protected:

			System::String^ m_Name;
			System::String^ m_ItemID;

			System::UInt32 m_ClientHandle;
			System::UInt32 m_ServerHandle;

			Results m_Result;

			System::Collections::Generic::List<ReadResult>^ m_Values;

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

			HDAReadResult(ItemDefinition^ i)
			{
				SetName(i->Name);
				SetItemID(i->ItemID);
				SetClientHandle(i->ClientHandle);
				SetServerHandle(i->ServerHandle);

				m_Values = gcnew System::Collections::Generic::List<ReadResult>();
			}

			HDAReadResult(ItemDefinition^ i, int capacity) 
			{
				SetName(i->Name);
				SetItemID(i->ItemID);
				SetClientHandle(i->ClientHandle);
				SetServerHandle(i->ServerHandle);

				m_Values = gcnew System::Collections::Generic::List<ReadResult>(capacity);
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

			property System::Collections::Generic::List<ReadResult>^ Values
			{
				System::Collections::Generic::List<ReadResult>^ get()
				{
					return m_Values;
				}
			}
		};
	}
}
