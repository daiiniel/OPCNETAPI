#pragma once

#include "Result.h"


namespace OPC
{
	public ref class ItemDefinition
	{

	protected:

		System::String^ m_Name;
		System::String^ m_ItemID;

		System::UInt32 m_ClientHandle;
		System::UInt32 m_ServerHandle;

		Results m_Result;

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
		void SetResult(Results v)
		{
			m_Result = v;
		}

	public:

		ItemDefinition(System::String^ itemID)
		{
			SetName(System::String::Empty);
			SetItemID(itemID);

			SetClientHandle(0);
			SetServerHandle(0);
		}

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
	};
}

