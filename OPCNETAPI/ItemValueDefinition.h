#pragma once

#include "ItemDefinition.h"

using namespace OPC;

public ref class ItemValueDefinition : ItemDefinition
{
protected:

	System::Object^ m_value;

public:

	ItemValueDefinition(ItemDefinition^ item, Object^ value) : ItemDefinition(item->Name)
	{
		this->SetClientHandle(item->ClientHandle);
		this->SetServerHandle(item->ServerHandle);

		this->Value = value;
	}
	
	property System::Object^ Value
	{
		System::Object^ get()
		{
			return m_value;
		}
		void set(System::Object^ value)
		{
			this->m_value = value;
		}
	}

};