#pragma once

#include "ItemDefinition.h"

using namespace OPC;

namespace OPC
{
	namespace HDA
	{
		public ref class AggregateItem
		{
			ItemDefinition^ m_item;
			System::UInt32 m_aggregate;

		public:

			AggregateItem(ItemDefinition^ i, System::UInt32 a)
			{
				m_item = i;
				m_aggregate = a;
			}

			property ItemDefinition^ Item
			{
				ItemDefinition^ get()
				{
					return m_item;
				}
			}

			property System::UInt32 Aggregate
			{
				System::UInt32 get()
				{
					return m_aggregate;
				}
			}

		};
	}
}
