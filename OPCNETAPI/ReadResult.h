#pragma once

#include "Qualities.h"

namespace OPC
{

	public value class ReadResult
	{

	protected:

		System::Object^ m_Value;
		Qualities m_Quality;
		System::DateTime m_TimeStamp;

	internal:

		void SetValue(System::Object^ value)
		{
			m_Value = value;
		}

		void SetQuality(Qualities quality)
		{
			m_Quality = quality;
		}

		void SetTimeStamp(System::DateTime time)
		{
			m_TimeStamp = time;
		}
		
	public:

		property System::Object^ Value
		{
			System::Object^ get()
			{
				return m_Value;
			}
		}

		property Qualities Quality
		{
			Qualities get()
			{
				return m_Quality;
			}
		}

		property System::DateTime TimeStamp
		{
			System::DateTime get()
			{
				return m_TimeStamp;
			}
		}

	};
}
