#pragma once

#include <Windows.h>

#include "Conversions.h"

namespace OPC
{
	namespace HDA
	{
		public ref class HistorianStatus
		{
			System::DateTime m_currentTime;
			System::DateTime m_startTime;

			System::UInt16 m_majorVersion;
			System::UInt16 m_minorVersion;

			System::UInt32 m_buildNumber;
			System::UInt32 m_maxReturnValues;

			System::String^ m_status;
			System::String^ m_vendorInfo;

		internal:

			HistorianStatus()
			{

			}

			static HistorianStatus^ Create(FILETIME* current, FILETIME* startTime, WORD major, WORD minor, WORD buildNumber, DWORD maxReturnValues, LPWSTR status, LPWSTR vendorInfo)
			{
				HistorianStatus^ h = gcnew HistorianStatus();

				h->m_startTime = FileTimeToDateTime(*startTime);
				h->m_currentTime = FileTimeToDateTime(*current);

				h->m_majorVersion = major;
				h->m_minorVersion = minor;

				h->m_buildNumber = buildNumber;
				h->m_maxReturnValues = maxReturnValues;

				h->m_status = System::Runtime::InteropServices::Marshal::PtrToStringAuto((System::IntPtr) status);
				h->m_vendorInfo = System::Runtime::InteropServices::Marshal::PtrToStringAuto((System::IntPtr) vendorInfo);

				return h;
			}

		public:

			property System::DateTime CurrentTime
			{
				System::DateTime get()
				{
					return m_currentTime;
				}
			}

			property System::DateTime StartTime
			{
				System::DateTime get()
				{
					return m_startTime;
				}
			}

			property System::UInt16 MajorVersion
			{
				System::UInt16 get()
				{
					return m_majorVersion;
				}
			}

			property System::UInt16 MinorVersion
			{
				System::UInt16 get()
				{
					return m_minorVersion;
				}
			}

			property System::UInt32 BuildNumber
			{
				System::UInt32 get()
				{
					return m_buildNumber;
				}
			}

			property System::UInt32 MaxReturnValues
			{
				System::UInt32 get()
				{
					return m_maxReturnValues;
				}
			}

			property System::String^ Status
			{
				System::String^ get()
				{
					return m_status;
				}
			}

			property System::String^ VendorInfo
			{
				System::String^ get()
				{
					return m_vendorInfo;
				}
			}
		};
	}
}