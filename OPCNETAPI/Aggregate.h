#pragma once

namespace OPC
{
	namespace HDA
	{
		public ref class Aggregate
		{
			System::UInt32 m_id;
			System::String^ m_name;
			System::String^ m_description;

		internal:

			Aggregate() : m_id(0), m_name(nullptr), m_description(nullptr)
			{

			}

			static Aggregate^ Create(DWORD id, LPWSTR name, LPWSTR description)
			{
				Aggregate^ a = gcnew Aggregate();

				a->m_id = id;

				a->m_name = System::Runtime::InteropServices::Marshal::PtrToStringUni((System::IntPtr) name);
				a->m_description = System::Runtime::InteropServices::Marshal::PtrToStringUni((System::IntPtr) description);

				return a;
			}

		public:			

			property System::UInt32 Id
			{
				System::UInt32 get()
				{
					return m_id;
				}
			}

			property System::String^ Name
			{
				System::String^ get()
				{
					return m_name;
				}
			}

			property System::String^ Description
			{
				System::String^ get()
				{
					return m_description;
				}
			}
		};
	}
}
