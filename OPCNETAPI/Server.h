#pragma once


namespace OPC
{

	public value class Server
	{

	internal:

		System::String^ m_ProgId;
		System::Guid m_Guid;
		System::String^ m_UserType;
		System::String^ m_VerProgId;

	public:

		property System::String^ ProgId
		{
			System::String^ get()
			{
				return m_ProgId;
			}
		}

		property System::Guid GUID
		{
			System::Guid get()
			{
				return m_Guid;
			}
		}

		property System::String^ UserType
		{
			System::String^ get()
			{
				return m_UserType;
			}
		}

		property System::String^ VerProgId
		{
			System::String^ get()
			{
				return m_VerProgId;
			}
		}

	};
}
