#pragma once

#include "Enumerator.h"
#include "Server.h"

namespace OPC
{
	public value class URL
	{
	internal:
		System::Uri^ m_uri;

	public:

		URL(System::Uri^  url)
		{
			this->m_uri = url;
		}

		property System::String^ Host
		{
			System::String^ get()
			{
				return m_uri->Host;
			}
		}

		Server ToServer()
		{
			Server s;

			System::String^ host = m_uri->Host;

			int position = m_uri->LocalPath->IndexOf("{");

			s.m_ProgId = m_uri->LocalPath->Substring(1, (position == -1) ? m_uri->LocalPath->Length - 1 : position - 1);

			int start = m_uri->LocalPath->IndexOf("{");
			int end = (start > 0) ? m_uri->LocalPath->IndexOf("}", start) + 1 : -1;

			if (start > 0 && end > 0)
			{
				s.m_Guid = System::Guid::Parse(m_uri->LocalPath->Substring(start, end - start));
			}
			else
			{
				OPC::Enum::Enumerator^ enumerator = gcnew OPC::Enum::Enumerator();

				s.m_Guid = enumerator->GetCLSIDFromProgID(s.m_ProgId);
			}

			return s;
		}

		Server ToServer(System::Net::NetworkCredential^ credentials)
		{
			Server s;

			System::String^ host = m_uri->Host;

			int position = m_uri->LocalPath->IndexOf("{");

			s.m_ProgId = m_uri->LocalPath->Substring(1, (position == -1) ? m_uri->LocalPath->Length - 1 : position - 1);

			int start = m_uri->LocalPath->IndexOf("{");
			int end = (start > 0) ? m_uri->LocalPath->IndexOf("}", start) + 1 : -1;

			if (start > 0 && end > 0)
			{
				s.m_Guid = System::Guid::Parse(m_uri->LocalPath->Substring(start, end - start));
			}
			else
			{
				OPC::Enum::Enumerator^ enumerator = gcnew OPC::Enum::Enumerator(host, credentials);

				s.m_Guid = enumerator->GetCLSIDFromProgID(s.m_ProgId);

				enumerator->Release();
			}

			return s;
		}

		System::String^ ToString() override
		{
			return this->m_uri->ToString();
		}
	};
}