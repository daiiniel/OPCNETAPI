#pragma once

#include "ComBase.h"
#include "ComUtils.h"
#include "OpcEnum.h"
#include "Specification.h"
#include "Server.h"

#include "opcda.h"

using namespace OPC;

namespace OPC
{
	namespace Enum
	{

		public ref class Enumerator : ComBase
		{

		protected:

			IOPCServerList2* GetServer();
						
		public:

			Enumerator();
			Enumerator(
				System::String^ host,
				System::Net::NetworkCredential^ credentials);

			~Enumerator();

			System::Collections::Generic::IList<Server>^ GetAvailableServers(... array<Specification^>^ specifications);
			System::Guid GetCLSIDFromProgID(System::String^ progID);

			void Release();

			static Server FromProgId(System::String^ progId);

		};
	}
}
