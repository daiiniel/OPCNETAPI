#pragma once

#include "opc_ae.h"
#include "opchda.h"
#include "opcda.h"

#include "ComUtils.h"

namespace OPC
{
	public ref class Specification
	{

	protected:

		System::Guid m_guid;

		Specification(System::Guid g)
		{
			m_guid = g;
		}

	public:

		static Specification^ OPC_DA10 = gcnew Specification(IIDToGuid(CATID_OPCDAServer10));
		static Specification^ OPC_DA20 = gcnew Specification(IIDToGuid(CATID_OPCDAServer20));
		static Specification^ OPC_DA30 = gcnew Specification(IIDToGuid(CATID_OPCDAServer30));

		static Specification^ OPC_AE10 = gcnew Specification(IIDToGuid(CATID_OPCAEServer10));
		static Specification^ OPC_HDA10 = gcnew Specification(IIDToGuid(CATID_OPCHDAServer10));

		property System::Guid GUID
		{
			System::Guid get()
			{
				return m_guid;
			}
		}
	};
}

