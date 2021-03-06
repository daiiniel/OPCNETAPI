#pragma once

namespace OPC
{
	public enum class Qualities 
	{
		OPC_STATUS_MASK = 0xfc,

		OPC_LIMIT_MASK = 0x3,

		OPC_QUALITY_BAD = 0,

		OPC_QUALITY_UNCERTAIN = 0x40,

		OPC_QUALITY_GOOD = 0xc0,

		OPC_QUALITY_CONFIG_ERROR = 0x4,

		OPC_QUALITY_NOT_CONNECTED = 0x8,

		OPC_QUALITY_DEVICE_FAILURE = 0xc,

		OPC_QUALITY_SENSOR_FAILURE = 0x10,

		OPC_QUALITY_LAST_KNOWN = 0x14,

		OPC_QUALITY_COMM_FAILURE = 0x18,

		OPC_QUALITY_OUT_OF_SERVICE = 0x1c,

		OPC_QUALITY_WAITING_FOR_INITIAL_DATA = 0x20,

		OPC_QUALITY_LAST_USABLE = 0x44,

		OPC_QUALITY_SENSOR_CAL = 0x50,

		OPC_QUALITY_EGU_EXCEEDED = 0x54,

		OPC_QUALITY_SUB_NORMAL = 0x58,

		OPC_QUALITY_LOCAL_OVERRIDE = 0xd8,

		OPC_LIMIT_OK = 0,

		OPC_LIMIT_LOW = 0x1,

		OPC_LIMIT_HIGH = 0x2,

		OPC_LIMIT_CONST = 0x3,
	};
}