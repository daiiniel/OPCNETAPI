#pragma once

#include <opchda.h>

tagOPCHDA_TIME DatetimeToOPCTime(System::DateTime d);
System::DateTime OPCTimeToDatetime(tagOPCHDA_TIME& t);

FILETIME DateTimeToFileTime(System::DateTime% d);
System::DateTime FileTimeToDateTime(FILETIME& ft);

FILETIME TimeSpanToFileTime(System::TimeSpan time);
