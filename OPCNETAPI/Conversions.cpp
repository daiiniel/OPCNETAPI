#include "Stdafx.h"
#include "Conversions.h"

tagOPCHDA_TIME DatetimeToOPCTime(System::DateTime d)
{
	tagOPCHDA_TIME t;

	t.bString = FALSE;
	t.szTime = NULL;

	SYSTEMTIME st;

	st.wYear = d.Year;
	st.wMonth = d.Month;
	st.wDay = d.Day;
	st.wHour = d.Hour;
	st.wMinute = d.Minute;
	st.wSecond = d.Second;
	st.wMilliseconds = d.Millisecond;

	SystemTimeToFileTime(&st, &t.ftTime);

	return t;
}

System::DateTime OPCTimeToDatetime(tagOPCHDA_TIME& t)
{
	if (t.bString)
	{
		throw gcnew System::NotImplementedException();
	}
	else
	{
		SYSTEMTIME st;

		FileTimeToSystemTime(&t.ftTime, &st);

		return System::DateTime(
			st.wYear,
			st.wMonth,
			st.wDay,
			st.wHour,
			st.wMinute,
			st.wSecond,
			st.wMilliseconds,
			System::DateTimeKind::Utc);
	}
}

FILETIME DateTimeToFileTime(System::DateTime% d)
{
	FILETIME ft;

	SYSTEMTIME st;

	st.wYear = d.Year;
	st.wMonth = d.Month;
	st.wDay = d.Day;
	st.wHour = d.Hour;
	st.wMinute = d.Minute;
	st.wSecond = d.Second;
	st.wMilliseconds = d.Millisecond;

	SystemTimeToFileTime(&st, &ft);

	return ft;
}

System::DateTime FileTimeToDateTime(FILETIME& ft)
{
	SYSTEMTIME st;

	FileTimeToSystemTime(&ft, &st);

	return System::DateTime(
		st.wYear,
		st.wMonth,
		st.wDay,
		st.wHour,
		st.wMinute,
		st.wSecond,
		st.wMilliseconds,
		System::DateTimeKind::Utc);
}

FILETIME TimeSpanToFileTime(System::TimeSpan time)
{
	FILETIME ft;

	UINT64 nanoSeconds = (UINT64)time.TotalMilliseconds;

	*(UINT64*)(&ft) = nanoSeconds * 1e4;

	return ft;
}