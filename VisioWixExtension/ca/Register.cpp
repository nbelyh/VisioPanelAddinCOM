
#include "stdafx.h"

/***************************************************************/

DWORD GetSAM(DWORD sam, REG_KEY_BITNESS iBitness)
{
	if (iBitness == REG_KEY_32BIT)
		sam |= KEY_WOW64_32KEY;

	if (iBitness == REG_KEY_64BIT)
		sam |= KEY_WOW64_64KEY;

	return sam;
}

HRESULT CreateOfficeRegistryKeyAtRoot(
	LPCWSTR pwzId, 
	LPCWSTR pwzFile, 
	LPCWSTR pwzFriendlyName, 
	LPCWSTR pwzDescription,
	int iCommandLineSafe,
	int iLoadBehavior,
	HKEY hKeyRoot, 
	REG_KEY_BITNESS iBitness)
{
	HRESULT hr = S_OK;
	HKEY hKey = NULL;

	LPWSTR wszRegPath = NULL;
	LPWSTR wszFileEntry = NULL;

    WcaLog(LOGMSG_VERBOSE, "CreateOfficeRegistryKey: Bitness=%ld, Id=%ls, Manifest=%ls, FriendlyName=%ls, Description=%ls", 
		iBitness, pwzId, pwzFile, pwzFriendlyName, pwzDescription);

    hr = StrAllocFormatted(&wszRegPath, L"Software\\Microsoft\\Visio\\Addins\\%ls", pwzId);
    ExitOnFailure(hr, "Failed to allocate string for registry path.");

	hr = RegCreateEx(hKeyRoot, wszRegPath, GetSAM(KEY_READ|KEY_WRITE, iBitness), FALSE, NULL, &hKey, NULL);
	ExitOnFailure(hr, "Failed to create or open the registry key: %ls. It looks like Visio is not installed or security issue", wszRegPath);

	WcaLog(LOGMSG_VERBOSE, "Created or opened registry key: %ls", wszRegPath);

    hr = StrAllocFormatted(&wszFileEntry, L"file:///%ls|vstolocal", pwzFile);
    ExitOnFailure(hr, "Failed to allocate string for registry file entry.");

	hr = RegWriteString(hKey, L"Manifest", wszFileEntry);
	ExitOnFailure(hr, "Failed set Manifest value.");

	hr = RegWriteString(hKey, L"FriendlyName", pwzFriendlyName);
	ExitOnFailure(hr, "Failed set FriendlyName value.");

	hr = RegWriteString(hKey, L"Description", pwzDescription);
	ExitOnFailure(hr, "Failed set Description value.");

	hr = RegWriteNumber(hKey, L"CommandLineSafe", iCommandLineSafe);
	ExitOnFailure(hr, "Failed set CommandLineSafe value value.");

	hr = RegWriteNumber(hKey, L"LoadBehavior", iLoadBehavior);
	ExitOnFailure(hr, "Failed set LoadBehavior value value.");

LExit:
	ReleaseStr(wszFileEntry);
	ReleaseStr(wszRegPath);
	ReleaseRegKey(hKey);
	return hr;
}

HRESULT DeleteOfficeRegistryKeyAtRoot(
	LPCWSTR pwzId, 
	HKEY hKeyRoot, 
	REG_KEY_BITNESS iBitness)
{
	HRESULT hr = S_OK;

    WcaLog(LOGMSG_VERBOSE, "DeleteOfficeRegistryKey: Bitness:%ld, Id=%ls", 
		iBitness, pwzId);

	LPWSTR wszRegPath = NULL;
    hr = StrAllocFormatted(&wszRegPath, L"Software\\Microsoft\\Visio\\Addins\\%ls", pwzId);
    ExitOnFailure(hr, "Failed to allocate registry path.");

	hr = RegDelete(hKeyRoot, wszRegPath, iBitness, FALSE);

	if (E_FILENOTFOUND != hr)
		ExitOnFailure(hr, "Failed to delete the registry key: %ls.", wszRegPath);

	WcaLog(LOGMSG_VERBOSE, "Deleted registry key: %ls", wszRegPath);

LExit:
	return hr;
}

HRESULT RegisterCOMAtRoot(
    LPCWSTR strClsId,
    LPCWSTR strProgId,
    LPCWSTR typeFullName, 
    LPCWSTR strAsmName, 
    LPCWSTR strAsmVersion, 
    LPCWSTR strAsmCodeBase, 
    LPCWSTR strRuntimeVersion,
	HKEY hKeyRoot,
	REG_KEY_BITNESS iBitness)
{
	HRESULT hr = S_OK;

	HKEY hKeyProgId = NULL;
	HKEY hKeyProgIdClassId = NULL;
	HKEY hKeyClassIdProgId = NULL;
	HKEY hKeyClsId = NULL;
	HKEY hKeyClsIdInprocServer = NULL;
	HKEY hKeyClsIdInprocServerVersion = NULL;

	LPWSTR wszRegPath = NULL;
	LPWSTR wszRegPathProgId = NULL;

	hr = StrAllocFormatted(&wszRegPath, L"Software\\Classes\\CLSID\\%ls", strClsId);
	ExitOnFailure(hr, "Failed to allocate string for registry path.");

	hr = RegCreateEx(hKeyRoot, wszRegPath, GetSAM(KEY_READ|KEY_WRITE, iBitness), FALSE, NULL, &hKeyClsId, NULL);
	ExitOnFailure(hr, "Failed to create or open the registry key: %ls", wszRegPath);

	hr = RegWriteString(hKeyClsId, L"", typeFullName);
	ExitOnFailure1(hr, "Failed to write to %ls", wszRegPath);

	hr = RegCreateEx(hKeyClsId, L"InprocServer32", KEY_READ|KEY_WRITE, FALSE, NULL, &hKeyClsIdInprocServer, NULL);
	ExitOnFailure(hr, "Failed to create key %ls\\InprocServer32", wszRegPath);

	hr = RegWriteString(hKeyClsIdInprocServer, L"", L"mscoree.dll");
	ExitOnFailure(hr, "Failed to write to %ls", wszRegPath);

	hr = RegWriteString(hKeyClsIdInprocServer, L"ThreadingModel", L"both");
	ExitOnFailure(hr, "Failed to write to %ls", wszRegPath);

	hr = RegWriteString(hKeyClsIdInprocServer, L"Class", typeFullName);
	ExitOnFailure(hr, "Failed to write to %ls", wszRegPath);

	hr = RegWriteString(hKeyClsIdInprocServer, L"Assembly", strAsmName);
	ExitOnFailure(hr, "Failed to write to %ls", wszRegPath);

	hr = RegWriteString(hKeyClsIdInprocServer, L"RuntimeVersion", strRuntimeVersion);
	ExitOnFailure(hr, "Failed to write to %ls", wszRegPath);

	if (strAsmCodeBase && *strAsmCodeBase)
	{
		hr = RegWriteString(hKeyClsIdInprocServer, L"CodeBase", strAsmCodeBase);
		ExitOnFailure(hr, "Failed to write to %ls", wszRegPath);
	}

	hr = RegCreateEx(hKeyClsIdInprocServer, strAsmVersion, KEY_READ|KEY_WRITE, FALSE, NULL, &hKeyClsIdInprocServerVersion, NULL);
	ExitOnFailure(hr, "Failed to create key %ls\\%ls", wszRegPath, strAsmVersion);

	hr = RegWriteString(hKeyClsIdInprocServerVersion, L"Class", typeFullName);
	ExitOnFailure(hr, "Failed to write to %ls", wszRegPath);

	hr = RegWriteString(hKeyClsIdInprocServerVersion, L"Assembly", strAsmName);
	ExitOnFailure(hr, "Failed to write to %ls", wszRegPath);

	hr = RegWriteString(hKeyClsIdInprocServerVersion, L"RuntimeVersion", strRuntimeVersion);
	ExitOnFailure(hr, "Failed to write to %ls", wszRegPath);

	if (strAsmCodeBase && *strAsmCodeBase)
	{
		hr = RegWriteString(hKeyClsIdInprocServerVersion, L"CodeBase", strAsmCodeBase);
		ExitOnFailure(hr, "Failed to write to %ls", wszRegPath);
	}

	if (strProgId && *strProgId)
	{
		hr = StrAllocFormatted(&wszRegPathProgId, L"Software\\Classes\\%ls", strProgId);
		ExitOnFailure(hr, "Failed to allocate string for registry path.");

		hr = RegCreateEx(hKeyRoot, wszRegPathProgId, KEY_READ|KEY_WRITE, FALSE, NULL, &hKeyProgId, NULL);
		ExitOnFailure(hr, "Failed to create or open the registry key: %ls", wszRegPathProgId);

		hr = RegWriteString(hKeyProgId, L"", typeFullName);
		ExitOnFailure1(hr, "Failed to write to %ls", wszRegPathProgId);

		hr = RegCreateEx(hKeyProgId, L"CLSID", KEY_READ|KEY_WRITE, FALSE, NULL, &hKeyProgIdClassId, NULL);
		ExitOnFailure1(hr, "Failed to create key %ls\\CLSID", wszRegPathProgId);

		hr = RegWriteString(hKeyProgIdClassId, L"", strClsId);
		ExitOnFailure1(hr, "Failed to write to %ls\\CLSID", wszRegPathProgId);

		hr = RegCreateEx(hKeyClsId, L"ProgId", KEY_READ|KEY_WRITE, FALSE, NULL, &hKeyClassIdProgId, NULL);
		ExitOnFailure(hr, "Failed to create key %ls\\ProgId", wszRegPath);

		hr = RegWriteString(hKeyClassIdProgId, L"", strProgId);
		ExitOnFailure(hr, "Failed to write to key %ls\\ProgId", wszRegPath);
	}

	hr = RegCreateEx(hKeyClsId, L"Implemented Categories\\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}", KEY_READ|KEY_WRITE, FALSE, NULL, &hKeyClsIdInprocServer, NULL);
	ExitOnFailure(hr, "Failed to create key", wszRegPath);

    // EnsureManagedCategoryExists();

LExit:
	ReleaseStr(wszRegPath);
	ReleaseStr(wszRegPathProgId);

	ReleaseRegKey(hKeyProgId);
	ReleaseRegKey(hKeyProgIdClassId);
	ReleaseRegKey(hKeyClassIdProgId);

	ReleaseRegKey(hKeyClsIdInprocServer);
	ReleaseRegKey(hKeyClsIdInprocServerVersion);
	ReleaseRegKey(hKeyClsId);

	return hr;
}

HRESULT UnregisterCOMAtRoot(
	LPCWSTR strClsId,
    LPCWSTR strProgId,
	LPCWSTR strAsmVersion,
	HKEY hKeyRoot,
	REG_KEY_BITNESS iBitness)
{
	HRESULT hr = S_OK;

	LPWSTR wszRegPathProgIdClass = NULL;

	LPWSTR wszRegPath = NULL;
	LPWSTR wszClsId = NULL;
	LPWSTR wszClsIdInprocServer = NULL;
	LPWSTR wszClsIdProgId = NULL;
	LPWSTR wszClsIdInprocServerVersion = NULL;

	hr = StrAllocFormatted(&wszClsId, L"Software\\Classes\\CLSID\\%ls", strClsId);
	ExitOnFailure(hr, "Failed to allocate string for registry path.");

	hr = StrAllocFormatted(&wszClsIdInprocServer, L"%ls\\InprocServer32", wszClsId);
	ExitOnFailure(hr, "Failed to allocate string for registry path.");

	hr = StrAllocFormatted(&wszClsIdInprocServerVersion, L"%ls\\%ls", wszClsIdInprocServer, strAsmVersion);
	ExitOnFailure(hr, "Failed to allocate string for registry path.");

	hr = RegDelete(hKeyRoot, wszClsIdInprocServerVersion, iBitness, FALSE);
	ExitOnFailure(hr, "Failed to deletethe registry key %ls", wszClsIdInprocServerVersion);

	hr = RegDelete(hKeyRoot, wszClsIdInprocServer, iBitness, FALSE);
	ExitOnFailure(hr, "Failed to deletethe registry key %ls", wszClsIdInprocServer);

    if (strProgId && *strProgId)
    {
		hr = StrAllocFormatted(&wszClsIdProgId, L"%ls\\ProgId", wszClsId);
		ExitOnFailure(hr, "Failed to allocate string for registry path.");

		hr = RegDelete(hKeyRoot, wszClsIdProgId, iBitness, FALSE);
		ExitOnFailure(hr, "Failed to deletethe registry key %ls", wszClsIdProgId);

		hr = StrAllocFormatted(&wszRegPathProgIdClass, L"%ls\\CLSID", strProgId);
		ExitOnFailure(hr, "Failed to allocate string for registry path.");

		hr = RegDelete(hKeyRoot, wszRegPathProgIdClass, REG_KEY_DEFAULT, FALSE);
		ExitOnFailure(hr, "Failed to allocate string for registry path.");

		hr = RegDelete(hKeyRoot, strProgId, iBitness, FALSE);
		ExitOnFailure(hr, "Failed to allocate string for registry path.");
    }

	hr = RegDelete(hKeyRoot, wszClsId, iBitness, FALSE);
	ExitOnFailure(hr, "Failed to deletethe registry key %ls", wszClsId);

LExit:
	ReleaseStr(wszRegPath);
	ReleaseStr(wszRegPathProgIdClass);
	ReleaseStr(wszClsIdInprocServer);
	ReleaseStr(wszClsIdInprocServerVersion);

	return hr;
}

HRESULT CreateOfficeRegistryKey(
	LPCWSTR pwzId, 
	LPCWSTR pwzFile, 
	LPCWSTR pwzFriendlyName, 
	LPCWSTR pwzDescription,
	int iCommandLineSafe,
	int iLoadBehavior,
	BOOL fPerUserInstall, 
	int iBitness)
{
	HRESULT hr = S_OK;

	if (fPerUserInstall)
	{
		hr = CreateOfficeRegistryKeyAtRoot(pwzId, pwzFile, pwzFriendlyName, pwzDescription, iCommandLineSafe, iLoadBehavior, HKEY_CURRENT_USER, REG_KEY_DEFAULT);
		ExitOnFailure1(hr, "failed to register addin (HKCU): %ls", pwzId);
	}
	else
	{
		if (iBitness == REG_KEY_32BIT || iBitness == REG_KEY_DEFAULT)
		{
			hr = CreateOfficeRegistryKeyAtRoot(pwzId, pwzFile, pwzFriendlyName, pwzDescription, iCommandLineSafe, iLoadBehavior, HKEY_LOCAL_MACHINE, REG_KEY_32BIT);
			ExitOnFailure1(hr, "failed to register addin (HKLM, 32bit): %ls", pwzId);
		}
		if (iBitness == REG_KEY_64BIT || iBitness == REG_KEY_DEFAULT)
		{
			hr = CreateOfficeRegistryKeyAtRoot(pwzId, pwzFile, pwzFriendlyName, pwzDescription, iCommandLineSafe, iLoadBehavior, HKEY_LOCAL_MACHINE, REG_KEY_64BIT);
			ExitOnFailure1(hr, "failed to register addin (HKLM, 64bit): %ls", pwzId);
		}
	}

LExit:
	return hr;
}

HRESULT DeleteOfficeRegistryKey(
	LPCWSTR pwzId, 
	BOOL fPerUserInstall, 
	int iBitness)
{
	HRESULT hr = S_OK;

	if (fPerUserInstall)
	{
		hr = DeleteOfficeRegistryKeyAtRoot(pwzId, HKEY_CURRENT_USER, REG_KEY_DEFAULT);
		ExitOnFailure1(hr, "failed to unregister addin (HKCU): %ls", pwzId);
	}
	else
	{
		if (iBitness == REG_KEY_32BIT || iBitness == REG_KEY_DEFAULT)
		{
			hr = DeleteOfficeRegistryKeyAtRoot(pwzId, HKEY_LOCAL_MACHINE, REG_KEY_32BIT);
			ExitOnFailure1(hr, "failed to unregister addin (HKLM, 32bit): %ls", pwzId);
		}
		if (iBitness == REG_KEY_64BIT || iBitness == REG_KEY_DEFAULT)
		{
			hr = DeleteOfficeRegistryKeyAtRoot(pwzId, HKEY_LOCAL_MACHINE, REG_KEY_64BIT);
			ExitOnFailure1(hr, "failed to unregister addin (HKLM, 64bit): %ls", pwzId);
		}
	}

LExit:
	return hr;
}

HRESULT RegisterCOM(
    LPCWSTR strClsId,
    LPCWSTR strProgId,
    LPCWSTR typeFullName, 
    LPCWSTR strAsmName, 
    LPCWSTR strAsmVersion, 
    LPCWSTR strAsmCodeBase, 
    LPCWSTR strRuntimeVersion,
	BOOL fPerUserInstall, 
	int iBitness)
{
	HRESULT hr = S_OK;

	if (fPerUserInstall)
	{
		hr = RegisterCOMAtRoot(strClsId, strProgId, typeFullName, strAsmName, strAsmVersion, strAsmCodeBase, strRuntimeVersion,
			HKEY_CURRENT_USER, REG_KEY_DEFAULT);
		ExitOnFailure1(hr, "failed to register COM");
	}
	else
	{
		if (iBitness == REG_KEY_32BIT || iBitness == REG_KEY_DEFAULT)
		{
			hr = RegisterCOMAtRoot(strClsId, strProgId, typeFullName, strAsmName, strAsmVersion, strAsmCodeBase, strRuntimeVersion,
				HKEY_LOCAL_MACHINE, REG_KEY_32BIT);
			ExitOnFailure1(hr, "failed to register COM");
		}
		if (iBitness == REG_KEY_64BIT || iBitness == REG_KEY_DEFAULT)
		{
			hr = RegisterCOMAtRoot(strClsId, strProgId, typeFullName, strAsmName, strAsmVersion, strAsmCodeBase, strRuntimeVersion,
				HKEY_LOCAL_MACHINE, REG_KEY_64BIT);
			ExitOnFailure1(hr, "failed to register COM");
		}
	}

LExit:
	return hr;
}

HRESULT UnregisterCOM(
	LPCWSTR strClsId,
    LPCWSTR strProgId,
	LPCWSTR strAsmVersion,
	BOOL fPerUserInstall, 
	int iBitness)
{
	HRESULT hr = S_OK;

	if (fPerUserInstall)
	{
		hr = UnregisterCOMAtRoot(strClsId, strProgId, strAsmVersion, HKEY_CURRENT_USER, REG_KEY_DEFAULT);
		ExitOnFailure1(hr, "failed to unregister COM");
	}
	else
	{
		if (iBitness == REG_KEY_32BIT || iBitness == REG_KEY_DEFAULT)
		{
			hr = UnregisterCOMAtRoot(strClsId, strProgId, strAsmVersion, HKEY_LOCAL_MACHINE, REG_KEY_32BIT);
			ExitOnFailure1(hr, "failed to unregister COM");
		}
		if (iBitness == REG_KEY_64BIT || iBitness == REG_KEY_DEFAULT)
		{
			hr = UnregisterCOMAtRoot(strClsId, strProgId, strAsmVersion, HKEY_LOCAL_MACHINE, REG_KEY_64BIT);
			ExitOnFailure1(hr, "failed to unregister COM");
		}
	}

LExit:
	return hr;
}
