
#pragma once

HRESULT CreateOfficeRegistryKey(
	LPCWSTR pwzId, 
	LPCWSTR pwzFile, 
	LPCWSTR pwzFriendlyName, 
	LPCWSTR pwzDescription,
	int iCommandLineSafe,
	int iLoadBehavior,
	BOOL fPerUserInstall, 
	int iBitness);

HRESULT DeleteOfficeRegistryKey(
	LPCWSTR pwzId, 
	BOOL fPerUserInstall, 
	int iBitness);

HRESULT RegisterCOM(
    LPCWSTR strClsId,
    LPCWSTR strProgId,
    LPCWSTR typeFullName, 
    LPCWSTR strAsmName, 
    LPCWSTR strAsmVersion, 
    LPCWSTR strAsmCodeBase, 
    LPCWSTR strRuntimeVersion,
	BOOL fPerUserInstall, 
	int iBitness);

HRESULT UnregisterCOM(
	LPCWSTR strClsId,
    LPCWSTR strProgId,
	LPCWSTR strAsmVersion,
	BOOL fPerUserInstall, 
	int iBitness);
