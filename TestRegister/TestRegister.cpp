// TestRegister.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include <windef.h>

#include "../VisioWixExtension/ca/Register.h"
#pragma comment(lib, "../Debug/VisioCustomActions.lib")

//HRESULT CreateOfficeRegistryKey(
//	LPCWSTR pwzId, 
//	LPCWSTR pwzFile, 
//	LPCWSTR pwzFriendlyName, 
//	LPCWSTR pwzDescription,
//	int iCommandLineSafe,
//	int iLoadBehavior,
//	BOOL fPerUserInstall, 
//	int iBitness);
//
//HRESULT DeleteOfficeRegistryKey(
//	LPCWSTR pwzId, 
//	BOOL fPerUserInstall, 
//	int iBitness);

int main(int argc, wchar_t* argv[])
{
	RegisterCOM(
		L"{1CFBA542-667B-4E57-A739-343819CC84E8}", L"MyTestApp.Id", 
		L"strTypeName",  L"strAsmName", L"1.0.0.0", L"strAsmCodeBase", L"strRuntimeVersion", TRUE, 0);

	UnregisterCOM(
		L"{1CFBA542-667B-4E57-A739-343819CC84E8}", L"MyTestApp.Id", L"1.0.0.0", TRUE, 0);
}