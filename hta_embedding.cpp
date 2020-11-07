// Launches an hta file via COM, using ProgID: htafile
// triggers the launch via ipersistmoniker

#include <ole2.h>
#include <comdef.h>
#pragma comment(lib, "Urlmon.lib")

int wmain(int argc, wchar_t* argv[], wchar_t* envp[])
{
	GUID gHtafile = { 0x3050f4d8,0x98b5,0x11cf,{ 0xbb,0x82,0x00,0xaa,0x00,0xbd,0xce,0x0b } };
	HRESULT hr;
	LPMONIKER pFmoniker;
	IPersistMonikerPtr pPersist;

	WCHAR pwszHta[MAX_PATH];

	wcscpy_s(pwszHta, MAX_PATH, L"c:\\users\\user\\desktop\\test.hta");

	hr = CoInitialize(0);

	hr = CreateURLMonikerEx(NULL, pwszHta, &pFmoniker, URL_MK_UNIFORM);
	hr = CoCreateInstance(gHtafile, NULL, CLSCTX_LOCAL_SERVER, IID_IPersistMoniker, reinterpret_cast<void**>(&pPersist));

	if (hr != S_OK)
	{
		printf("Creating htafile COM object failed");
	}

	hr = pPersist->Load(FALSE, pFmoniker, NULL, STGM_READ);

	pFmoniker = NULL;
	pPersist = NULL;
}
