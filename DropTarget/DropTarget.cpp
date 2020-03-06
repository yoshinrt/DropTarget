/*****************************************************************************

	DropTarget --- shell extension drop handler

*****************************************************************************/

#define _CRT_SECURE_NO_WARNINGS

#include <windows.h>
#include <windowsx.h>
#include <shlobj.h>
#include <string.h>
#include "dds.h"

// Initialize GUIDs (should be done only and at-least once per DLL/EXE)

#pragma data_seg( ".text" )
#define INITGUID
#include <initguid.h>
#include <shlguid.h>
#include "DropTarget.h"
#pragma data_seg()

/*** constant definition ****************************************************/

#define BUFSIZE		256
#define CMDBUFSIZE	1024

/*** gloval var definition **************************************************/

UINT		g_cRefThisDll  = 0;			// Reference count of this DLL.
HINSTANCE	g_hInstThisDll = nullptr;	// Handle to this DLL itself.

/*** search a ext ***********************************************************/

wchar_t *SearchExt( wchar_t *szFile ){
	wchar_t	*pExt;
	
	return(
		( UINT )( pExt = wcsrchr( szFile, '.' )) >
		( UINT )wcsrchr( szFile, '\\' ) ?
		pExt : nullptr
	);
}

/*** get a string from REGISTORY ********************************************/

wchar_t *GetRegStr(
	HKEY	hKey,
	LPTSTR	pSubKey,
	LPTSTR	pValueName,
	wchar_t	*szBuf ){
	
	HKEY	hSubKey;					/* for getting default action		*/
	DWORD	dwBufSize = BUFSIZE;		/* data buf size for RegQueryValue	*/
	
	if( RegOpenKeyEx( hKey, pSubKey, 0, KEY_READ, &hSubKey )== ERROR_SUCCESS ){
		LONG lResult = RegQueryValueEx(
			hSubKey, pValueName, nullptr, nullptr,
			( LPBYTE )szBuf, &dwBufSize
		);
		RegCloseKey( hSubKey );
		return( lResult == ERROR_SUCCESS ? szBuf : nullptr );
	}
	return( nullptr );
}

/*** get default action of the file type ************************************/

wchar_t *GetDefaultAction( wchar_t *szFile, wchar_t *szDst ){
	
	wchar_t	szVerb[ BUFSIZE ],
			*pExt;
	
	if(( pExt = SearchExt( szFile ))&&
		GetRegStr( HKEY_CLASSES_ROOT, pExt, nullptr, szDst )){
		
		wcscat( szDst, L"\\shell" );
		
		if( GetRegStr( HKEY_CLASSES_ROOT, szDst, nullptr, szVerb )){
			wcscat( wcscat( wcscat(
				szDst, L"\\" ), *szVerb ? szVerb : L"open" ), L"\\command" );
			goto GetAction;
		}
	}
	wcscpy( szDst, L"Unknown\\shell\\open\\command" );
	
  GetAction:
	GetRegStr( HKEY_CLASSES_ROOT, szDst, nullptr, szDst );
	return( szDst );
}

/*** DllMain ****************************************************************/

int WINAPI DllMain(
	HINSTANCE	hInstance,
	DWORD		dwReason,
	LPVOID		lpReserved ){
	
	if( dwReason == DLL_PROCESS_ATTACH ) g_hInstThisDll = hInstance;
	return( 1 );	// ok
}

/*** DllCanUnloadNow & DllGetClassObject ************************************/

STDAPI DllCanUnloadNow( void ){
    return( g_cRefThisDll ? S_FALSE : S_OK );
}

STDAPI DllGetClassObject( REFCLSID rclsid, REFIID riid, LPVOID *ppvOut ){
	
	*ppvOut = nullptr;
	
	if( IsEqualIID( rclsid, CLSID_DropTarget )){
		CShellExtClassFactory *pcf = new CShellExtClassFactory;
		return( pcf->QueryInterface( riid, ppvOut ));
	}
	
	return( CLASS_E_CLASSNOTAVAILABLE );
}

/*** CShellExtClassFactory methods ******************************************/

CShellExtClassFactory::CShellExtClassFactory(){
	m_cRef = 0L;
	g_cRefThisDll++;
}

CShellExtClassFactory::~CShellExtClassFactory(){
	g_cRefThisDll--;
}

STDMETHODIMP CShellExtClassFactory::QueryInterface(
	REFIID		riid,
	LPVOID FAR	*ppv ){
	
	*ppv = nullptr;
	
	// Any interface on this object is the object pointer
	if(
		IsEqualIID( riid, IID_IUnknown )||
		IsEqualIID( riid, IID_IClassFactory )
	){
		*ppv =( LPCLASSFACTORY )this;
		AddRef();
		return( NOERROR );
	}
	return( E_NOINTERFACE );
}

STDMETHODIMP_(ULONG) CShellExtClassFactory::AddRef(){
	return ++m_cRef;
}

STDMETHODIMP_(ULONG) CShellExtClassFactory::Release(){
	if ( --m_cRef ) return( m_cRef );
	
	delete this;
	return 0L;
}

STDMETHODIMP CShellExtClassFactory::CreateInstance(
	LPUNKNOWN	pUnkOuter,
	REFIID		riid,
	LPVOID		*ppvObj ){
	
	*ppvObj = nullptr;
	
	// Shell extensions typically don't support aggregation (inheritance)
	
	if( pUnkOuter )	return( CLASS_E_NOAGGREGATION );
	
	// Create the main shell extension object.  The shell will then call
	// QueryInterface with IID_IShellExtInit--this is how shell extensions are
	// initialized.
	
	LPCSHELLEXT pShellExt = new CShellExt();  //Create the CShellExt object
	
	if( nullptr == pShellExt )	return( E_OUTOFMEMORY );
	
	return( pShellExt->QueryInterface( riid, ppvObj ));
}


STDMETHODIMP CShellExtClassFactory::LockServer( BOOL fLock ){
	return( NOERROR );
}

/*** IUnKnown ***************************************************************/

CShellExt::CShellExt(){
	m_cRef = 0L;
	g_cRefThisDll++;
}

CShellExt::~CShellExt(){
	g_cRefThisDll--;
}

STDMETHODIMP CShellExt::QueryInterface( REFIID riid, LPVOID FAR *ppv ){
	
	*ppv = nullptr;
	if     ( IsEqualIID( riid, IID_IUnknown     )) *ppv =( LPSHELLEXTINIT )this;
	else if( IsEqualIID( riid, IID_IPersistFile )) *ppv =( LPPERSISTFILE  )this;
	else if( IsEqualIID( riid, IID_IDropTarget  )) *ppv =( LPDROPTARGET   )this;
	
	if( *ppv ){
		AddRef();
		return NOERROR;
	}
	
	return E_NOINTERFACE;
}

STDMETHODIMP_( ULONG ) CShellExt::AddRef(){
	return ++m_cRef;
}

STDMETHODIMP_( ULONG ) CShellExt::Release(){
	if ( --m_cRef ) return m_cRef;
	
	delete this;
	return 0L;
}

/*** IPersistFile ***********************************************************/

STDMETHODIMP CShellExt::GetClassID( LPCLSID lpClassID ){
	return E_FAIL;
}

STDMETHODIMP CShellExt::IsDirty(){
	return S_FALSE;
}

STDMETHODIMP CShellExt::Load( LPCOLESTR lpszFileName, DWORD grfMode ){
	wcscpy( m_szFileUserClickedOn, lpszFileName );
	return NOERROR;
}

STDMETHODIMP CShellExt::Save( LPCOLESTR lpszFileName, BOOL fRemember ){
	return E_FAIL;
}

STDMETHODIMP CShellExt::SaveCompleted( LPCOLESTR lpszFileName ){
	return E_FAIL;
}

STDMETHODIMP CShellExt::GetCurFile( LPOLESTR FAR* lplpszFileName ){
	return E_FAIL;
}

/*** IDropTarget ************************************************************/

STDMETHODIMP CShellExt::DragEnter(
	IDataObject	*pDataObject,
	DWORD		grfKeyState,
	POINTL		pt,
	DWORD		*pdwEffect ){
	
	*pdwEffect = DROPEFFECT_MOVE;
	return( S_OK );
}

STDMETHODIMP CShellExt::DragOver(
	DWORD		grfKeyState,
	POINTL		pt,
	DWORD		*pdwEffect ){
	
	*pdwEffect = DROPEFFECT_MOVE;
	return( S_OK );
}

STDMETHODIMP CShellExt::DragLeave( void ){
	return( S_OK );
}

STDMETHODIMP CShellExt::Drop(
	IDataObject	*pDataObject,
	DWORD		grfKeyState,
	POINTL		pt,
	DWORD		*pdwEffect ){
	
	FORMATETC	fmtetc;
	STGMEDIUM	stgmed;
	HRESULT		hrErr;
	
	wchar_t	szExecCmd[ BUFSIZE ],
			szExecCmdLine[ CMDBUFSIZE ];
	
	STARTUPINFO			si;
	PROCESS_INFORMATION	pi;
	
	// do a GetData for CF_OBJECTDESCRIPTOR format to get the size of the
	// object as displayed in the source. using this size, initialize the
	// size for the drag feedback rectangle.
	
	fmtetc.cfFormat	= CF_HDROP;
	fmtetc.ptd		= nullptr;
	fmtetc.lindex	= -1;
	fmtetc.dwAspect	= DVASPECT_CONTENT;
	fmtetc.tymed	= TYMED_HGLOBAL;
	
	hrErr = pDataObject->GetData( &fmtetc, &stgmed );
	if( hrErr == NOERROR ){
		
		LPDROPFILES pDropFiles =( LPDROPFILES )GlobalLock( stgmed.hGlobal );
		
		if( pDropFiles ){
			/*** make command line ******************************************/
			
			GetDefaultAction( m_szFileUserClickedOn, szExecCmd );
			if( !wcsstr( szExecCmd, L"%1" )) wcscat( szExecCmd, L" \"%1\"" );
			if( !wcsstr( szExecCmd, L"%*" )) wcscat( szExecCmd, L" %*" );
			
			wchar_t	*pDropFileName,
					*pCmd	  = szExecCmd,
					*pCmdLine = szExecCmdLine;
			
			WCHAR	*pwcDropFileName;
			
			do{
				if( *pCmd == '%' ){
					switch( *++pCmd ){
					  case '1':
						wcscpy( pCmdLine, m_szFileUserClickedOn );
						pCmdLine = wcschr( pCmdLine, '\0' );
						
					  Case '*':
						*pCmdLine = '\0';
						
						if( pDropFiles->fWide ){
							pwcDropFileName =( WCHAR *)(( wchar_t *)pDropFiles + pDropFiles->pFiles );
							
							for(; *pwcDropFileName; pwcDropFileName += wcslen( pwcDropFileName )+ 1 ){
								wcscat( pCmdLine, L" \"" );
								pCmdLine = wcschr( pCmdLine, '\0' );
								wcscpy( pCmdLine, pwcDropFileName );
								wcscat( pCmdLine, L"\"" );
							}
						}else{
							pDropFileName =( wchar_t *)pDropFiles + pDropFiles->pFiles;
							
							for(; *pDropFileName; pDropFileName += wcslen( pDropFileName )+ 1 ){
								wcscat( pCmdLine, L" \"" );
								wcscat( pCmdLine, pDropFileName );
								wcscat( pCmdLine, L"\"" );
							}
						}
						pCmdLine = wcschr( pCmdLine, '\0' );
						
					  Default:
						*pCmdLine++ = *pCmd;
					}
				}else{
					*pCmdLine++ = *pCmd;
				}
			}while( *pCmd++ );
			
			DebugMsgW( szExecCmdLine );
			
			/*** create process *********************************************/
			
			GetStartupInfo( &si );
			CreateProcess(
				nullptr, szExecCmdLine,			/* exec file, params		*/
				nullptr, nullptr,						/* securities				*/
				FALSE,							/* inherit flag				*/
				NORMAL_PRIORITY_CLASS,			/* creation flags			*/
				nullptr, nullptr,						/* env, cur.dir.			*/
				&si, &pi
			);
			CloseHandle( pi.hProcess );
			CloseHandle( pi.hThread );
		}
		GlobalUnlock( stgmed.hGlobal );
		ReleaseStgMedium( &stgmed );
	}
	
	*pdwEffect = DROPEFFECT_MOVE;
	return( S_OK );
}
