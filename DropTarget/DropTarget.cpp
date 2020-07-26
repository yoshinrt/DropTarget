/*****************************************************************************

	DropTarget --- shell extension drop handler

*****************************************************************************/

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
		( pExt = wcsrchr( szFile, '.' )) > wcsrchr( szFile, '\\' ) ?
		pExt : nullptr
	);
}

/*** get a string from REGISTORY ********************************************/

wchar_t *GetRegStr(
	HKEY	hKey,
	LPTSTR	pSubKey,
	LPTSTR	pValueName,
	wchar_t	*szBuf
){
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
	
	wchar_t	szVerb[ BUFSIZE ];
	wchar_t	*pExt;
	
	if( !(
		// HKLM\.ext を取得
		( pExt = SearchExt( szFile )) &&
		GetRegStr( HKEY_CLASSES_ROOT, pExt, nullptr, szDst ) &&
		
		// HKLM\extfile\shell を取得
		wcscat_s( szDst, BUFSIZE, L"\\shell" ) == 0 &&
		GetRegStr( HKEY_CLASSES_ROOT, szDst, nullptr, szVerb ) &&
		
		// HKLM\extfile\shell\open\command を取得
		wcscat_s( szDst, BUFSIZE, L"\\" ) == 0 &&
		wcscat_s( szDst, BUFSIZE, *szVerb ? szVerb : L"open" ) == 0 &&
		wcscat_s( szDst, BUFSIZE, L"\\command" )
	)){
		wcscpy_s( szDst, BUFSIZE, L"Unknown\\shell\\open\\command" );
	}
	
	GetRegStr( HKEY_CLASSES_ROOT, szDst, nullptr, szDst );
	return( szDst );
}

/*** DllMain ****************************************************************/

int WINAPI DllMain(
	HINSTANCE	hInstance,
	DWORD		dwReason,
	LPVOID		lpReserved
){
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
	LPVOID FAR	*ppv
){
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
	LPVOID		*ppvObj
){
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
	m_pszFileUserClickedOn = nullptr;
}

CShellExt::~CShellExt(){
	g_cRefThisDll--;
	
	if( m_pszFileUserClickedOn ){
		delete [] m_pszFileUserClickedOn;
		m_pszFileUserClickedOn = nullptr;
	}
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
	
	if( m_pszFileUserClickedOn ){
		delete [] m_pszFileUserClickedOn;
		m_pszFileUserClickedOn = nullptr;
	}
	
	size_t	nLen = wcslen( lpszFileName );
	m_pszFileUserClickedOn = new wchar_t[ nLen + 1 ];
	
	wcscpy_s( m_pszFileUserClickedOn, nLen + 1, lpszFileName );
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
	DWORD		*pdwEffect
){
	*pdwEffect = DROPEFFECT_MOVE;
	return( S_OK );
}

STDMETHODIMP CShellExt::DragOver(
	DWORD		grfKeyState,
	POINTL		pt,
	DWORD		*pdwEffect
){
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
	DWORD		*pdwEffect
){
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
			
			GetDefaultAction( m_pszFileUserClickedOn, szExecCmd );
			
			// %1 %* がないときは追加
			if( !wcsstr( szExecCmd, L"%1" )) wcscat_s( szExecCmd, BUFSIZE, L" \"%1\"" );
			if( !wcsstr( szExecCmd, L"%*" )) wcscat_s( szExecCmd, BUFSIZE, L" %*" );
			
			wchar_t	*pCmd	  = szExecCmd,
					*pCmdLine = szExecCmdLine;
			
			for( ; *pCmd; ++pCmd ){
				if( *pCmd == '%' ){
					switch( *++pCmd ){
					  case '1':
						if( m_pszFileUserClickedOn ){
							wcscpy_s( pCmdLine, CMDBUFSIZE, m_pszFileUserClickedOn );
						}
						pCmdLine = wcschr( pCmdLine, L'\0' );
						
					  Case '*':
						*pCmdLine = L'\0';
						
						if( pDropFiles->fWide ){
							wchar_t	*pwcDropFileName = ( wchar_t *)(( char *)pDropFiles + pDropFiles->pFiles );
							
							for(; *pwcDropFileName; pwcDropFileName += wcslen( pwcDropFileName )+ 1 ){
								wcscat_s( pCmdLine, CMDBUFSIZE, L" \"" );
								wcscpy_s( pCmdLine, CMDBUFSIZE, pwcDropFileName );
								wcscat_s( pCmdLine, CMDBUFSIZE, L"\"" );
							}
						}else{
							char *pDropFileName = ( char *)pDropFiles + pDropFiles->pFiles;
							
							for(; *pDropFileName; pDropFileName += strlen( pDropFileName )+ 1 ){
								wcscat_s( pCmdLine, CMDBUFSIZE, L" \"" );
								
								pCmdLine = wcschr( pCmdLine, L'\0' );
								MultiByteToWideChar(
									CP_ACP,				// CodePage
									0,					// dwFlags
									pDropFileName,		// lpMultiByteStr
									-1,					// cbMultiByte
									pCmdLine,			// lpWideCharStr
									CMDBUFSIZE			// cchWideChar
								);
								
								wcscat_s( pCmdLine, CMDBUFSIZE, L"\"" );
							}
						}
						pCmdLine = wcschr( pCmdLine, L'\0' );
						
					  Default:
						*pCmdLine++ = *pCmd;
					}
				}else{
					*pCmdLine++ = *pCmd;
				}
			}
			
			DebugMsgW( szExecCmdLine );
			
			if( m_pszFileUserClickedOn ){
				delete [] m_pszFileUserClickedOn;
				m_pszFileUserClickedOn = nullptr;
			}
			
			/*** create process *********************************************/
			
			GetStartupInfo( &si );
			CreateProcess(
				nullptr, szExecCmdLine,			/* exec file, params		*/
				nullptr, nullptr,				/* securities				*/
				FALSE,							/* inherit flag				*/
				NORMAL_PRIORITY_CLASS,			/* creation flags			*/
				nullptr, nullptr,				/* env, cur.dir.			*/
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
