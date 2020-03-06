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

UINT		g_cRefThisDll  = 0;		// Reference count of this DLL.
HINSTANCE	g_hInstThisDll = NULL;	// Handle to this DLL itself.

/*** search a ext ***********************************************************/

char *SearchExt( char *szFile ){
	char	*pExt;
	
	return(
		( UINT )( pExt = strrchr( szFile, '.' )) >
		( UINT )strrchr( szFile, '\\' ) ?
		pExt : NULL
	);
}

/*** get a string from REGISTORY ********************************************/

char *GetRegStr(
	HKEY	hKey,
	LPTSTR	pSubKey,
	LPTSTR	pValueName,
	char	*szBuf ){
	
	HKEY	hSubKey;					/* for getting default action		*/
	DWORD	dwBufSize = BUFSIZE;		/* data buf size for RegQueryValue	*/
	
	if( RegOpenKeyEx( hKey, pSubKey, 0, KEY_READ, &hSubKey )== ERROR_SUCCESS ){
		LONG lResult = RegQueryValueEx(
			hSubKey, pValueName, NULL, NULL,
			( LPBYTE )szBuf, &dwBufSize
		);
		RegCloseKey( hSubKey );
		return( lResult == ERROR_SUCCESS ? szBuf : NULL );
	}
	return( NULL );
}

/*** get default action of the file type ************************************/

char *GetDefaultAction( char *szFile, char *szDst ){
	
	char	szVerb[ BUFSIZE ],
			*pExt;
	
	if(( pExt = SearchExt( szFile ))&&
		GetRegStr( HKEY_CLASSES_ROOT, pExt, NULL, szDst )){
		
		strcat( szDst, "\\shell" );
		
		if( GetRegStr( HKEY_CLASSES_ROOT, szDst, NULL, szVerb )){
			strcat( strcat( strcat(
				szDst, "\\" ), *szVerb ? szVerb : "open" ), "\\command" );
			goto GetAction;
		}
	}
	strcpy( szDst, "Unknown\\shell\\open\\command" );
	
  GetAction:
	GetRegStr( HKEY_CLASSES_ROOT, szDst, NULL, szDst );
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
	
	*ppvOut = NULL;
	
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
	
	*ppv = NULL;
	
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
	
	*ppvObj = NULL;
	
	// Shell extensions typically don't support aggregation (inheritance)
	
	if( pUnkOuter )	return( CLASS_E_NOAGGREGATION );
	
	// Create the main shell extension object.  The shell will then call
	// QueryInterface with IID_IShellExtInit--this is how shell extensions are
	// initialized.
	
	LPCSHELLEXT pShellExt = new CShellExt();  //Create the CShellExt object
	
	if( NULL == pShellExt )	return( E_OUTOFMEMORY );
	
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
	
	*ppv = NULL;
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
	WideCharToMultiByte(
		CP_ACP,								// CodePage
		0,									// dwFlags
		lpszFileName,						// lpWideCharStr
		-1,									// cchWideChar
		m_szFileUserClickedOn,				// lpMultiByteStr
		sizeof( m_szFileUserClickedOn ),	// cchMultiByte,
		NULL,								// lpDefaultChar,
		NULL );								// lpUsedDefaultChar
	
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
	
	char	szExecCmd[ BUFSIZE ],
			szExecCmdLine[ CMDBUFSIZE ];
	
	STARTUPINFO			si;
	PROCESS_INFORMATION	pi;
	
	// do a GetData for CF_OBJECTDESCRIPTOR format to get the size of the
	// object as displayed in the source. using this size, initialize the
	// size for the drag feedback rectangle.
	
	fmtetc.cfFormat	= CF_HDROP;
	fmtetc.ptd		= NULL;
	fmtetc.lindex	= -1;
	fmtetc.dwAspect	= DVASPECT_CONTENT;
	fmtetc.tymed	= TYMED_HGLOBAL;
	
	hrErr = pDataObject->GetData( &fmtetc, &stgmed );
	if( hrErr == NOERROR ){
		
		LPDROPFILES pDropFiles =( LPDROPFILES )GlobalLock( stgmed.hGlobal );
		
		if( pDropFiles ){
			/*** make command line ******************************************/
			
			GetDefaultAction( m_szFileUserClickedOn, szExecCmd );
			if( !strstr( szExecCmd, "%1" )) strcat( szExecCmd, " \"%1\"" );
			if( !strstr( szExecCmd, "%*" )) strcat( szExecCmd, " %*" );
			
			char	*pDropFileName,
					*pCmd	  = szExecCmd,
					*pCmdLine = szExecCmdLine;
			
			WCHAR	*pwcDropFileName;
			
			do{
				if( *pCmd == '%' ){
					switch( *++pCmd ){
					  case '1':
						strcpy( pCmdLine, m_szFileUserClickedOn );
						pCmdLine = strchr( pCmdLine, '\0' );
						
					  Case '*':
						*pCmdLine = '\0';
						
						if( pDropFiles->fWide ){
							pwcDropFileName =( WCHAR *)(( char *)pDropFiles + pDropFiles->pFiles );
							
							for(; *pwcDropFileName; pwcDropFileName += wcslen( pwcDropFileName )+ 1 ){
								strcat( pCmdLine, " \"" );
								pCmdLine = strchr( pCmdLine, '\0' );
								
								WideCharToMultiByte(
									CP_ACP,				// CodePage
									0,					// dwFlags
									pwcDropFileName,	// lpWideCharStr
									-1,					// cchWideChar
									pCmdLine,			// lpMultiByteStr
									CMDBUFSIZE,			// cchMultiByte,
									NULL,				// lpDefaultChar,
									NULL );				// lpUsedDefaultChar
								
								strcat( pCmdLine, "\"" );
							}
						}else{
							pDropFileName =( char *)pDropFiles + pDropFiles->pFiles;
							
							for(; *pDropFileName; pDropFileName += strlen( pDropFileName )+ 1 ){
								strcat( pCmdLine, " \"" );
								strcat( pCmdLine, pDropFileName );
								strcat( pCmdLine, "\"" );
							}
						}
						pCmdLine = strchr( pCmdLine, '\0' );
						
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
				NULL, szExecCmdLine,			/* exec file, params		*/
				NULL, NULL,						/* securities				*/
				FALSE,							/* inherit flag				*/
				NORMAL_PRIORITY_CLASS,			/* creation flags			*/
				NULL, NULL,						/* env, cur.dir.			*/
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
