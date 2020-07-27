/*****************************************************************************

	DropTarget --- shell extension drop handler

*****************************************************************************/

#include <windows.h>
#include <windowsx.h>
#include <shlobj.h>
#include <wchar.h>
#include "dds.h"

// Initialize GUIDs (should be done only and at-least once per DLL/EXE)

#pragma data_seg( ".text" )
#define INITGUID
#include <initguid.h>
#include <shlguid.h>
#include "DropTarget.h"
#pragma data_seg()

/*** constant definition ****************************************************/

#define BUFSIZE	256

/*** gloval var definition **************************************************/

UINT		g_cRefThisDll  = 0;			// Reference count of this DLL.
HINSTANCE	g_hInstThisDll = nullptr;	// Handle to this DLL itself.

/*** ローレベル string ******************************************************/

class CBuffer {
public:
	CBuffer( size_t iSize = 0 ){
		
		if( iSize ){
			m_pBuf = new wchar_t[ iSize ];
			*m_pBuf = L'\0';
		}
		else m_pBuf = nullptr;
		
		m_iSize		= iSize;
		m_iLength	= 0;
	}
	
	~CBuffer(){
		if( m_pBuf ) delete [] m_pBuf;
		m_pBuf = nullptr;
	}
	
	// iSize 以上で 2^n サイズ確保する
	BOOL Resize( size_t iSize ){
		if( iSize <= m_iSize ) return TRUE;	// すでに確保済み
		
		// 2^n サイズ
		for( m_iSize = BUFSIZE; m_iSize < iSize; m_iSize <<= 1 ){
			// 最大 16MB
			if( iSize > 8 * 1024 * 1024 ){
				m_iSize >>= 1;
				return FALSE;
			}
		}
		
		wchar_t *pTmp = new wchar_t[ m_iSize ];
		
		// コピー
		if( m_pBuf ){
			wmemcpy_s( pTmp, m_iSize, m_pBuf, m_iLength + 1 );
			delete [] m_pBuf;
		}else{
			*pTmp		= L'\0';
			m_iLength	= 0;
		}
		m_pBuf = pTmp;
		
		return TRUE;
	}
	
	// cat
	CBuffer& operator += ( const wchar_t *b ){
		// バッファ伸長
		size_t iStrLen = wcslen( b );
		size_t iRequiredSize = m_iLength + iStrLen + 1;
		if( m_iSize < iRequiredSize ) Resize( iRequiredSize );
		
		wmemcpy_s( m_pBuf + m_iLength, m_iSize - m_iLength, b, iStrLen + 1 );
		m_iLength += iStrLen;
		
		return *this;
	}
	
	CBuffer& operator += ( const wchar_t b ){
		if( !Resize( m_iLength + 2 )) return *this;
		
		m_pBuf[ m_iLength++ ] = b;
		m_pBuf[ m_iLength   ] = L'\0';
		
		return *this;
	}
	
	CBuffer& Set( const wchar_t *b ){
		if( m_pBuf ) *m_pBuf = L'\0';
		*this += b;
		return *this;
	}
	
	// メンバ変数
	
	size_t	m_iSize;
	size_t	m_iLength;
	wchar_t* m_pBuf;
};

/*** search a ext ***********************************************************/

wchar_t *SearchExt( wchar_t *szFile ){
	wchar_t	*pExt;
	
	return(
		( pExt = wcsrchr( szFile, '.' )) > wcsrchr( szFile, '\\' ) ?
		pExt : nullptr
	);
}

/*** get a string from REGISTORY ********************************************/

BOOL GetRegStr(
	HKEY	hKey,
	LPTSTR	pSubKey,
	LPTSTR	pValueName,
	CBuffer	&strBuf
){
	HKEY	hSubKey;	/* for getting default action		*/
	DWORD	dwBufSize;	/* data buf size for RegQueryValue	*/
	
	if( RegOpenKeyEx( hKey, pSubKey, 0, KEY_READ, &hSubKey ) != ERROR_SUCCESS ) return FALSE;
	
	LONG	lResult;
	DWORD	dwType;
	
	while( 1 ){
		dwBufSize = ( DWORD )( strBuf.m_iSize * 2 );
		
		if(
			( lResult = RegQueryValueEx(
				hSubKey, pValueName, nullptr, &dwType,
				( LPBYTE )strBuf.m_pBuf, &dwBufSize
			)) != ERROR_MORE_DATA &&
			strBuf.m_pBuf != nullptr
		) break;
		
		if( !strBuf.Resize( dwBufSize / 2 )){
			lResult = ~ERROR_SUCCESS;
			break;
		}
	}
	
	RegCloseKey( hSubKey );
	
	strBuf.m_iLength = dwBufSize / 2 - 1;
	
	if( lResult != ERROR_SUCCESS ) return FALSE;
	
	// 環境変数展開
	if( dwType == REG_EXPAND_SZ ){
		wchar_t *szPreExpand = strBuf.m_pBuf;
		
		strBuf.m_pBuf		= nullptr;
		strBuf.m_iSize		= 0;
		strBuf.m_iLength	= 0;
		
		strBuf.Resize( ExpandEnvironmentStrings( szPreExpand, nullptr, 0 ));
		ExpandEnvironmentStrings( szPreExpand, strBuf.m_pBuf, strBuf.m_iSize );
		
		delete [] szPreExpand;
	}
	
	return TRUE;
}

/*** get default action of the file type ************************************/

BOOL GetDefaultAction( wchar_t *szFile, CBuffer &strDst ){
	
	CBuffer strVerb( BUFSIZE );
	wchar_t	*pExt;
	
	if(
		// HKLM\.ext を取得
		( pExt = SearchExt( szFile )) &&
		GetRegStr( HKEY_CLASSES_ROOT, pExt, nullptr, strDst )
	){
		strDst += L"\\shell";
		BOOL bVerb = GetRegStr( HKEY_CLASSES_ROOT, strDst.m_pBuf, nullptr, strVerb );
		
		// HKLM\extfile\shell\open\command を取得
		(( strDst += L"\\" ) += ( bVerb ? strVerb.m_pBuf : L"open" )) += L"\\command";
	}else{
		// unknown を取得
		strDst.Set( L"Unknown\\shell\\open\\command" );
	}
	
	return GetRegStr( HKEY_CLASSES_ROOT, strDst.m_pBuf, nullptr, strDst );
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
			CBuffer strExecFile( BUFSIZE );
			CBuffer strCmdLine( BUFSIZE );
			
			/*** make command line ******************************************/
			
			GetDefaultAction( m_pszFileUserClickedOn, strExecFile );
			
			// %1 %* がないときは追加
			if( !wcsstr( strExecFile.m_pBuf, L"%1" )) strExecFile += L" \"%1\"";
			if( !wcsstr( strExecFile.m_pBuf, L"%*" )) strExecFile += L" %*";
			
			wchar_t	*pCmd	  = strExecFile.m_pBuf;
			
			for( ; *pCmd; ++pCmd ){
				if( *pCmd == '%' ){
					switch( *++pCmd ){
					  case '1':
						if( m_pszFileUserClickedOn ){
							strCmdLine += m_pszFileUserClickedOn;
						}
						
					  Case '*':
						if( pDropFiles->fWide ){
							wchar_t	*pwcDropFileName = ( wchar_t *)(( char *)pDropFiles + pDropFiles->pFiles );
							
							while( *pwcDropFileName ){
								size_t iFileLen = wcslen( pwcDropFileName );
								if( !strCmdLine.Resize( strCmdLine.m_iLength + iFileLen + 3 + 1 )) break;
								
								(( strCmdLine += L" \"" ) += pwcDropFileName ) += L'\"';
								
								pwcDropFileName += iFileLen + 1;
							}
						}else{
							char *pDropFileName = ( char *)pDropFiles + pDropFiles->pFiles;
							
							while( *pDropFileName ){
								
								// wchar 文字数 (null を含む)
								int iFileLen = MultiByteToWideChar(
									CP_ACP,										// CodePage
									0,											// dwFlags
									pDropFileName,								// lpMultiByteStr
									-1,											// cbMultiByte
									nullptr,									// lpWideCharStr
									0											// cchWideChar
								);
								
								if( !strCmdLine.Resize( strCmdLine.m_iLength + iFileLen + 3 )) break;
								
								strCmdLine += L" \"";
								
								// 変換
								MultiByteToWideChar(
									CP_ACP,										// CodePage
									0,											// dwFlags
									pDropFileName,								// lpMultiByteStr
									-1,											// cbMultiByte
									strCmdLine.m_pBuf + strCmdLine.m_iLength,	// lpWideCharStr
									strCmdLine.m_iSize - strCmdLine.m_iLength	// cchWideChar
								);
								
								strCmdLine += L'\"';
								
								pDropFileName += strlen( pDropFileName ) + 1;
								strCmdLine.m_iLength += iFileLen - 1;
							}
						}
						
					  Default:
						strCmdLine += *pCmd;
					}
				}else{
					strCmdLine += *pCmd;
				}
			}
			
			DebugMsgD( strCmdLine.m_pBuf );
			
			if( m_pszFileUserClickedOn ){
				delete [] m_pszFileUserClickedOn;
				m_pszFileUserClickedOn = nullptr;
			}
			
			/*** create process *********************************************/
			
			GetStartupInfo( &si );
			CreateProcess(
				nullptr, strCmdLine.m_pBuf,		/* exec file, params		*/
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
