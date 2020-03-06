/*****************************************************************************

		DropTarget.h --- shell extension drop handler class definition

*****************************************************************************/

// {26250700-66AA-11d2-A84A-9B0573E67737}
DEFINE_GUID( CLSID_DropTarget, 
	0x26250700, 0x66aa, 0x11d2, 0xa8, 0x4a, 0x9b, 0x5, 0x73, 0xe6, 0x77, 0x37);

/*** CShellExtClassFactory class ********************************************/

class CShellExtClassFactory : public IClassFactory{
protected:
	ULONG	m_cRef;
	
public:
	CShellExtClassFactory();
	~CShellExtClassFactory();
	
	// IUnknown members
	STDMETHODIMP			QueryInterface( REFIID, LPVOID FAR * );
	STDMETHODIMP_( ULONG )	AddRef();
	STDMETHODIMP_( ULONG )	Release();
	
	// IClassFactory members
	STDMETHODIMP		CreateInstance( LPUNKNOWN, REFIID, LPVOID FAR * );
	STDMETHODIMP		LockServer( BOOL );
};

typedef CShellExtClassFactory *LPCSHELLEXTCLASSFACTORY;

/*** CShellExt class ********************************************************/

class CShellExt : public IPersistFile, IDropTarget{
protected:
	ULONG			m_cRef;
	wchar_t			m_szFileUserClickedOn[ MAX_PATH ];
	
public:
	CShellExt();
	~CShellExt();
	
	// IUnknown
	STDMETHODIMP			QueryInterface( REFIID, LPVOID FAR * );
	STDMETHODIMP_( ULONG )	AddRef();
	STDMETHODIMP_( ULONG )	Release();
	
	// IPersistFile
	STDMETHODIMP GetClassID( LPCLSID lpClassID );
	STDMETHODIMP IsDirty();
	STDMETHODIMP Load( LPCOLESTR lpszFileName, DWORD grfMode );
	STDMETHODIMP Save( LPCOLESTR lpszFileName, BOOL fRemember );
	STDMETHODIMP SaveCompleted( LPCOLESTR lpszFileName );
	STDMETHODIMP GetCurFile( LPOLESTR FAR* lplpszFileName );
	
	//IDropTarget
	STDMETHODIMP			DragEnter(
								IDataObject	*pDataObject,
								DWORD		grfKeyState,
								POINTL		pt,
								DWORD		*pdwEffect );
	
	STDMETHODIMP			DragOver(
								DWORD		grfKeyState,
								POINTL		pt,
								DWORD		*pdwEffect );
	
	STDMETHODIMP			DragLeave( void );
	
	STDMETHODIMP			Drop(
								IDataObject	*pDataObject,
								DWORD		grfKeyState,
								POINTL		pt,
								DWORD		*pdwEffect );
};

typedef CShellExt *LPCSHELLEXT;
