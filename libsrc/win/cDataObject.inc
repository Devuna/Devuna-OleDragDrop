
!ABCIncludeFile(COM)

! ================================================================================ 
!                        Devuna OLE Drag and Drop Classes
! ================================================================================ 
! Notice : Copyright (C) 2017, Devuna 
!          Distributed under LGPLv3 (http://www.gnu.org/licenses/lgpl.html) 
! 
!    This file is part of Devuna-OleDragDrop (https://github.com/Devuna/Devuna-OleDragDrop) 
! 
!    Devuna-OleDragDrop is free software: you can redistribute it and/or modify 
!    it under the terms of the GNU General Public License as published by 
!    the Free Software Foundation, either version 3 of the License, or 
!    (at your option) any later version. 
! 
!    Devuna-OleDragDrop is distributed in the hope that it will be useful, 
!    but WITHOUT ANY WARRANTY; without even the implied warranty of 
!    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the 
!    GNU General Public License for more details. 
!
!    You should have received a copy of the GNU General Public License 
!    along with Devuna-Devuna-OleDragDrop.  If not, see <http://www.gnu.org/licenses/>. 
! ================================================================================
!
!Author:        Randy Rogers <rrogers@devuna.com>
!Creation Date: 2006.04.07
!
!Revisions:
!==========
!2017.02.02 Devuna rebuild for Clarion 10.0.12463
!2015.03.06 Devuna rebuild for Clarion 10
!2006.05.04 KCR corrected bugs in IDropTargetClass
!           KCR provide feedback based on keystate modifiers
!2006.04.17 KCR fixed problems in IDropTargetClass
!2006.04.07 KCR initial version
!================================================================

    INCLUDE('svcomdef.inc'),ONCE
    INCLUDE('svcom.inc'),ONCE

LPWSTR                          EQUATE(LONG)
DWORD                           EQUATE(LONG)
LPOLESTR                        EQUATE(LONG)
HDROP                           EQUATE(HANDLE)
UINT                            EQUATE(UNSIGNED)


!==========================================================================
! Clipboard format which may be supported by IDataObject from system
! defined shell folders (such as directories, network, ...).
!==========================================================================
CFSTR_SHELLIDLIST                   CSTRING('Shell IDList Array')      ! CF_IDLIST
CFSTR_SHELLIDLISTOFFSET             CSTRING('Shell Object Offsets')    ! CF_OBJECTPOSITIONS
CFSTR_NETRESOURCES                  CSTRING('Net Resource')            ! CF_NETRESOURCE
CFSTR_FILEDESCRIPTORA               CSTRING('FileGroupDescriptor')     ! CF_FILEGROUPDESCRIPTORA
CFSTR_FILEDESCRIPTORW               CSTRING('FileGroupDescriptorW')    ! CF_FILEGROUPDESCRIPTORW
CFSTR_FILECONTENTS                  CSTRING('FileContents')            ! CF_FILECONTENTS
CFSTR_FILENAMEA                     CSTRING('FileName')                ! CF_FILENAMEA
CFSTR_FILENAMEW                     CSTRING('FileNameW')               ! CF_FILENAMEW
CFSTR_PRINTERGROUP                  CSTRING('PrinterFriendlyName')     ! CF_PRINTERS
CFSTR_FILENAMEMAPA                  CSTRING('FileNameMap')             ! CF_FILENAMEMAPA
CFSTR_FILENAMEMAPW                  CSTRING('FileNameMapW')            ! CF_FILENAMEMAPW
CFSTR_SHELLURL                      CSTRING('UniformResourceLocator')
CFSTR_INETURLA                      CSTRING('UniformResourceLocator')
CFSTR_INETURLW                      CSTRING('UniformResourceLocatorW')
CFSTR_PREFERREDDROPEFFECT           CSTRING('Preferred DropEffect')
CFSTR_PERFORMEDDROPEFFECT           CSTRING('Performed DropEffect')
CFSTR_PASTESUCCEEDED                CSTRING('Paste Succeeded')
CFSTR_INDRAGLOOP                    CSTRING('InShellDragLoop')
CFSTR_DRAGCONTEXT                   CSTRING('DragContext')
CFSTR_MOUNTEDVOLUME                 CSTRING('MountedVolume')
CFSTR_PERSISTEDDATAOBJECT           CSTRING('PersistedDataObject')
CFSTR_TARGETCLSID                   CSTRING('TargetCLSID')                       ! HGLOBAL with a CLSID of the drop target
CFSTR_LOGICALPERFORMEDDROPEFFECT    CSTRING('Logical Performed DropEffect')
CFSTR_AUTOPLAY_SHELLIDLISTS         CSTRING('Autoplay Enumerated IDList Array')  !  (HGLOBAL with LPIDA)
CFSTR_UNTRUSTEDDRAGDROP             CSTRING('UntrustedDragDrop')                 !  DWORD

  COMPILE('EndUnicode', _Unicode_)
CFSTR_FILEDESCRIPTOR    CSTRING('FileGroupDescriptorW')    ! CF_FILEGROUPDESCRIPTORW
CFSTR_FILENAME          CSTRING('FileNameW')               ! CF_FILENAMEW
CFSTR_FILENAMEMAP       CSTRING('FileNameMapW')            ! CF_FILENAMEMAPW
CFSTR_INETURL           CSTRING('UniformResourceLocatorW')
  !EndUnicode
  OMIT('EndAnsi', _Unicode_)
CFSTR_FILEDESCRIPTOR    CSTRING('FileGroupDescriptor')     ! CF_FILEGROUPDESCRIPTORA
CFSTR_FILENAME          CSTRING('FileName')                ! CF_FILENAMEA
CFSTR_FILENAMEMAP       CSTRING('FileNameMap')             ! CF_FILENAMEMAPA
CFSTR_INETURL           CSTRING('UniformResourceLocator')
  !EndAnsi

DVASPECT                        ITEMIZE
DVASPECT_CONTENT                    EQUATE
DVASPECT_THUMBNAIL                  EQUATE
DVASPECT_ICON                       EQUATE(4)
DVASPECT_DOCPRINT                   EQUATE(8)
                                END

DVTARGETDEVICE                  GROUP,TYPE
tdSize                              LONG
tdDriverNameOffset                  SHORT
tdDeviceNameOffset                  SHORT
tdPortNameOffset                    SHORT
tdExtDevmodeOffset                  SHORT
tdData                              BYTE
                                END

DV_E_FORMATETC                  EQUATE(080040064h)
DATA_E_FORMATETC                EQUATE(DV_E_FORMATETC)
OLE_E_ADVISENOTSUPPORTED        EQUATE(080040003h)

DRAGDROP_E_NOTREGISTERED        EQUATE(080040100h)
DRAGDROP_E_ALREADYREGISTERED    EQUATE(080040101h)
DRAGDROP_E_INVALIDHWND          EQUATE(080040102h)

DATADIR_GET                     EQUATE(1)
DATADIR_SET                     EQUATE(2)

DRAGDROP_S_DROP                 EQUATE(000040100h)
DRAGDROP_S_CANCEL               EQUATE(000040101h)
DRAGDROP_S_USEDEFAULTCURSORS    EQUATE(000040102h)
DATA_S_SAMEFORMATETC            EQUATE(000040130h)

DROPEFFECT_NONE                 EQUATE(0)
DROPEFFECT_COPY                 EQUATE(1)
DROPEFFECT_MOVE                 EQUATE(2)
DROPEFFECT_LINK                 EQUATE(4)
DROPEFFECT_SCROLL               EQUATE(080000000h)

DROPFILES                       GROUP,TYPE
pFiles                              DWORD       ! offset of file list
pt                                  LIKE(POINT) ! drop point (client coords)
fNC                                 BOOL        ! is it on NonClient area and pt is in screen coords
fWide                               BOOL        ! wide character flag
                                END

FORMATETC                       GROUP,TYPE
cfFormat                            LONG
ptd                                 LONG              
dwAspect                            ULONG
lindex                              LONG
tymed                               ULONG
                                END
tagFORMATETC                    GROUP(FORMATETC),TYPE   !added for backward compatibility
                                END

STATSTG                         GROUP,TYPE
pwcsName                            LPOLESTR
type                                DWORD
cbSize                              LIKE(_ULARGE_INTEGER)
mtime                               LIKE(_FILETIME)
ctime                               LIKE(_FILETIME)
atime                               LIKE(_FILETIME)
grfMode                             DWORD
grfLocksSupported                   DWORD
clsid                               LIKE(_CLSID)
grfStateBits                        DWORD
reserved                            DWORD
                                END
_tagSTATSTG_                    EQUATE(1)
tagSTATSTG                      GROUP(STATSTG),TYPE
                                END


FILEDESCRIPTOR                  GROUP,TYPE
dwFlags                             DWORD
clsid                               LIKE(_CLSID)
sizel                               LIKE(SIZE)
pointl                              LIKE(POINT)
dwFileAttributes                    DWORD
ftCreationTime                      LIKE(_FILETIME)
ftLastAccessTime                    LIKE(_FILETIME)
ftLastWriteTime                     LIKE(_FILETIME)
nFileSizeHigh                       DWORD
nFileSizeLow                        DWORD
cFileName                           CSTRING(MAX_PATH)
                                END

FILEGROUPDESCRIPTOR             GROUP,TYPE
cItems                              UNSIGNED
fd                                  LIKE(FILEDESCRIPTOR)
                                END

STGMEDIUM                       GROUP,TYPE
tymed                             ULONG
u                                 GROUP
hMetaFilePict                       LONG
hHEnhMetaFile                       LONG,OVER(hMetaFilePict)
hGdiHandle                          LONG,OVER(hMetaFilePict)
hGlobal                             LONG,OVER(hMetaFilePict)
lpszFileName                        LPWSTR,OVER(hMetaFilePict)
pstm                                LONG,OVER(hMetaFilePict)
pstg                                LONG,OVER(hMetaFilePict)
                                  END
pUnkForRelease                    LONG
                                END

kcr_IStorage                    INTERFACE(IUnknown),COM
CreateStream                      PROCEDURE(LONG pwcsName, LONG grfMode, LONG reserved1, LONG reserved2, *LONG pstm),HRESULT
OpenStream                        PROCEDURE(LONG pwcsName, LONG preserved1, LONG grfMode, LONG reserved2, *LONG pstm),HRESULT
CreateStorage                     PROCEDURE(LONG pwcsName, LONG grfMode, LONG reserved1, LONG reserved2, *LONG pstg),HRESULT
OpenStorage                       PROCEDURE(LONG pwcsName, *kcr_IStorage stgPriority, LONG grfMode, LONG snbExclude, LONG reserved, *LONG pstg),HRESULT,RAW
CopyTo                            PROCEDURE(LONG ciidExclude, LONG rgiidExclude, LONG snbExclude, *kcr_IStorage stgDest),HRESULT,RAW
MoveElementTo                     PROCEDURE(LONG pwcsName, *kcr_IStorage stgDest, LONG pwcsNewName, LONG grfFlags),HRESULT,RAW
Commit                            PROCEDURE(LONG grfCommitFlags),HRESULT
Revert                            PROCEDURE,HRESULT
EnumElements                      PROCEDURE(LONG reserved1, LONG reserved2, LONG reserved3, *LONG penum),HRESULT
DestroyElement                    PROCEDURE(LONG pwcsName),HRESULT
RenameElement                     PROCEDURE(LONG pwcsOldName, LONG pwcsNewName),HRESULT
SetElementTimes                   PROCEDURE(LONG pwcsName, LONG pctime, LONG patime, LONG pmtime),HRESULT
SetClass                          PROCEDURE(LONG rclsid),HRESULT
SetStateBits                      PROCEDURE(LONG grfStateBits, LONG grfMask),HRESULT
Stat                              PROCEDURE(LONG pstatstg, LONG grfStatFlag),HRESULT
                                END


IEnumSTATSTG                    INTERFACE(IUnknown),COM,TYPE
Next                              PROCEDURE(ULONG celt, *STATSTG rgelt, *ULONG pceltFetched),HRESULT,RAW,PROC
Skip                              PROCEDURE(ULONG celt),HRESULT,RAW,PROC
Reset                             PROCEDURE(),HRESULT,RAW,PROC
Clone                             PROCEDURE(*LONG ppenum),HRESULT,RAW,PROC
                                END

IID_IAdviseSink                 GROUP
data1                             LONG(10FH)
data2                             SHORT(0H)
data3                             SHORT(0H)
data4                             STRING('<0C0H><0H><0H><0H><0H><0H><0H><46H>')
                                END

IAdviseSink                     INTERFACE(IUnknown),COM,TYPE
OnDataChange                      PROCEDURE(*FORMATETC pformatetc,*STGMEDIUM pStgmed),HRESULT,RAW,PROC
OnViewChange                      PROCEDURE(ULONG dwAspect,LONG lindex),HRESULT,RAW,PROC
OnRename                          PROCEDURE(*IMoniker pmk),HRESULT,RAW,PROC
OnSave                            PROCEDURE(),HRESULT,RAW,PROC
OnClose                           PROCEDURE(),HRESULT,RAW,PROC
                                END

IID_IDropSource                 GROUP
data1                             LONG(121H)
data2                             SHORT(0H)
data3                             SHORT(0H)
data4                             STRING('<0C0H><0H><0H><0H><0H><0H><0H><46H>')
                                END

IDropSource                     INTERFACE(IUnknown),COM,TYPE
QueryContinueDrag                 PROCEDURE(BOOL fEscapePressed, DWORD grfKeyState),HRESULT,PROC
GiveFeedback                      PROCEDURE(DWORD dwEffect),HRESULT,PROC
                                END

IID_IDropTarget                 GROUP
data1                             LONG(122H)
data2                             SHORT(0H)
data3                             SHORT(0H)
data4                             STRING('<0C0H><0H><0H><0H><0H><0H><0H><46H>')
                                END

kcr_IDropTarget                 INTERFACE(IUnknown),COM
DragEnter                         PROCEDURE(LONG pDataObj, LONG grfKeyState, LONG xPos, LONG yPos, *LONG pdwEffect),HRESULT
DragOver                          PROCEDURE(LONG grfKeyState, LONG xPOs, LONG yPos, *LONG pdwEffect),HRESULT
DragLeave                         PROCEDURE,HRESULT
Drop                              PROCEDURE(LONG pDataObj, LONG grfKeyState, LONG xPos, LONG yPos, *LONG pdwEffect),HRESULT
                                END


IID_IDataObject                 GROUP
data1                             LONG(10EH)
data2                             SHORT(0H)
data3                             SHORT(0H)
data4                             STRING('<0C0H><0H><0H><0H><0H><0H><0H><46H>')
                                END

IDataObject                     INTERFACE(IUnknown),COM,TYPE
GetData                           PROCEDURE(LONG pformatetcIn,LONG pRemoteMedium),HRESULT,PROC
GetDataHere                       PROCEDURE(LONG pformatetc,LONG pRemoteMedium),HRESULT,PROC
QueryGetData                      PROCEDURE(LONG pformatetc),HRESULT,PROC
GetCanonicalFormatEtc             PROCEDURE(LONG pformatectIn,LONG pformatetcOut),HRESULT,PROC
SetData                           PROCEDURE(LONG pformatetc,LONG pmedium,LONG fRelease),HRESULT,PROC
EnumFormatEtc                     PROCEDURE(ULONG dwDirection,LONG ppenumFormatEtc),HRESULT,PROC
DAdvise                           PROCEDURE(LONG pformatetc,ULONG advf,LONG pAdvSink,LONG pdwConnection),HRESULT,PROC
DUnadvise                         PROCEDURE(ULONG dwConnection),HRESULT,PROC
EnumDAdvise                       PROCEDURE(LONG ppenumAdvise),HRESULT,PROC
                                END

IID_IEnumFORMATETC              GROUP
data1                             LONG(103H)
data2                             SHORT(0H)
data3                             SHORT(0H)
data4                             STRING('<0C0H><0H><0H><0H><0H><0H><0H><46H>')
                                END

IEnumFORMATETC                  INTERFACE(IUnknown),COM,TYPE
RemoteNext                        PROCEDURE(ULONG celt,LONG rgelt,LONG pceltFetched),HRESULT,PROC
Skip                              PROCEDURE(ULONG celt),HRESULT,PROC
Reset                             PROCEDURE(),HRESULT,PROC
Clone                             PROCEDURE(LONG ppenum),HRESULT,PROC
                                END


IDropSourceClass                CLASS(CCOMUSerObject),IMPLEMENTS(IDropSource),TYPE,MODULE('cDataObject.clw'),LINK('cDataObject.clw',_COMLinkMode_),DLL(_COMDllMode_)
IsInitialized                     BOOL,PROTECTED
debug                             BOOL
IDropSourceObj                    &IDropSource,PROTECTED
Construct                         PROCEDURE()
Destruct                          PROCEDURE(),VIRTUAL
Init                              PROCEDURE(),HRESULT,PROC,VIRTUAL
Kill                              PROCEDURE(),HRESULT,PROC,VIRTUAL
GetInterfaceObject                PROCEDURE(),LONG  !*IDropSource
GetLibLocation                    PROCEDURE(),STRING,VIRTUAL
_ShowErrorMessage                 PROCEDURE(STRING pMethodName,HRESULT pHR,BYTE pUseDebugMode=true),HRESULT,PROC,VIRTUAL
!Interface methods
QueryInterface                    PROCEDURE(LONG riid, *LONG pvObject),HRESULT,DERIVED
QueryContinueDrag                 PROCEDURE(BOOL fEscapePressed, DWORD grfKeyState),HRESULT,PROC
GiveFeedback                      PROCEDURE(DWORD dwEffect),HRESULT,PROC
                                END

IDropTargetClass                CLASS(CCOMUSerObject),IMPLEMENTS(kcr_IDropTarget),TYPE,MODULE('cDataObject.clw'),LINK('cDataObject.clw',_COMLinkMode_),DLL(_COMDllMode_)
m_hWnd                            HWND,PROTECTED
m_OleDropEvent                    LONG,PROTECTED    !event to fire on drop
m_fAllowDrop                      BOOL
IsInitialized                     BOOL,PROTECTED
debug                             BOOL
IDropTargetObj                    &IDropTarget,PROTECTED
Construct                         PROCEDURE()
Destruct                          PROCEDURE(),VIRTUAL
Init                              PROCEDURE(HWND hWnd, LONG OleDropEvent),HRESULT,PROC,VIRTUAL
Kill                              PROCEDURE(),HRESULT,PROC,VIRTUAL
GetInterfaceObject                PROCEDURE(),LONG  !*IDropTarget
GetLibLocation                    PROCEDURE(),STRING,VIRTUAL
Register                          PROCEDURE(),HRESULT,PROC,PROTECTED
DropEffect                        PROCEDURE(DWORD grfKeyState, *POINT pt, DWORD dwAllowed),DWORD,VIRTUAL
QueryDataObject                   PROCEDURE(*IDataObject pDataObject),BOOL,VIRTUAL
_ShowErrorMessage                 PROCEDURE(STRING pMethodName,HRESULT pHR,BYTE pUseDebugMode=true),HRESULT,PROC,VIRTUAL
!Interface methods
QueryInterface                    PROCEDURE(LONG riid, *LONG pvObject),HRESULT,DERIVED
DragEnter                         PROCEDURE(*IDataObject pDataObject, DWORD grfKeyState, POINT pt, *DWORD pdwEffect),DWORD,PROC,VIRTUAL
DragOver                          PROCEDURE(DWORD grfKeyState, POINT pt, *DWORD pdwEffect),HRESULT,PROC,VIRTUAL
DragLeave                         PROCEDURE(),HRESULT,PROC,VIRTUAL
Drop                              PROCEDURE(*IDataObject pDataObject, DWORD grfKeyState, POINT pt, *DWORD pdwEffect),HRESULT,PROC,VIRTUAL
                                END

MAX_NUM_FORMAT                  EQUATE(10)   !Maximum number of formats supported by IDataObjectClass
IDataObjectClass                CLASS,IMPLEMENTS(IDataObject),TYPE,MODULE('cDataObject.clw'),LINK('cDataObject.clw',_COMLinkMode_),DLL(_COMDllMode_)
m_cRef                            LONG,PRIVATE
m_nNumFormats                     LONG
m_nMaxFormats                     LONG
m_pFormatEtc                      LONG	!LIKE(FORMATETC),DIM(MAX_NUM_FORMAT)
m_pStgMedium                      LONG	!LIKE(STGMEDIUM),DIM(MAX_NUM_FORMAT)
IsInitialized                     BOOL,PROTECTED
debug                             BOOL
IDataObjectObj                    &IDataObject,PROTECTED

Construct                         PROCEDURE()
Destruct                          PROCEDURE(),VIRTUAL
Init                              PROCEDURE(),HRESULT,PROC,VIRTUAL
Init                              PROCEDURE(*tagFORMATETC fmtetc, *STGMEDIUM stgmed, SIGNED count),HRESULT,PROC,VIRTUAL
Kill                              PROCEDURE(),HRESULT,PROC,VIRTUAL
GetInterfaceObject                PROCEDURE(),LONG                  !*IDataObject
GetLibLocation                    PROCEDURE(),STRING
LookupFormatEtc                   PROCEDURE(*FORMATETC pFormatEtc),SIGNED
_ShowErrorMessage                 PROCEDURE(STRING pMethodName,HRESULT pHR,BYTE pUseDebugMode=true),HRESULT,PROC,VIRTUAL
IsEqualIID                        PROCEDURE(REFIID riid1, REFIID riid2),BOOL
!IUnknown Interface derived methods
QueryInterface                    PROCEDURE(LONG riid, *LONG pvObject),HRESULT,VIRTUAL
AddRef                            PROCEDURE(),LONG,PROC,VIRTUAL
Release                           PROCEDURE(),LONG,PROC,VIRTUAL
!IDataObject Interface methods
GetData                           PROCEDURE(*FORMATETC pformatetcIn,*STGMEDIUM pRemoteMedium),HRESULT,PROC,VIRTUAL
GetDataHere                       PROCEDURE(*FORMATETC pformatetc,*STGMEDIUM pRemoteMedium),HRESULT,PROC,VIRTUAL
QueryGetData                      PROCEDURE(*FORMATETC pformatetc),HRESULT,PROC,VIRTUAL
GetCanonicalFormatEtc             PROCEDURE(*FORMATETC pformatectIn,*FORMATETC pformatetcOut),HRESULT,PROC,VIRTUAL
SetData                           PROCEDURE(*FORMATETC pformatetc,*STGMEDIUM pmedium,LONG fRelease=1),HRESULT,PROC,VIRTUAL
EnumFormatEtc                     PROCEDURE(ULONG dwDirection,LONG ppenumFormatEtc),HRESULT,PROC,VIRTUAL
DAdvise                           PROCEDURE(*FORMATETC pformatetc,ULONG advf,*IAdviseSink pAdvSink,*ULONG pdwConnection),HRESULT,PROC,VIRTUAL
DUnadvise                         PROCEDURE(ULONG dwConnection),HRESULT,PROC,VIRTUAL
EnumDAdvise                       PROCEDURE(LONG ppenumAdvise),HRESULT,PROC,VIRTUAL
                                END


IEnumFORMATETCClass             CLASS(CCOMUSerObject),IMPLEMENTS(IEnumFORMATETC),TYPE,MODULE('cDataObject.clw'),LINK('cDataObject.clw',_COMLinkMode_),DLL(_COMDllMode_)
m_nNumFormats                     LONG
m_pFormatEtc                      LONG  !*FORMATETC[]
m_nIndex                          LONG
IsInitialized                     BOOL,PROTECTED
debug                             BOOL
IEnumFORMATETCObj                 &IEnumFORMATETC,PROTECTED
Construct                         PROCEDURE()
Destruct                          PROCEDURE(),VIRTUAL
Init                              PROCEDURE(*FORMATETC paFormatEtc, SIGNED count),HRESULT,PROC,VIRTUAL
Kill                              PROCEDURE(),HRESULT,PROC,VIRTUAL
GetInterfaceObject                PROCEDURE(),LONG  !*IEnumFORMATETC
GetLibLocation                    PROCEDURE(),STRING,VIRTUAL
_ShowErrorMessage                 PROCEDURE(STRING pMethodName,HRESULT pHR,BYTE pUseDebugMode=true),HRESULT,PROC,VIRTUAL
!Interface methods
QueryInterface                    PROCEDURE(LONG riid, *LONG pvObject),HRESULT,DERIVED
RemoteNext                        PROCEDURE(ULONG celt,*FORMATETC rgelt,LONG pceltFetched),HRESULT,PROC,VIRTUAL
Skip                              PROCEDURE(ULONG celt),HRESULT,PROC,VIRTUAL
Reset                             PROCEDURE(),HRESULT,PROC,VIRTUAL
Clone                             PROCEDURE(LONG ppenum),HRESULT,PROC,VIRTUAL
                                END

OleDataGroup                    GROUP,TYPE
UseVariable                         LONG        !drop target control use variable
UseFieldEquate                      LONG        !drop target control field equate value
                                END

OleDataSourceClass              CLASS,TYPE,MODULE('cDataObject.clw'),LINK('cDataObject.clw',_COMLinkMode_),DLL(_COMDllMode_)
m_SelStart                        LONG
m_SelEnd                          LONG
GetOleData                        PROCEDURE(),LONG,PROC,VIRTUAL
OnLButtonDown                     PROCEDURE(),LONG,PROC,VIRTUAL
                                END

TextDataObjectClass             CLASS(IDataObjectClass),TYPE,MODULE('cDataObject.clw'),LINK('cDataObject.clw',_COMLinkMode_),DLL(_COMDllMode_)
Init                              PROCEDURE(*CSTRING szText),HRESULT,PROC,VIRTUAL
                                END

cOleIniter                      CLASS,TYPE,MODULE('cDataObject.clw'),LINK('cDataObject.clw',_COMLinkMode_),DLL(_COMDllMode_)
Construct                         PROCEDURE()
Destruct                          PROCEDURE,VIRTUAL
                                END