   MEMBER

! ================================================================================
!                        Devuna OLE Drag and Drop Classes
! ================================================================================
! Notice : Copyright (C) 2017, Devuna
!          Distributed under the MIT License (https://opensource.org/licenses/MIT)
!
!    This file is part of Devuna-OleDragDrop (https://github.com/Devuna/Devuna-OleDragDrop)
!
!    Devuna-OleDragDrop is free software: you can redistribute it and/or modify
!    it under the terms of the MIT License as published by
!    the Open Source Initiative.
!
!    Devuna-OleDragDrop is distributed in the hope that it will be useful,
!    but WITHOUT ANY WARRANTY; without even the implied warranty of
!    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
!    MIT License for more details.
!
!    You should have received a copy of the MIT License
!    along with Devuna-Devuna-OleDragDrop.  If not, see <https://opensource.org/licenses/MIT>.
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
  INCLUDE('cDataObject.inc'),ONCE

  MAP
    MODULE('WinAPI')
      kcr_GlobalHandle(LONG lpvMem),HGLOBAL,NAME('GlobalHandle'),PASCAL,NAME('GlobalHandle')
      kcr_GlobalSize(HGLOBAL hMem),DWORD,NAME('GlobalSize'),PASCAL,NAME('GlobalSize')
      kcr_memcpy(LONG lpDest,LONG lpSource,LONG nCount),LONG,PROC,NAME('_memcpy')
      kcr_OleInitialize(LONG pvReserved),LONG,PASCAL,PROC,NAME('OleInitialize')
      kcr_OleUninitialize(),PASCAL,NAME('OleUninitialize')
      kcr_SHCreateStdEnumFmtEtc(LONG nNumFormats, LONG paFormatEtc, LONG ppenumFormatEtc),HRESULT,PROC,PASCAL,NAME('SHCreateStdEnumFmtEtc')
      kcr_ZeroMemory(LONG DestinationPtr,LONG dwLength),RAW,PASCAL,NAME('RtlZeroMemory')
    END

    INCLUDE('svapifnc.inc'),ONCE

    CreateEnumFormatEtc(LONG nNumFormats, LONG paFormatEtc, LONG ppenumFormatEtc),HRESULT,PROC
    Dec2Hex(ULONG),STRING
    DeepCopyFormatEtc(LONG dest, LONG source)
    DupMem(HGLOBAL hMem),HGLOBAL
  END

CreateEnumFormatEtc PROCEDURE(LONG nNumFormats, LONG paFormatEtc, LONG ppIEnumFORMATETC)
hr                              HRESULT(S_OK)
cEnumFormatEtc                  &IEnumFORMATETCClass
fmtetc                          &FORMATETC
pIEnumFORMATETC                 LONG

  CODE
    IF nNumFormats           = 0 OR |
       paFormatEtc           = 0 OR |
       ppIEnumFORMATETC      = 0
       hr = E_INVALIDARG
    ELSE
       cEnumFormatEtc &= NEW IEnumFORMATETCClass
       IF ~cEnumFormatEtc &= NULL
          fmtetc &= (paFormatEtc)
          hr = cEnumFormatEtc.Init(fmtetc, nNumFormats)
          IF hr = S_OK
             pIEnumFORMATETC = cEnumFormatEtc.GetInterfaceObject()
             kcr_memcpy(ppIEnumFORMATETC,ADDRESS(pIEnumFORMATETC),4)
          END
       ELSE
          hr = E_OUTOFMEMORY
       END
    END
    RETURN hr


Dec2Hex                         PROCEDURE(ULONG pDec)
locHex                          STRING(30)
  CODE
  LOOP UNTIL(~pDec)
    locHex = SUB('0123456789ABCDEF',1+pDec % 16,1) & CLIP(locHex)
    pDec = INT(pDec / 16)
  END
  RETURN CLIP(locHex)


DeepCopyFormatEtc   PROCEDURE(LONG dest, LONG source)
destFormatEtc       &FORMATETC
sourceFormatEtc     &FORMATETC

  CODE
    !copy the source FORMATETC into dest
    kcr_memcpy(dest, source, SIZE(FORMATETC))
    destFormatEtc   &= (dest)
    sourceFormatEtc &= (source)
    IF sourceFormatEtc.ptd <> 0
       destFormatEtc = CoTaskMemAlloc(SIZE(DVTARGETDEVICE))
       kcr_memcpy(destFormatEtc.ptd,sourceFormatEtc.ptd,SIZE(DVTARGETDEVICE))
    END
    RETURN


DupMem  PROCEDURE(HGLOBAL hMem) !,HGLOBAL
source      LONG
dest        LONG
dwBytes     DWORD

  CODE
    !lock the source memory object
    source = GlobalLock(hMem)

    !create a fixed "global" block - i.e. just
    !a regular lump of our process heap
    dwBytes = kcr_GlobalSize(hMem)
    dest    = GlobalAlloc(GMEM_FIXED, dwBytes)
    kcr_memcpy(dest, source, dwBytes)

    GlobalUnlock(hMem)

    RETURN kcr_GlobalHandle(dest)


!========================================================!
! IDataObjectClass implementation                        !
!========================================================!
IDataObjectClass.Construct      PROCEDURE()
  CODE
    SELF.debug = TRUE


IDataObjectClass.Destruct       PROCEDURE() !,VIRTUAL
  CODE
    IF SELF.IsInitialized = TRUE
       SELF.Kill()
    END


IDataObjectClass.Init               PROCEDURE() !,VIRTUAL
hr      HRESULT(S_OK)

  CODE
    SELF.m_nNumFormats = 0
    SELF.m_nMaxFormats = MAX_NUM_FORMAT
    SELF.m_pFormatEtc = GlobalAlloc(GPTR, SIZE(tagFORMATETC) * SELF.m_nMaxFormats )
    SELF.m_pStgMedium = GlobalAlloc(GPTR, SIZE(STGMEDIUM) * SELF.m_nMaxFormats )
    kcr_ZeroMemory(SELF.m_pFormatEtc,SIZE(FORMATETC) * SELF.m_nMaxFormats)
    kcr_ZeroMemory(SELF.m_pStgMedium,SIZE(STGMEDIUM) * SELF.m_nMaxFormats)
    SELF.IDataObjectObj &= ADDRESS(SELF.IDataObject)
    SELF.IsInitialized = TRUE
    RETURN hr


IDataObjectClass.Init           PROCEDURE(*tagFORMATETC fmtetc, *STGMEDIUM stgmed, SIGNED count)   !,HRESULT,PROC,VIRTUAL
hr      HRESULT(S_OK)

  CODE
    SELF.m_nNumFormats = count
    SELF.m_nMaxFormats = count
    SELF.m_pFormatEtc = GlobalAlloc(GPTR, SIZE(tagFORMATETC) * count )
    SELF.m_pStgMedium = GlobalAlloc(GPTR, SIZE(STGMEDIUM) * count )
    kcr_memcpy(SELF.m_pFormatEtc,ADDRESS(fmtetc), SIZE(tagFORMATETC) * count)
    kcr_memcpy(SELF.m_pStgMedium,ADDRESS(stgmed), SIZE(STGMEDIUM) * count)
    SELF.IDataObjectObj &= ADDRESS(SELF.IDataObject)
    SELF.IsInitialized = TRUE
    RETURN hr



IDataObjectClass.Kill               PROCEDURE() !,VIRTUAL
hr      HRESULT

  CODE
    IF SELF.IsInitialized = FALSE
       hr = E_NOINTERFACE
    ELSE
       IF SELF.m_pFormatEtc
          GlobalFree(SELF.m_pFormatEtc)
       END
       IF SELF.m_pStgMedium
          GlobalFree(SELF.m_pStgMedium)
       END
       IF NOT (SELF.IDataObjectObj &= NULL)
          SELF.IDataObjectObj &= NULL
       END
       SELF.IsInitialized = FALSE
       hr = S_OK
    END
    RETURN hr


IDataObjectClass.GetInterfaceObject PROCEDURE()
  CODE
    RETURN ADDRESS(SELF.IDataObjectObj)


IDataObjectClass.GetLibLocation     PROCEDURE()
  CODE
    RETURN ''


IDataObjectClass.LookupFormatEtc    PROCEDURE(*FORMATETC pFormatEtc) !,SIGNED
pElement    LONG
_FormatEtc  &tagFORMATETC
I           LONG
szcfFormat  CSTRING(64)
AllMedia    ULONG(0FFFFFFFFh)

  CODE
    LOOP I = 1 TO SELF.m_nNumFormats
       pElement = SELF.m_pFormatEtc + ((I-1) * SIZE(tagFORMATETC))
       _FormatEtc &= (pElement)
       IF ((_FormatEtc.tymed = pFormatEtc.tymed) OR (pFormatEtc.tymed = AllMedia))  AND    |
          (_FormatEtc.cfFormat = pFormatEtc.cfFormat) AND    |
          (_FormatEtc.dwaspect = pFormatEtc.dwaspect)
          BREAK
       END
    END

    RETURN CHOOSE(I <= SELF.m_nNumFormats,I,-1)


IDataObjectClass._ShowErrorMessage  PROCEDURE(STRING pMethodName,HRESULT pHR,BYTE pUseDebugMode=1)  !,VIRTUAL
hr                              HRESULT(S_OK)
  CODE
    IF NOT(SELF.debug=FALSE AND pUseDebugMode=TRUE)
       MESSAGE(pMethodName & ' failed (' & Dec2Hex(pHR) & ')','COM Error',ICON:EXCLAMATION)
    END
    RETURN S_OK


IDataObjectClass.IsEqualIID PROCEDURE(REFIID riid1, REFIID riid2)

Guid1       LIKE(_GUIDL)
Guid2       LIKE(_GUIDL)

  CODE
    kcr_memcpy(ADDRESS(Guid1), riid1, SIZE(_GUID))
    kcr_memcpy(ADDRESS(Guid2), riid2, SIZE(_GUID))
    IF (Guid1.Data1 = Guid2.Data1)
       IF (Guid1.Data2 = Guid2.Data2)
          IF (Guid1.Data3 = Guid2.Data3)
             IF (Guid1.Data4 = Guid2.Data4)
                RETURN TRUE
             END
          END
       END
    END
    RETURN FALSE


!--------------------------------------------------------------------------
!IUnknown Interface Implementation
!--------------------------------------------------------------------------

IDataObjectClass.QueryInterface procedure(REFIID riid, *long pvObject)  !,HRESULT,VIRTUAL
hr          HRESULT(E_FAIL)

  CODE
    pvObject = 0
    IF (SELF.IsEqualIID(riid, ADDRESS(_IUnknown)))
       pvObject = ADDRESS(SELF.IDataObject)
    ELSIF (SELF.IsEqualIID(riid, ADDRESS(IID_IDataObject)))
       pvObject = ADDRESS(SELF.IDataObject)
    END
    IF (pvObject = 0)
       hr = E_NOINTERFACE
    ELSE
       SELF.AddRef()
       hr = COM_NOERROR
    END
    RETURN hr

IDataObjectClass.AddRef PROCEDURE   !,LONG,PROC,VIRTUAL

  CODE
    RETURN InterlockedIncrement(SELF.m_cRef)


IDataObjectClass.Release PROCEDURE  !,LONG,PROC,VIRTUAL

  CODE
    IF ~SELF.m_cRef
       RETURN 0
    END
    IF InterlockedDecrement(SELF.m_cRef)
       RETURN SELF.m_cRef
    END
    RETURN 0


!--------------------------------------------------------------------------
!IDataObject Interface Implementation
!--------------------------------------------------------------------------

IDataObjectClass.GetData    PROCEDURE(*FORMATETC pformatetc,*STGMEDIUM pRemoteMedium)  !,HRESULT,PROC,VIRTUAL
hr          HRESULT(S_OK)
idx         LONG
pMedium     &STGMEDIUM

  CODE
    IF SELF.IsInitialized = FALSE
       hr = E_FAIL
    ELSE
       idx = SELF.LookupFormatEtc(pformatetc)
       IF idx = -1
          hr = DV_E_FORMATETC
       ELSE
          pMedium &= (SELF.m_pStgMedium + ((idx-1) * SIZE(STGMEDIUM)))
          pRemoteMedium.tymed = pMedium.tymed
          pRemoteMedium.pUnkForRelease = 0
          CASE pMedium.tymed
            OF TYMED_HGLOBAL
               pRemoteMedium.u.hGlobal = DupMem(pMedium.u.hGlobal)
          ELSE
             hr = DV_E_FORMATETC
          END
       END
    END

    !check for error
    CASE hr
      OF S_OK OROF DV_E_FORMATETC
         !expected results
    ELSE
         SELF._ShowErrorMessage('IDataObjectClass.GetData',HR)
    END

    RETURN hr


IDataObjectClass.GetDataHere    PROCEDURE(*FORMATETC pformatetc,*STGMEDIUM pRemoteMedium)  !,HRESULT,PROC,VIRTUAL
hr          HRESULT(S_OK)

  CODE
    IF SELF.IsInitialized = FALSE
       hr = E_FAIL
    ELSE
       hr = DATA_E_FORMATETC
    END

    !check for error
    CASE hr
      OF S_OK OROF DV_E_FORMATETC
         !expected results
    ELSE
         SELF._ShowErrorMessage('IDataObjectClass.GetDataHere',HR)
    END

    RETURN hr


IDataObjectClass.QueryGetData   PROCEDURE(*FORMATETC pFormatEtc) !,HRESULT,PROC,VIRTUAL
hr          HRESULT(S_OK)

  CODE
    IF SELF.IsInitialized = FALSE
       hr = E_FAIL
    ELSE
       hr = CHOOSE(SELF.LookupFormatEtc(pFormatEtc) = -1, DV_E_FORMATETC, S_OK)
    END

    !check for error
    CASE hr
      OF S_OK OROF DV_E_FORMATETC
         !expected results
    ELSE
         SELF._ShowErrorMessage('IDataObjectClass.QueryGetData',HR)
    END

    RETURN hr



IDataObjectClass.GetCanonicalFormatEtc    PROCEDURE(*FORMATETC pFormatEtcIn,*FORMATETC pFormatEtcOut) !,HRESULT,PROC,VIRTUAL
hr          HRESULT(S_OK)

  CODE
    IF SELF.IsInitialized = FALSE
       hr = E_FAIL
    ELSE
       !Apparently we have to set this field to NULL even though we don't do anything
       pFormatEtcOut.ptd = 0
       hr = E_NOTIMPL
    END

    !check for error
    CASE hr
      OF S_OK OROF E_NOTIMPL
         !expected results
    ELSE
         SELF._ShowErrorMessage('IDataObjectClass.GetCanonicalFormatEtc',HR)
    END

    RETURN hr


IDataObjectClass.SetData  PROCEDURE(*FORMATETC pFormatEtc,*STGMEDIUM pMedium,LONG fRelease=1)  !,HRESULT,PROC,VIRTUAL
hr          HRESULT(S_OK)
i           LONG
_StgMedium  &STGMEDIUM
_FormatEtc  &FORMATETC
lpDest      LONG
lpSource    LONG


  CODE
    IF SELF.IsInitialized = FALSE
       hr = E_FAIL
    ELSE
       IF ~fRelease
          hr = E_FAIL
       ELSE
          LOOP i = 1 TO SELF.m_nNumFormats
             _FormatEtc &= (SELF.m_pFormatEtc + ((i-1) * SIZE(FORMATETC)))
             _StgMedium &= (SELF.m_pStgMedium + ((i-1) * SIZE(STGMEDIUM)))
             IF _FormatEtc.cfFormat = pFormatEtc.cfFormat
                lpDest = SELF.m_pFormatEtc + ((i-1) * SIZE(FORMATETC))
                lpSource = ADDRESS(pformatetc)
                kcr_memcpy(lpDest,lpSource, SIZE(FORMATETC))

                lpDest = SELF.m_pStgMedium + ((i-1) * SIZE(STGMEDIUM))
                lpSource = ADDRESS(pMedium)
                kcr_memcpy(lpDest,lpSource, SIZE(STGMEDIUM))

                hr = S_OK
                BREAK
             END
          END
          IF i > SELF.m_nNumFormats
             IF i <= SELF.m_nMaxFormats
                lpDest = SELF.m_pFormatEtc + ((i-1) * SIZE(FORMATETC))
                lpSource = ADDRESS(pformatetc)
                kcr_memcpy(lpDest,lpSource, SIZE(FORMATETC))

                lpDest = SELF.m_pStgMedium + ((i-1) * SIZE(STGMEDIUM))
                lpSource = ADDRESS(pMedium)
                kcr_memcpy(lpDest,lpSource, SIZE(STGMEDIUM))

                SELF.m_nNumFormats += 1
                hr = S_OK
             ELSE
                hr = E_OUTOFMEMORY
             END
          END
       END
    END

    !check for error
    CASE hr
      OF S_OK
         !expected results
    ELSE
         SELF._ShowErrorMessage('IDataObjectClass.SetData',HR)
    END

    RETURN hr


IDataObjectClass.EnumFormatEtc  PROCEDURE(ULONG dwDirection,LONG ppIEnumFORMATETC)   !,HRESULT,PROC,VIRTUAL
hr              HRESULT(S_OK)
pIEnumFormatEtc LONG

  CODE
    IF SELF.IsInitialized = FALSE
       hr = E_FAIL
    ELSE
       IF dwDirection = DATADIR_GET
          pIEnumFormatEtc = ppIEnumFormatEtc
          hr = CreateEnumFormatEtc(SELF.m_nNumFormats, SELF.m_pFormatEtc, ppIEnumFORMATETC)
          !hr = kcr_SHCreateStdEnumFmtEtc(SELF.m_nNumFormats, SELF.m_pFormatEtc, ppIEnumFORMATETC)
       ELSE
          hr = E_NOTIMPL
       END
    END

    !check for error
    CASE hr
      OF S_OK OROF E_NOTIMPL
         !expected results
    ELSE
         SELF._ShowErrorMessage('IDataObjectClass.EnumFormatEtc',HR)
    END

    RETURN hr


IDataObjectClass.DAdvise        PROCEDURE(*FORMATETC pformatetc,ULONG advf,*IAdviseSink pAdvSink,*ULONG pdwConnection)   !,HRESULT,PROC,VIRTUAL
hr          HRESULT(S_OK)

  CODE
    IF SELF.IsInitialized = FALSE
       hr = E_FAIL
    ELSE
       hr = OLE_E_ADVISENOTSUPPORTED
    END

    !check for error
    CASE hr
      OF S_OK OROF OLE_E_ADVISENOTSUPPORTED
         !expected results
    ELSE
         SELF._ShowErrorMessage('IDataObjectClass.DAdvise',HR)
    END

    RETURN hr


IDataObjectClass.DUnadvise      PROCEDURE(ULONG dwConnection)   !,HRESULT,PROC,VIRTUAL
hr          HRESULT(S_OK)

  CODE
    IF SELF.IsInitialized = FALSE
       hr = E_FAIL
    ELSE
       hr = OLE_E_ADVISENOTSUPPORTED
    END

    !check for error
    CASE hr
      OF S_OK OROF OLE_E_ADVISENOTSUPPORTED
         !expected results
    ELSE
         SELF._ShowErrorMessage('IDataObjectClass.DUnadvise',HR)
    END

    RETURN hr


IDataObjectClass.EnumDAdvise    PROCEDURE(LONG ppenumAdvise)    !,HRESULT,PROC,VIRTUAL
hr          HRESULT(S_OK)

  CODE
    IF SELF.IsInitialized = FALSE
       hr = E_FAIL
    ELSE
       hr = OLE_E_ADVISENOTSUPPORTED
    END

    !check for error
    CASE hr
      OF S_OK OROF OLE_E_ADVISENOTSUPPORTED
         !expected results
    ELSE
         SELF._ShowErrorMessage('IDataObjectClass.EnumDAdvise',HR)
    END

    RETURN hr


!--------------------------------------------------------------------------
!IDataObject Interface Implementation
!--------------------------------------------------------------------------
IDataObjectClass.IDataObject.QueryInterface PROCEDURE(REFIID riid, *LONG pvObject)  !,HRESULT,VIRTUAL
  CODE
    RETURN SELF.QueryInterface(riid, pvObject)

IDataObjectClass.IDataObject.AddRef     PROCEDURE() !,LONG,PROC,VIRTUAL
  CODE
    RETURN SELF.AddRef()

IDataObjectClass.IDataObject.Release    PROCEDURE() !,LONG,PROC,VIRTUAL
  CODE
    RETURN SELF.Release()

!IDataObjectClass.IDataObject.GetData  PROCEDURE(*FORMATETC pFormatEtc,*STGMEDIUM pRemoteMedium)  !,HRESULT,PROC,VIRTUAL
IDataObjectClass.IDataObject.GetData  PROCEDURE(LONG pFormatEtc,LONG pRemoteMedium)  !,HRESULT,PROC,VIRTUAL
fmtetc          &FORMATETC
RemoteMedium    &STGMEDIUM
  CODE
    fmtetc &= (pFormatEtc)
    RemoteMedium &= (pRemoteMedium)
    RETURN SELF.GetData(fmtetc, RemoteMedium)

!IDataObjectClass.IDataObject.GetDataHere  PROCEDURE(*FORMATETC pformatetc,*STGMEDIUM pRemoteMedium)    !,HRESULT,PROC,VIRTUAL
IDataObjectClass.IDataObject.GetDataHere  PROCEDURE(LONG pFormatEtc,LONG pRemoteMedium)    !,HRESULT,PROC,VIRTUAL
fmtetc          &FORMATETC
RemoteMedium    &STGMEDIUM
  CODE
    fmtetc &= (pFormatEtc)
    RemoteMedium &= (pRemoteMedium)
    RETURN SELF.GetDataHere(fmtetc, RemoteMedium)

!IDataObjectClass.IDataObject.QueryGetData   PROCEDURE(*FORMATETC pformatetc) !,HRESULT,PROC,VIRTUAL
IDataObjectClass.IDataObject.QueryGetData   PROCEDURE(LONG pFormatEtc) !,HRESULT,PROC,VIRTUAL
fmtetc          &FORMATETC
  CODE
    fmtetc &= (pFormatEtc)
    RETURN SELF.QueryGetData(fmtetc)

!IDataObjectClass.IDataObject.GetCanonicalFormatEtc  PROCEDURE(*FORMATETC pFormatEctIn,*FORMATETC pFormatEtcOut)   !,HRESULT,PROC,VIRTUAL
IDataObjectClass.IDataObject.GetCanonicalFormatEtc  PROCEDURE(LONG pFormatEtcIn,LONG pFormatEtcOut)   !,HRESULT,PROC,VIRTUAL
FormatEtcIn     &FORMATETC
FormatEtcOut    &FORMATETC
  CODE
    FormatEtcIn  &= (pFormatEtcIn)
    FormatEtcOut &= (pFormatEtcOut)
    RETURN SELF.GetCanonicalFormatEtc(FormatEtcIn, FormatEtcOut)

!IDataObjectClass.IDataObject.SetData  PROCEDURE(*FORMATETC pformatetc,*STGMEDIUM pmedium,long fRelease)   !,HRESULT,PROC,VIRTUAL
IDataObjectClass.IDataObject.SetData  PROCEDURE(LONG pFormatEtc,LONG pMedium,LONG fRelease)   !,HRESULT,PROC,VIRTUAL
fmtetc          &FORMATETC
RemoteMedium    &STGMEDIUM

  CODE
    fmtetc &= (pFormatEtc)
    RETURN SELF.SetData(fmtetc, RemoteMedium, fRelease)

IDataObjectClass.IDataObject.EnumFormatEtc  PROCEDURE(ULONG dwDirection,LONG ppenumFormatEtc)   !,HRESULT,PROC,VIRTUAL

  CODE
    RETURN SELF.EnumFormatEtc(dwDirection, ppenumFormatEtc)

!IDataObjectClass.IDataObject.DAdvise    PROCEDURE(*FORMATETC pformatetc,ULONG advf,*IAdviseSink pAdvSink,*ULONG pdwConnection)   !,HRESULT,PROC,VIRTUAL
IDataObjectClass.IDataObject.DAdvise    PROCEDURE(LONG pFormatEtc,ULONG advf,LONG pAdvSink,LONG pdwConnection)   !,HRESULT,PROC,VIRTUAL
fmtetc          &FORMATETC
AdvSink         &IAdviseSink
dwConnection    &ULONG
  CODE
    fmtetc &= (pFormatEtc)
    AdvSink   &= (pAdvSink)
    dwConnection &= (pdwConnection)
    RETURN SELF.DAdvise(fmtetc, advf, AdvSink, dwConnection)

IDataObjectClass.IDataObject.DUnadvise  PROCEDURE(ULONG dwConnection)   !,HRESULT,PROC,VIRTUAL
  CODE
    RETURN SELF.DUnadvise(dwConnection)

IDataObjectClass.IDataObject.EnumDAdvise    PROCEDURE(LONG ppenumAdvise)    !,HRESULT,PROC,VIRTUAL
  CODE
    RETURN SELF.EnumDAdvise(ppenumAdvise)


!==============================================================================
! IEnumFORMATETCClass implementation
!==============================================================================
IEnumFORMATETCClass.Construct   PROCEDURE()
  CODE
    SELF.debug = TRUE


IEnumFORMATETCClass.Destruct    PROCEDURE()
  CODE
    IF SELF.IsInitialized = TRUE
       SELF.Kill()
    END


IEnumFORMATETCClass.Init        PROCEDURE(*FORMATETC paFormatEtc, SIGNED count)
hr          HRESULT(S_OK)
I           LONG,AUTO
dest        LONG,AUTO
source      LONG,AUTO

  CODE
    SELF.m_nIndex = 0
    SELF.m_nNumFormats = count
    SELF.m_pFormatEtc = GlobalAlloc(GPTR, SIZE(FORMATETC) * count )
    IF ~SELF.m_pFormatEtc
       hr = E_OUTOFMEMORY
    ELSE
       LOOP I = 1 TO count
          dest = SELF.m_pFormatEtc + ((I-1) * SIZE(FORMATETC))
          source = ADDRESS(paFormatEtc) + ((I-1) * SIZE(FORMATETC))
          DeepCopyFormatEtc(dest, source)
       END
       SELF.IUnk &= ADDRESS(SELF.IEnumFormatEtc)
       SELF.IEnumFORMATETCObj &= ADDRESS(SELF.IEnumFormatEtc)
       SELF.IsInitialized = TRUE
    END

    !check for error
    CASE hr
      OF S_OK
         !expected results
    ELSE
         SELF._ShowErrorMessage('IEnumFORMATETCClass.Init',HR)
    END

    RETURN hr


IEnumFORMATETCClass.Kill        PROCEDURE()
hr                              HRESULT(S_OK)
I                               LONG,AUTO
fmtetc                          &tagFORMATETC

  CODE
    IF SELF.IsInitialized = FALSE
       hr = E_NOINTERFACE
    ELSE
       LOOP I = 1 TO SELF.m_nNumFormats
          fmtetc &= SELF.m_pFormatEtc + ((I-1) * SIZE(FORMATETC))
          IF fmtetc.ptd <> 0
             CoTaskMemFree(fmtetc.ptd)
          END
       END
       GlobalFree(SELF.m_pFormatEtc)
       SELF.m_pFormatEtc = 0
       IF NOT (SELF.IEnumFORMATETCObj &= NULL)
          SELF.IUnk &= NULL
          SELF.IEnumFORMATETCObj &= NULL
       END
       SELF.IsInitialized = FALSE
       hr = S_OK
    END
    RETURN hr


IEnumFORMATETCClass.GetInterfaceObject    PROCEDURE()

  CODE
    RETURN ADDRESS(SELF.IEnumFORMATETCObj)


IEnumFORMATETCClass.GetLibLocation    PROCEDURE()

  CODE
    RETURN ''


IEnumFORMATETCClass._ShowErrorMessage    PROCEDURE(STRING pMethodName,HRESULT pHR,BYTE pUseDebugMode=1)
hr          HRESULT(S_OK)

  CODE
    IF NOT(SELF.debug = FALSE AND pUseDebugMode = TRUE)
       MESSAGE(pMethodName & ' failed (' & Dec2Hex(pHR) & ')','COM Error',ICON:EXCLAMATION)
    END
    RETURN hr


IEnumFORMATETCClass.QueryInterface procedure(REFIID riid, *LONG pvObject)
hr          HRESULT(S_OK)

  CODE
    pvObject = 0
    IF (SELF.IsEqualIID(riid, ADDRESS(_IUnknown)))
       pvObject = ADDRESS(SELF.IEnumFORMATETCObj)
    ELSIF (SELF.IsEqualIID(riid, ADDRESS(IID_IEnumFORMATETC)))
       pvObject = ADDRESS(SELF.IEnumFORMATETCObj)
    END
    IF (pvObject = 0)
       hr = E_NOINTERFACE
    ELSE
       SELF.AddRef()
       hr = COM_NOERROR
    END

    RETURN hr


IEnumFORMATETCClass.RemoteNext  PROCEDURE(ULONG celt,*FORMATETC rgelt,LONG pceltFetched)
hr          HRESULT(S_OK)
copied      ULONG(0)
dest        LONG,AUTO
source      LONG,AUTO
celtFetched &LONG

  CODE
    IF SELF.IsInitialized = FALSE
       hr = E_FAIL
    ELSE
       LOOP WHILE ((SELF.m_nIndex < SELF.m_nNumFormats) AND (copied < celt))
          source = SELF.m_pFormatEtc + (SELF.m_nIndex * SIZE(FORMATETC))
          dest   = ADDRESS(rgelt) +  (copied * SIZE(FORMATETC))
          DeepCopyFormatEtc(dest, source)
          copied += 1
          SELF.m_nIndex += 1
       END
       IF pceltFetched
          !store result
          celtFetched &= (pceltFetched)
          celtFetched = copied
       END
       hr = CHOOSE(copied = celt,S_OK,S_FALSE)
    END

    !check for error
    CASE hr
      OF S_OK OROF S_FALSE
         !expected results
    ELSE
         SELF._ShowErrorMessage('IEnumFORMATETCClass.RemoteNext',HR)
    END

    RETURN hr


IEnumFORMATETCClass.Skip        PROCEDURE(ULONG celt)
hr          HRESULT(S_OK)

  CODE
    IF SELF.IsInitialized = FALSE
       hr = E_FAIL
    ELSE
       SELF.m_nIndex += celt
       hr = CHOOSE(SELF.m_nIndex < SELF.m_nNumFormats,S_OK,S_FALSE)
    END

    !check for error
    CASE hr
      OF S_OK
         !expected results
    ELSE
         SELF._ShowErrorMessage('IEnumFORMATETCClass.Skip',HR)
    END

    RETURN hr


IEnumFORMATETCClass.Reset       PROCEDURE()
hr          HRESULT(S_OK)

  CODE
    IF SELF.IsInitialized = FALSE
       hr = E_FAIL
    ELSE
       SELF.m_nIndex = 0
    END

    !check for error
    CASE hr
      OF S_OK
         !expected results
    ELSE
       SELF._ShowErrorMessage('IEnumFORMATETCClass.Reset',HR)
    END

    RETURN hr


IEnumFORMATETCClass.Clone       PROCEDURE(LONG ppenum)
hr              HRESULT(S_OK)
cEnumFormatEtc  &IEnumFORMATETCClass

  CODE
    IF SELF.IsInitialized = FALSE
       hr = E_FAIL
    ELSE
       !make a duplicate enumerator
       hr = CreateEnumFormatEtc(SELF.m_nNumFormats, SELF.m_pFormatEtc, ppenum)
       IF hr = S_OK
          !manually set the index state
          cEnumFormatEtc &= (ppenum)
          cEnumFormatEtc.m_nIndex = SELF.m_nIndex
       END
    END

    !check for error
    CASE hr
      OF S_OK
         !expected results
    ELSE
         SELF._ShowErrorMessage('IEnumFORMATETCClass.Clone',HR)
    END

    RETURN hr


!==============================================================================
! IEnumFormatEtc Interface Implementation
!==============================================================================
IEnumFORMATETCClass.IEnumFORMATETC.QueryInterface PROCEDURE(REFIID riid, *LONG pvObject)
  CODE
    RETURN SELF.QueryInterface(riid, pvObject)

IEnumFORMATETCClass.IEnumFORMATETC.AddRef     PROCEDURE() !,LONG,PROC
  CODE
    RETURN SELF.AddRef()

IEnumFORMATETCClass.IEnumFORMATETC.Release    PROCEDURE() !,LONG,PROC
  CODE
    RETURN SELF.Release()

!IEnumFORMATETCClass.IEnumFORMATETC.RemoteNext  PROCEDURE(ULONG celt,*FORMATETC rgelt,*ULONG pceltFetched)
IEnumFORMATETCClass.IEnumFORMATETC.RemoteNext  PROCEDURE(ULONG celt,LONG rgelt,LONG pceltFetched)
Element         &FORMATETC
  CODE
    Element &= (rgelt)
    RETURN SELF.RemoteNext(celt, Element, pceltFetched)

IEnumFORMATETCClass.IEnumFORMATETC.Skip        PROCEDURE(ULONG celt)
  CODE
    RETURN SELF.Skip(celt)

IEnumFORMATETCClass.IEnumFORMATETC.Reset       PROCEDURE()
  CODE
    RETURN SELF.Reset()

IEnumFORMATETCClass.IEnumFORMATETC.Clone       PROCEDURE(LONG ppenum)
  CODE
    RETURN SELF.Clone(ppenum)


!==============================================================================
! IDropSourceClass implementation
!==============================================================================
IDropSourceClass.Construct   PROCEDURE()
  CODE
    SELF.debug = TRUE


IDropSourceClass.Destruct    PROCEDURE()
  CODE
    IF SELF.IsInitialized = TRUE
       SELF.Kill()
    END


IDropSourceClass.Init        PROCEDURE()
hr          HRESULT(S_OK)

  CODE
    SELF.IUnk &= ADDRESS(SELF.IDropSource)
    SELF.IDropSourceObj &= ADDRESS(SELF.IDropSource)
    SELF.IsInitialized = TRUE
    RETURN hr


IDropSourceClass.Kill        PROCEDURE()
hr          HRESULT(S_OK)

  CODE
    IF SELF.IsInitialized = FALSE
       hr = E_NOINTERFACE
    ELSE
       IF NOT (SELF.IDropSourceObj &= NULL)
          SELF.IUnk &= NULL
          SELF.IDropSourceObj &= NULL
       END
       SELF.IsInitialized = FALSE
       hr = S_OK
    END
    RETURN hr


IDropSourceClass.GetInterfaceObject    PROCEDURE()

  CODE
    RETURN ADDRESS(SELF.IDropSourceObj)


IDropSourceClass.GetLibLocation    PROCEDURE()

  CODE
    RETURN ''


IDropSourceClass._ShowErrorMessage    PROCEDURE(STRING pMethodName,HRESULT pHR,BYTE pUseDebugMode=1)
hr                              HRESULT(S_OK)

  CODE
    IF NOT(SELF.debug = FALSE AND pUseDebugMode = TRUE)
       MESSAGE(pMethodName & ' failed (' & Dec2Hex(pHR) & ')','COM Error',ICON:EXCLAMATION)
    END
    RETURN hr


IDropSourceClass.QueryInterface procedure(REFIID riid, *long pvObject)
hr          HRESULT(S_OK)

  CODE
    pvObject = 0
    IF (SELF.IsEqualIID(riid, ADDRESS(_IUnknown)))
       pvObject = ADDRESS(SELF.IDropSourceObj)
    ELSIF (SELF.IsEqualIID(riid, ADDRESS(IID_IDropSource)))
       pvObject = ADDRESS(SELF.IDropSourceObj)
    END

    IF (pvObject = 0)
       hr = E_NOINTERFACE
    ELSE
       SELF.AddRef()
       hr = COM_NOERROR
    END

    RETURN hr


IDropSourceClass.QueryContinueDrag  PROCEDURE(BOOL fEscapePressed, DWORD grfKeyState)  !,HRESULT,PROC
hr          HRESULT(S_OK)

  CODE
    IF SELF.IsInitialized = FALSE
       hr = E_FAIL
    ELSE
       !if the Escape key has been pressed since the last call, cancel the drop
       IF fEscapePressed = TRUE
          hr = DRAGDROP_S_CANCEL
       ELSIF BAND(grfKeyState,MK_LBUTTON) = 0
          hr = DRAGDROP_S_DROP
       ELSE
          hr = S_OK
       END
    END
    RETURN hr


IDropSourceClass.GiveFeedback   PROCEDURE(DWORD dwEffect)   !,HRESULT,PROC
hr          HRESULT(S_OK)

  CODE
    IF SELF.IsInitialized = FALSE
       hr = E_FAIL
    ELSE
       hr = DRAGDROP_S_USEDEFAULTCURSORS
    END
    RETURN hr


!==============================================================================
! IDropSource Interface Implementation
!==============================================================================
IDropSourceClass.IDropSource.QueryInterface PROCEDURE(REFIID riid, *LONG pvObject)
  CODE
    RETURN SELF.QueryInterface(riid, pvObject)

IDropSourceClass.IDropSource.AddRef     PROCEDURE() !,LONG,PROC
  CODE
    RETURN SELF.AddRef()

IDropSourceClass.IDropSource.Release    PROCEDURE() !,LONG,PROC
  CODE
    RETURN SELF.Release()

IDropSourceClass.IDropSource.QueryContinueDrag  PROCEDURE(BOOL fEscapePressed, DWORD grfKeyState)  !,HRESULT,PROC
  CODE
    RETURN SELF.QueryContinueDrag(fEscapePressed, grfKeyState)

IDropSourceClass.IDropSource.GiveFeedback   PROCEDURE(DWORD dwEffect)   !,HRESULT,PROC
  CODE
    RETURN SELF.GiveFeedback(dwEffect)



!==============================================================================
! IDropTargetClass implementation
!==============================================================================
IDropTargetClass.Construct   PROCEDURE()
  CODE
    SELF.debug = TRUE


IDropTargetClass.Destruct    PROCEDURE()
  CODE
    IF SELF.IsInitialized = TRUE
       SELF.Kill()
    END


IDropTargetClass.Init        PROCEDURE(HWND hWnd, LONG OleDropEvent)
hr          HRESULT(S_OK)

  CODE
    SELF.IUnk &= ADDRESS(SELF.kcr_IDropTarget)
    SELF.IDropTargetObj &= ADDRESS(SELF.kcr_IDropTarget)
    SELF.m_hWnd = hWnd
    SELF.m_OleDropEvent = OleDropEvent
    hr = SELF.Register()
    IF hr = S_OK
       SetWindowPos(hWnd,HWND_TOP,0,0,0,0,BOR(BOR(SWP_NOMOVE,SWP_NOSIZE),SWP_NOREDRAW))
       SELF.IsInitialized = TRUE
    END
    RETURN hr


IDropTargetClass.Kill        PROCEDURE()
hr          HRESULT(S_OK)

  CODE
    IF SELF.IsInitialized = FALSE
       hr = E_NOINTERFACE
    ELSE
       IF NOT (SELF.IDropTargetObj &= NULL)
          IF RevokeDragDrop(SELF.m_hwnd) = S_OK
          END
          SELF.IUnk &= NULL
          SELF.IDropTargetObj &= NULL
       END
       SELF.IsInitialized = FALSE
       hr = S_OK
    END
    RETURN hr


IDropTargetClass.GetInterfaceObject    PROCEDURE()

  CODE
    RETURN ADDRESS(SELF.IDropTargetObj)


IDropTargetClass.GetLibLocation    PROCEDURE()

  CODE
    RETURN ''


IDropTargetClass.Register PROCEDURE()
hr          HRESULT

  CODE
    hr = RegisterDragDrop(SELF.m_hWnd, ADDRESS(SELF.IDropTargetObj))
    IF hr = DRAGDROP_E_ALREADYREGISTERED
       hr = S_OK
    END
    RETURN hr

IDropTargetClass.QueryDataObject    PROCEDURE(*IDataObject pDataObject)
hr          HRESULT
fmtetc      LIKE(FORMATETC)
  CODE
    fmtetc.cfFormat  = CF_TEXT
    fmtetc.ptd       = 0
    fmtetc.dwAspect  = DVASPECT_CONTENT
    fmtetc.lindex    = -1
    fmtetc.tymed     = TYMED_HGLOBAL
    hr = CHOOSE(pDataObject.QueryGetData(ADDRESS(fmtetc)) = S_OK,TRUE,FALSE)
    RETURN hr


IDropTargetClass.DropEffect         PROCEDURE(DWORD grfKeyState, *POINT pt, DWORD dwAllowed)
dwEffect            DWORD
  CODE
    !1. check 'pt' -> do we allow a drop at the specified coordinates?

    !2. work out what the drop-effect should be based on grfKeyState
    IF BAND(grfKeyState,MK_CONTROL) AND BAND(grfKeyState,MK_SHIFT)
       dwEffect = BAND(dwAllowed,DROPEFFECT_LINK)
    ELSIF BAND(grfKeyState,MK_CONTROL)
       dwEffect = BAND(dwAllowed,DROPEFFECT_COPY)
    ELSIF BAND(grfKeyState,MK_SHIFT)
       dwEffect = BAND(dwAllowed,DROPEFFECT_MOVE)
    ELSE
       !3. no key-modifiers were specified (or drop effect not allowed), so
       !   base the effect on those allowed by the dropsource
       IF BAND(dwAllowed, DROPEFFECT_COPY)
          dwEffect = DROPEFFECT_COPY
       END
       IF BAND(dwAllowed, DROPEFFECT_MOVE)
          dwEffect = DROPEFFECT_MOVE
       END
       IF BAND(dwAllowed, DROPEFFECT_LINK)
          dwEffect = DROPEFFECT_LINK
       END
    END
    RETURN dwEffect


IDropTargetClass._ShowErrorMessage    PROCEDURE(STRING pMethodName,HRESULT pHR,BYTE pUseDebugMode=1)
hr                              HRESULT(S_OK)

  CODE
    IF NOT(SELF.debug = FALSE AND pUseDebugMode = TRUE)
       MESSAGE(pMethodName & ' failed (' & Dec2Hex(pHR) & ')','COM Error',ICON:EXCLAMATION)
    END
    RETURN hr


IDropTargetClass.QueryInterface procedure(REFIID riid, *long pvObject)
hr          HRESULT

  CODE
    pvObject = 0
    IF (SELF.IsEqualIID(riid, ADDRESS(_IUnknown)))
       pvObject = ADDRESS(SELF.IDropTargetObj)
    ELSIF (SELF.IsEqualIID(riid, ADDRESS(IID_IDropTarget)))
       pvObject = ADDRESS(SELF.IDropTargetObj)
    END
    IF (pvObject = 0)
       hr = E_NOINTERFACE
    ELSE
       SELF.AddRef()
       hr = COM_NOERROR
    END
    RETURN hr


IDropTargetClass.DragEnter         PROCEDURE(*IDataObject pDataObject, DWORD grfKeyState, POINT pt, *DWORD pdwEffect)
hr          HRESULT(S_OK)
  CODE
    !does the dataobject contain data we want?
    SELF.m_fAllowDrop = SELF.QueryDataObject(pDataObject)
    IF SELF.m_fAllowDrop
       pdwEffect = SELF.DropEffect(grfKeyState, pt, pdwEffect)
       SetFocus(SELF.m_hWnd)
    ELSE
       pdwEffect = DROPEFFECT_NONE
    END
    RETURN hr


IDropTargetClass.DragOver          PROCEDURE(DWORD grfKeyState, POINT pt, *DWORD pdwEffect)
hr          HRESULT(S_OK)
  CODE
    IF SELF.m_fAllowDrop
       pdwEffect = SELF.DropEffect(grfKeyState, pt, pdwEffect)
    ELSE
       pdwEffect = DROPEFFECT_NONE
    END
    RETURN hr

IDropTargetClass.DragLeave         PROCEDURE()
hr          HRESULT(S_OK)
  CODE
    RETURN hr

IDropTargetClass.Drop              PROCEDURE(*IDataObject pDataObject, DWORD grfKeyState, POINT pt, *DWORD pdwEffect)
hr          HRESULT(S_OK)
  CODE
    IF SELF.m_fAllowDrop
       pdwEffect = SELF.DropEffect(grfKeyState, pt, pdwEffect)
    ELSE
       pdwEffect = DROPEFFECT_NONE
    END

    RETURN hr

!==============================================================================
! IDropTarget Interface Implementation
!==============================================================================
IDropTargetClass.kcr_IDropTarget.QueryInterface PROCEDURE(REFIID riid, *LONG pvObject)
  CODE
    RETURN SELF.QueryInterface(riid, pvObject)

IDropTargetClass.kcr_IDropTarget.AddRef     PROCEDURE() !,LONG,PROC
  CODE
    RETURN SELF.AddRef()

IDropTargetClass.kcr_IDropTarget.Release    PROCEDURE() !,LONG,PROC
  CODE
    RETURN SELF.Release()

IDropTargetClass.kcr_IDropTarget.DragEnter  PROCEDURE(LONG pDataObject, DWORD grfKeyState, LONG xPos, LONG yPos, *LONG pdwEffect)
DataObject  &IDataObject
pt          LIKE(POINT)
dwEffect    &DWORD
  CODE
    DataObject &= (pDataObject)
    pt.x = xPos
    pt.y = yPos
    dwEffect &= ADDRESS(pdwEffect)
    RETURN SELF.DragEnter(DataObject, grfKeyState, pt, dwEffect)

IDropTargetClass.kcr_IDropTarget.DragOver   PROCEDURE(DWORD grfKeyState, LONG xPos, LONG yPos, *LONG pdwEffect)
pt          LIKE(POINT)
dwEffect    &DWORD
  CODE
    pt.x = xPos
    pt.y = yPos
    dwEffect &= ADDRESS(pdwEffect)
    RETURN SELF.DragOver(grfKeyState, pt, dwEffect)

IDropTargetClass.kcr_IDropTarget.DragLeave  PROCEDURE()
  CODE
    RETURN SELF.DragLeave()

IDropTargetClass.kcr_IDropTarget.Drop       PROCEDURE(LONG pDataObject, DWORD grfKeyState, LONG xPos, LONG yPos, *LONG pdwEffect)
DataObject  &IDataObject
pt          LIKE(POINT)
dwEffect    &DWORD
  CODE
    DataObject &= (pDataObject)
    pt.x = xPos
    pt.y = yPos
    dwEffect &= ADDRESS(pdwEffect)
    RETURN SELF.Drop(DataObject, grfKeyState, pt, dwEffect)


!==============================================================================
! OleDataSourceClass Implementation
!==============================================================================
OleDataSourceClass.GetOleData PROCEDURE()  !,LONG,VIRTUAL
!Virtual Place Holder
!used to prepare data buffer before drag operation
!returns pointer to heap data
  CODE
    RETURN 0


OleDataSourceClass.OnLButtonDown PROCEDURE()  !,LONG,VIRTUAL
!Virtual Place Holder
!used to detect the drag operation
  CODE
    RETURN 0


!==============================================================================
! TextDataObjectClass Implementation
!==============================================================================
TextDataObjectClass.Init    PROCEDURE(*CSTRING szText)  !,HRESULT,PROC,VIRTUAL
hr  HRESULT
  CODE
    hr = PARENT.Init()
    RETURN hr


!==============================================================================
! cOleIniter Implementation
!==============================================================================
cOleIniter.Construct    PROCEDURE()
  CODE
    kcr_OleInitialize(0)
    RETURN

cOleIniter.Destruct     PROCEDURE() !,VIRTUAL
  CODE
    kcr_OleUnInitialize()
