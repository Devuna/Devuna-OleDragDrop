#! ================================================================================
#!                        Devuna OLE Drag and Drop Templates
#! ================================================================================
#! Notice : Copyright (C) 2017, Devuna
#!          Distributed under the MIT License (https://opensource.org/licenses/MIT)
#!
#!    This file is part of Devuna-OleDragDrop (https://github.com/Devuna/Devuna-OleDragDrop)
#!
#!    Devuna-OleDragDrop is free software: you can redistribute it and/or modify
#!    it under the terms of the MIT License as published by
#!    the Open Source Initiative.
#!
#!    Devuna-OleDragDrop is distributed in the hope that it will be useful,
#!    but WITHOUT ANY WARRANTY; without even the implied warranty of
#!    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#!    MIT License for more details.
#!
#!    You should have received a copy of the MIT License
#!    along with Devuna-OleDragDrop.  If not, see <https://opensource.org/licenses/MIT>.
#! ================================================================================
#!
#!Author:        Randy Rogers <rrogers@devuna.com>
#!
#!Revisions:
#!==========
#!2017.02.02 Devuna rebuild for Clarion 10.0.12463
#!2015.03.06 Devuna rebuild for Clarion 10
#!2011.03.22 KCR Changed fmtetc and stgmed in .DROP method to be reference variables.
#!2011.02.15 KCR Added more virtual methods and modified to support enhanced data object class
#!2006.05.04 KCR changed generated code to use REGISTERed event handler
#!           KCR Added more embeds
#!           KCR removed Key State Modifier as drag trigger
#!           KCR template changed so programmer can decide what drag effects are ok
#!2006.04.27 KCR Added more embeds and derived methods to the IDropTarget Class to allow
#!               the programmer to support other drop types besides text.
#!2006.04.25 KCR Added support for LEGACY templates (thanks to John Hickey for his assistance)
#!2006.04.24 KCR Added ability to select file field as data buffer
#!2006.04.20 KCR Added ability to select an existing global CSTRING as default buffer
#!               Added buffer override to KCR_OleDragDrop template
#!               Added new Text Replace Selected default behaviour
#!               Cleaned up template appearance
#!2006.04.19 KCR fixed api prototype that caused gpf
#!               use NEW for FORMATETC and STGMEDIUM groups to force 8-byte boundry
#!               added some more embed points and template options
#!2006.04.17 KCR template enhancements and bug fixes
#!2006.04.07 KCR initial version
#!=================================================================================================
#TEMPLATE(KCR_OleDragDrop,'Devuna OLE Drag and Drop ABC Templates'),FAMILY('ABC'),FAMILY('CW20')
#HELP('OLEDRAGDROP.HLP')
#!
#!
#!
#!
#!=================================================================================================
#EXTENSION(KCR_GlobalOleDragDrop,'Devuna Global OLE Drag and Drop Extension'),APPLICATION
#!=================================================================================================
#BOXED('Default MakeHead Prompts'),AT(0,0),WHERE(%False),HIDE
  #INSERT(%MakeHeadHiddenPrompts)
#ENDBOXED
#PREPARE
  #INSERT (%MakeHead,'KCR_GlobalOleDragDrop (Devuna)','OLE Drag and Drop Extension','2017.02.02')
  #DECLARE(%GlobalCSTRINGs),UNIQUE
  #DECLARE(%GlobalStringSize,%GlobalCSTRINGs)
  #FOR(%GlobalData),WHERE(EXTRACT(%GlobalDataStatement, 'CSTRING') AND EXTRACT(%GlobalDataStatement, 'THREAD'))
    #ADD(%GlobalCSTRINGs, %GlobalData)
    #SET(%GlobalStringSize, (EXTRACT(%GlobalDataStatement, 'CSTRING', 1) - 1))
  #ENDFOR
#ENDPREPARE
#BOXED('Devuna')
  #INSERT (%Head)
  #DISPLAY('This adds the required OLE Drag and Drop')
  #DISPLAY('global declarations to support the CF_TEXT format.')
  #DISPLAY('')
  #DISPLAY('The template needs a buffer to facilitate data transfer.')
  #DISPLAY('You can Create or Select a global variable here to be')
  #DISPLAY('used as the default for the KCR_OleDragDrop template.')
  #DISPLAY('The default can be overridden.  If you choose None,')
  #DISPLAY('then no default Ole Data Buffer will be supplied.')
  #PROMPT('Ole Data Buffer Options',OPTION),%OleDataBufferOption
  #PROMPT('Create',RADIO),AT(16)
  #PROMPT('Select',RADIO),AT(16)
  #PROMPT('None',RADIO),AT(16)
  #BOXED,WHERE(%OleDataBufferOption = 'Create'),AT(,160)
    #PROMPT('Ole Data Buffer label:',@S40),%OleDataBufferCreate,DEFAULT('OleDataBuffer'),WHENACCEPTED(%SetOleDataBuffer(%OleDataBufferCreate))
    #PROMPT('Ole Data Buffer size :',SPIN(@N10,1,4194303,1024)),%OleDataBufferCreateSize,DEFAULT(4096),WHENACCEPTED(%SetOleDataBufferSize(%OleDataBufferCreateSize))
    #DISPLAY('Make the size big enough to accommodate the data')
    #DISPLAY('you expect to drag or be dropped on your application.')
    #DISPLAY('Drops on your application will be truncated if the')
    #DISPLAY('data exceeds the supplied buffer size.')
  #ENDBOXED
  #BOXED,WHERE(%OleDataBufferOption = 'Select'),AT(,160)
    #PROMPT('Ole Data Buffer label:',FROM(%GlobalCSTRINGs)),%OleDataBufferSelect,WHENACCEPTED(%SetOleData(%OleDataBufferSelect))
  #ENDBOXED
  #PROMPT('&Adjust Drag Detect Rectangle',CHECK),%AdjustDragDetect,AT(10),DEFAULT(%FALSE)
  #BOXED,WHERE(%AdjustDragDetect)
    #PROMPT('Drag Height',@N3),%DragHeight,DEFAULT(4)
    #PROMPT('Drag Width',@N3),%DragWidth,DEFAULT(4)
  #ENDBOXED
  #DISPLAY('')
  #PROMPT('&Do Not Generate Any Code',CHECK),%DoNotGenerate,AT(10),DEFAULT(%FALSE)
  #DISPLAY ('')
  #BOXED('Hidden Prompts'),HIDE
  #PROMPT('',@S40),%OleDataBuffer,DEFAULT('OleDataBuffer')
  #PROMPT('',@N10),%OleDataBufferSize,DEFAULT(4096)
  #PROMPT('',@N1),%KeystoneApiPresent,DEFAULT(%FALSE)
  #ENDBOXED
#ENDBOXED
#!
#!
#ATSTART
  #FOR(%ApplicationTemplate),WHERE(%ApplicationTemplate = 'KCR_Win32(KCR)')
    #SET(%KeystoneApiPresent,%TRUE)
    #BREAK
  #ENDFOR
#ENDAT
#!
#!
#AT(%CustomGlobalDeclarations),WHERE(NOT %DoNotGenerate)
    #PDEFINE('_svDllMode_',0)
    #PDEFINE('_svLinkMode_',1)
    #PDEFINE('_ComDllMode_',0)
    #PDEFINE('_ComLinkMode_',1)
#ENDAT
#!
#!
#AT(%AfterGlobalIncludes),WHERE(NOT %DoNotGenerate)
   INCLUDE('svcom.inc'),ONCE
   INCLUDE('cDataObject.inc'),ONCE

#IF(%AdjustDragDetect)
kcrDragHeight       LONG
kcrDragWidth        LONG
#ENDIF
szPropFieldEquate   CSTRING('FieldEquate')
szPropDragData      CSTRING('DragData')
szPropDataObject    CSTRING('DataObject')
szPropDropTarget    CSTRING('DropTarget')
  #IF(%OleDataBufferOption = 'Create')
%20OleDataBufferCreate CSTRING(%OleDataBufferCreateSize + 1),THREAD
  #ENDIF
#ENDAT
#!
#!
#AT(%GlobalMap),WHERE((NOT %DoNotGenerate) AND (NOT %KeystoneApiPresent))
MODULE('kernel32')
  kcr_CallWindowProc(LONG lpPrevWndProc, HWND hwnd, UNSIGNED nMsg, UNSIGNED wParam, LONG lParam),LONG,PASCAL,NAME('CallWindowProcA')
  kcr_ClientToScreen(HWND hwnd, *POINT pt),BOOL,PASCAL,RAW,PROC,NAME('ClientToScreen')
  kcr_DoDragDrop(LONG pDataObj, LONG pDropSource, LONG dwOKEffects, *DWORD dwEffect),LONG,PASCAL,RAW,NAME('DoDragDrop')
  kcr_DragDetect(HWND hWnd, LONG xPos, LONG yPos),BOOL,PASCAL,RAW,NAME('DragDetect')
  kcr_DragQueryFile(HDROP hDrop, UNSIGNED iFile, *CSTRING lpszFile,UNSIGNED cch),UNSIGNED,PASCAL,RAW,PROC,NAME('DragQueryFile')
  kcr_GetCapture(),HWND,PASCAL,NAME('GetCapture')
  kcr_GetCursorPos(*POINT lpPoint),BOOL,RAW,PASCAL,PROC,NAME('GetCursorPos')
  kcr_GetProp(HWND hwnd, *CSTRING szPropertyName),HANDLE,RAW,PASCAL,NAME('GetPropA')
  kcr_GetSystemMetrics(LONG nIndex),LONG,PASCAL,NAME('GetSystemMetrics')
  kcr_GlobalAlloc(LONG uFlags, LONG dwBytes),LONG,PASCAL,NAME('GlobalAlloc')
  kcr_GlobalFree(HGLOBAL hMem),HGLOBAL,PASCAL,PROC,NAME('GlobalFree')
  kcr_GlobalHandle(LONG lpvMem),HGLOBAL,NAME('GlobalHandle'),PASCAL,NAME('GlobalHandle')
  kcr_GlobalLock(HGLOBAL hMem),LONG,PASCAL,PROC,NAME('GlobalLock')
  kcr_GlobalSize(HGLOBAL hMem),DWORD,NAME('GlobalSize'),PASCAL,NAME('GlobalSize')
  kcr_GlobalUnlock(HGLOBAL hMem),BOOL,PASCAL,PROC,NAME('GlobalUnlock')
  kcr_MemCpy(ULONG pDest, ULONG pSrc, UNSIGNED nCount),name('_memcpy')
  kcr_MemSet(ULONG pDest, LONG ch, UNSIGNED nBytes),name('_memset')
  kcr_OleInitialize(LONG pvReserved),LONG,PASCAL,PROC,NAME('OleInitialize')
  kcr_OleUninitialize(),PASCAL,NAME('OleUninitialize')
  kcr_OleGetClipboard(*IDataObject ppDataObject),HRESULT,RAW,PASCAL,PROC,NAME('OleGetClipboard')
  kcr_OleSetClipboard(*IDataObject ppDataObject),HRESULT,RAW,PASCAL,PROC,NAME('OleSetClipboard')
  kcr_OleSetClipboard(LONG ppDataObject),HRESULT,RAW,PASCAL,PROC,NAME('OleSetClipboard')
  kcr_OleFlushClipboard(),HRESULT,RAW,PASCAL,PROC,NAME('OleFlushClipboard')
  kcr_PostMessage(HWND hwnd, UNSIGNED wmag, UNSIGNED wparam, LONG lparam),BOOL,PASCAL,PROC,NAME('PostMessageA')
  kcr_RegisterClipboardFormat(LONG lpszFormat),LONG,PASCAL,NAME('RegisterClipboardFormatA')
  kcr_ReleaseCapture(),PASCAL,NAME('ReleaseCapture')
  kcr_ReleaseStgMedium(*STGMEDIUM stgmed),RAW,PASCAL,NAME('ReleaseStgMedium')
  kcr_RemoveProp(HWND hwnd, *CSTRING szPropertyName),HANDLE,RAW,PASCAL,PROC,NAME('RemovePropA')
  kcr_SendMessage(HWND hwnd, UNSIGNED wmag, UNSIGNED wparam, LONG lparam),LONG,PASCAL,PROC,NAME('SendMessageA')
  kcr_SetCapture(HWND hwnd),HWND,PASCAL,PROC,NAME('SetCapture')
  kcr_SetFocus(HWND hwnd),HWND,PASCAL,PROC,NAME('SetFocus')
  kcr_SetProp(HWND hwnd, *CSTRING szPropertyName, HANDLE hData),BOOL,RAW,PASCAL,PROC,NAME('SetPropA')
  kcr_SystemParametersInfo(UNSIGNED uAction, UNSIGNED uParam, LONG pvParam, UNSIGNED fUpdateProfile),BOOL,PROC,PASCAL,RAW,NAME('SystemParametersInfoA')
  kcr_WindowFromPoint(POINT pt),HWND,RAW,PASCAL,NAME('WindowFromPoint')
  kcr_ZeroMemory(LONG DestinationPtr,LONG dwLength),RAW,PASCAL,NAME('RtlZeroMemory')
END
#ENDAT
#!
#!
#AT(%ProgramSetup)
#IF(~%DoNotGenerate)
  #IF(%ProgramExtension = 'EXE' AND %AdjustDragDetect)
kcrDragWidth  = kcr_GetSystemMetrics(SM_CXDRAG)
kcrDragHeight = kcr_GetSystemMetrics(SM_CYDRAG)
kcr_SystemParametersInfo(76,%DragWidth,0,1)     #<!SPI_SETDRAGWIDTH
kcr_SystemParametersInfo(77,%DragHeight,0,3)    #<!SPI_SETDRAGHEIGHT
  #ENDIF
#ENDIF
#ENDAT
#!
#!
#AT(%ProgramEnd)
#IF(~%DoNotGenerate)
  #IF(%ProgramExtension = 'EXE' AND %AdjustDragDetect)
kcr_SystemParametersInfo(76,4,0,1)   #<!SPI_SETDRAGWIDTH
kcr_SystemParametersInfo(77,4,0,3)  #<!SPI_SETDRAGHEIGHT
!kcr_SystemParametersInfo(76,kcrDragWidth,0,1)   #<!SPI_SETDRAGWIDTH
!kcr_SystemParametersInfo(77,kcrDragHeight,0,3)  #<!SPI_SETDRAGHEIGHT
  #ENDIF
#ENDIF
#ENDAT
#!
#!
#GROUP(%SetOleDataBuffer,%DataBuffer,%DataBufferSize)
#SET(%OleDataBuffer,%DataBuffer)
#!
#!
#GROUP(%SetOleDataBufferSize,%DataBufferSize)
#SET(%OleDataBufferSize,%DataBufferSize)
#!
#!
#GROUP(%SetOleData,%DataBuffer)
#SET(%OleDataBuffer,%DataBuffer)
#FIX(%GlobalCSTRINGs,%DataBuffer)
#SET(%OleDataBufferSize,%GlobalStringSize)
#!
#!
#!
#!
#!=================================================================================================
#EXTENSION(KCR_OleDragDrop,'Devuna OLE Drag and Drop CF_TEXT Extension'),REQ(KCR_GlobalOleDragDrop),MULTI
#!=================================================================================================
#!
#PREPARE
#DECLARE(%TemplateCount)
#SET(%TemplateCount,0)
#DECLARE(%EventDefault)
#SET(%EventDefault,36863)
#DECLARE(%CurrentInstance)
#SET(%CurrentInstance,%ActiveTemplateInstance)
#FOR(%ActiveTemplate),WHERE(%ActiveTemplate = 'KCR_OleDragDrop(KCR_OleDragDrop)')
  #FOR(%ActiveTemplateInstance)
    #SET(%TemplateCount, %TemplateCount + 1)
    #IF(%ActiveTemplateInstance = %CurrentInstance)
      #IF(NOT %EventOleDrop)
        #SET(%EventOleDrop,%EventDefault+%TemplateCount)
      #ENDIF
    #ENDIF
  #ENDFOR
#ENDFOR
#!
#DECLARE(%DragDropControls),MULTI
#FOR(%Control)
  #CASE(%ControlType)
  #OF('LIST')
  #OROF('TEXT')
  #OROF('ENTRY')
    #ADD(%DragDropControls,%Control)
  #ENDCASE
#ENDFOR
#DECLARE(%AllCSTRINGs),UNIQUE
#DECLARE(%StringSize,%AllCSTRINGs)
#CALL(%LoadCSTRINGs)  #!Call a #GROUP,PRESERVE just in case
#IF(NOT %OleDataBufferOverride)
  #SET(%OleDataBufferOverride, %OleDataBuffer)
  #SET(%OleDataBufferOverrideSize, %OleDataBufferSize)
#ENDIF
#ENDPREPARE
  #DISPLAY('This template adds OLE Drag and Drop CF_TEXT support to')
  #DISPLAY('a window control.')
  #DISPLAY('')
  #PROMPT('Ole Data Buffer:',FROM(%AllCSTRINGs)),%OleDataBufferOverride,WHENACCEPTED(%SetOleDataOverrideSize(%OleDataBufferOverride)),REQ
  #DISPLAY('')
  #PROMPT('Drag and Drop &Control:',FROM(%DragDropControls)),%DragDropControl,REQ
  #DISPLAY('')
  #PROMPT('&Include Drag Support',CHECK),%IncludeDrag,AT(10),DEFAULT(%TRUE)
  #ENABLE(%IncludeDrag)
    #BOXED('Effects Allowed'),AT(10,,180)
      #PROMPT('Copy',CHECK),%CopyEffect,DEFAULT(%TRUE),AT(10)
      #PROMPT('Move',CHECK),%MoveEffect,DEFAULT(%FALSE),AT(10)
      #PROMPT('Link',CHECK),%LinkEffect,DEFAULT(%FALSE),AT(10)
    #ENDBOXED
    #PROMPT('&Generate Default Drag Behaviour',CHECK),%GenerateDragDefault,AT(10),DEFAULT(%TRUE)
    #BOXED('Default Behaviour'),WHERE(%GenerateDragDefault),AT(10,,180)
      #PROMPT('Include Column Heading',CHECK),%IncludeColumnHeading,DEFAULT(%FALSE),AT(10)
      #PROMPT('&Field Delimiter',OPTION),%FieldDelimiter,DEFAULT('CRLF'),AT(10,,170,35)
      #PROMPT('CRLF',RADIO),AT(16,178)
      #PROMPT('TAB',RADIO),AT(16)
    #ENDBOXED
  #ENDENABLE
  #PROMPT('&Include Drop Support',CHECK),%IncludeDrop,AT(10),DEFAULT(%TRUE)
  #ENABLE(%IncludeDrop)
    #PROMPT('&Generate Default Drop Behaviour',CHECK),%GenerateDropDefault,AT(10),DEFAULT(%TRUE)
    #BOXED('Default Behaviour'),WHERE(%GenerateDropDefault),AT(10,,180)
      #PROMPT('OleDrop Event Value:',SPIN(@n5,32768,49151)),%EventOleDrop,REQ,AT(,,40)
      #PROMPT('Text/Entry Default Drop Behaviour',OPTION),%DefaultDropBehaviour,DEFAULT('Replace'),AT(10,,170,48)
      #PROMPT('Append',RADIO),AT(16,264)
      #PROMPT('Replace All',RADIO),AT(16)
      #PROMPT('Replace Selection',RADIO),AT(16)
    #ENDBOXED
  #ENDENABLE
  #BOXED('Hidden Prompts'),HIDE
  #PROMPT('',@N3),%FirstInstance
  #PROMPT('',@N10),%OleDataBufferOverrideSize
  #PROMPT('',@N3),%GenerateLegacy
  #ENDBOXED
#!
#!
#ATSTART
  #DECLARE(%EmbedNumber)
  #DECLARE(%CurrentInstance)
  #DECLARE(%TemplateCount)
  #SET(%TemplateCount,0)
  #SET(%FirstInstance,0)
  #FOR(%ActiveTemplate),WHERE(%ActiveTemplate = 'KCR_OleDragDrop(KCR_OleDragDrop)')
    #FOR(%ActiveTemplateInstance)
      #SET(%TemplateCount, %TemplateCount + 1)
      #IF(%FirstInstance = 0)
        #SET(%FirstInstance,%ActiveTemplateInstance)
      #ENDIF
    #ENDFOR
  #ENDFOR
#!
  #DECLARE(%DataObject)
  #SET(%DataObject,'OleDataObject' & %ActiveTemplateInstance)
  #DECLARE(%DropTarget)
  #SET(%DropTarget,'OleDropTarget' & %ActiveTemplateInstance)
  #DECLARE(%DropTargetEventName)
  #SET(%DropTargetEventName,'EVENT:OleDropTarget' & %ActiveTemplateInstance & ':OleDrop')
#!
  #DECLARE(%Delimiter)
  #DECLARE(%TempConstruct)
#!
  #CASE(%FieldDelimiter)
  #OF('CRLF')
    #SET(%Delimiter,'''<<0DH,0AH>''')
  #OF('TAB')
    #SET(%Delimiter,'''<<09H>''')
  #ENDCASE
#!
  #IF(SUB(%CWTemplateVersion,1,2)='v2' OR (VAREXISTS(%AppTemplateFamily) AND %AppTemplateFamily='CLARION'))
    #SET(%GenerateLegacy,%TRUE)
  #ENDIF
#!
  #IF( NOT((%CopyEffect = %TRUE) OR (%MoveEffect = %TRUE) OR (%LinkEffect = %TRUE)) )
    #SET(%ValueConstruct,'Devuna OLE Drag and Drop Extension(' & %ActiveTemplateInstance & ')')
    #ERROR(%ValueConstruct)
    #ERROR('At least one ''Effects Allowed'' must be selected')
    #ABORT
  #ENDIF
#ENDAT
#!
#!
#AT(%CustomModuleDeclarations),PRIORITY(5000),DESCRIPTION('Ole Drag and Drop'),WHERE((NOT %DoNotGenerate) AND %IncludeDrag)
    #ADD(%CustomModuleMapModule,'CURRENT MODULE')
    #SET(%ValueConstruct,SUB(%DragDropControl,2,LEN(%DragDropControl)-1) & ':WndProc' & %ActiveTemplateInstance)
    #ADD(%CustomModuleMapProcedure,%ValueConstruct)
    #IF(SUB(%CWTemplateVersion,1,2)='v2' OR (VAREXISTS(%AppTemplateFamily) AND %AppTemplateFamily='CLARION'))
      #SET(%CustomModuleMapProcedurePrototype,'(UNSIGNED hWnd, UNSIGNED wMsg, UNSIGNED wParam, LONG lParam),LONG,PASCAL')
    #ELSE
      #SET(%CustomModuleMapProcedurePrototype,'PROCEDURE(UNSIGNED hWnd, UNSIGNED wMsg, UNSIGNED wParam, LONG lParam),LONG,PASCAL')
    #ENDIF
#ENDAT
#!
#!
#AT(%ModuleDataSection),PRIORITY(5000),DESCRIPTION('Ole Drag and Drop'),WHERE(NOT %DoNotGenerate)
  #IF(%IncludeDrag)
    #SET(%ValueConstruct,SUB(%DragDropControl,2,LEN(%DragDropControl)-1) & ':origWndProc' & %ActiveTemplateInstance)
%30ValueConstruct LONG
  #ENDIF
#ENDAT
#!
#!
#AT(%DataSectionBeforeWindow),WHERE(NOT %DoNotGenerate AND %GenerateLegacy)
  #IF(%ActiveTemplateInstance = %FirstInstance)
!------------------------------------------------------------------------------
OleIniter            cOleIniter #<!Ole Initialization

  #ENDIF
  #IF(%IncludeDrop)
%20DropTargetEventName EQUATE(%EventOleDrop)

  #ENDIF
  #IF(%IncludeDrag OR %IncludeDrop)
#SET(%ValueConstruct,'OleDataObject' & %ActiveTemplateInstance)
%20ValueConstruct CLASS(TextDataObjectClass)
Init                    PROCEDURE(*CSTRING szText),HRESULT,PROC,DERIVED
Kill                    PROCEDURE(),HRESULT,PROC,DERIVED
_ShowErrorMessage       PROCEDURE(STRING pMethodName,HRESULT pHR,BYTE pUseDebugMode=true),HRESULT,PROC,DERIVED
QueryInterface          PROCEDURE(LONG riid, *LONG pvObject),HRESULT,DERIVED
AddRef                  PROCEDURE(),LONG,PROC,DERIVED
Release                 PROCEDURE(),LONG,PROC,DERIVED
GetData                 PROCEDURE(*FORMATETC pformatetcIn,*STGMEDIUM pRemoteMedium),HRESULT,PROC,DERIVED
GetDataHere             PROCEDURE(*FORMATETC pformatetc,*STGMEDIUM pRemoteMedium),HRESULT,PROC,DERIVED
QueryGetData            PROCEDURE(*FORMATETC pformatetc),HRESULT,PROC,DERIVED
GetCanonicalFormatEtc   PROCEDURE(*FORMATETC pformatectIn,*FORMATETC pformatetcOut),HRESULT,PROC,DERIVED
SetData                 PROCEDURE(*FORMATETC pformatetc,*STGMEDIUM pmedium,LONG fRelease),HRESULT,PROC,DERIVED
EnumFormatEtc           PROCEDURE(ULONG dwDirection,LONG ppenumFormatEtc),HRESULT,PROC,DERIVED
DAdvise                 PROCEDURE(*FORMATETC pformatetc,ULONG advf,*IAdviseSink pAdvSink,*ULONG pdwConnection),HRESULT,PROC,DERIVED
DUnadvise               PROCEDURE(ULONG dwConnection),HRESULT,PROC,DERIVED
EnumDAdvise             PROCEDURE(LONG ppenumAdvise),HRESULT,PROC,DERIVED
                     END

  #ENDIF
  #IF(%IncludeDrag)
#SET(%ValueConstruct,'OleDataSource' & %ActiveTemplateInstance)
%20ValueConstruct CLASS(OleDataSourceClass)
GetOleData              PROCEDURE(),LONG,DERIVED
OnLButtonDown           PROCEDURE(),LONG,DERIVED
                     END
  #ENDIF
  #IF(%IncludeDrop)
#SET(%ValueConstruct,'OleDropTarget' & %ActiveTemplateInstance)
%20ValueConstruct CLASS(IDropTargetClass)
DropEffect              PROCEDURE(DWORD grfKeyState, *POINT pt, DWORD dwAllowed),DWORD,DERIVED
QueryDataObject         PROCEDURE(*IDataObject pDataObject),BOOL,DERIVED
DragEnter               PROCEDURE(*IDataObject pDataObject, DWORD grfKeyState, POINT pt, *DWORD pdwEffect),HRESULT,PROC,DERIVED
DragOver                PROCEDURE(DWORD grfKeyState, POINT pt, *DWORD pdwEffect),HRESULT,PROC,DERIVED
DragLeave               PROCEDURE(),HRESULT,PROC,DERIVED
Drop                    PROCEDURE(*IDataObject pDataObject, DWORD grfKeyState, POINT pt, *DWORD pdwEffect),HRESULT,PROC,DERIVED
#SET(%EmbedNumber,1)
#EMBED(%OleDropTargetClassDeclaration,'Ole Drag and Drop: DropTarget Class Declaration'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Class Declaration)')
                     END

  #ENDIF
#ENDAT
#!
#!
#AT(%LocalDataAfterClasses),WHERE(NOT %DoNotGenerate AND NOT %GenerateLegacy)
  #IF(%ActiveTemplateInstance = %FirstInstance)
!------------------------------------------------------------------------------
OleIniter            cOleIniter #<!Ole Initialization

  #ENDIF
  #IF(%IncludeDrop)
%20DropTargetEventName EQUATE(%EventOleDrop)

  #ENDIF
  #IF(%IncludeDrag OR %IncludeDrop)
#SET(%ValueConstruct,'OleDataObject' & %ActiveTemplateInstance)
%20ValueConstruct CLASS(TextDataObjectClass)
Init                    PROCEDURE(*CSTRING szText),HRESULT,PROC,DERIVED
Kill                    PROCEDURE(),HRESULT,PROC,DERIVED
_ShowErrorMessage       PROCEDURE(STRING pMethodName,HRESULT pHR,BYTE pUseDebugMode=true),HRESULT,PROC,DERIVED
QueryInterface          PROCEDURE(LONG riid, *LONG pvObject),HRESULT,DERIVED
AddRef                  PROCEDURE(),LONG,PROC,DERIVED
Release                 PROCEDURE(),LONG,PROC,DERIVED
GetData                 PROCEDURE(*FORMATETC pformatetcIn,*STGMEDIUM pRemoteMedium),HRESULT,PROC,DERIVED
GetDataHere             PROCEDURE(*FORMATETC pformatetc,*STGMEDIUM pRemoteMedium),HRESULT,PROC,DERIVED
QueryGetData            PROCEDURE(*FORMATETC pformatetc),HRESULT,PROC,DERIVED
GetCanonicalFormatEtc   PROCEDURE(*FORMATETC pformatectIn,*FORMATETC pformatetcOut),HRESULT,PROC,DERIVED
SetData                 PROCEDURE(*FORMATETC pformatetc,*STGMEDIUM pmedium,LONG fRelease),HRESULT,PROC,DERIVED
EnumFormatEtc           PROCEDURE(ULONG dwDirection,LONG ppenumFormatEtc),HRESULT,PROC,DERIVED
DAdvise                 PROCEDURE(*FORMATETC pformatetc,ULONG advf,*IAdviseSink pAdvSink,*ULONG pdwConnection),HRESULT,PROC,DERIVED
DUnadvise               PROCEDURE(ULONG dwConnection),HRESULT,PROC,DERIVED
EnumDAdvise             PROCEDURE(LONG ppenumAdvise),HRESULT,PROC,DERIVED
                     END

  #ENDIF
  #IF(%IncludeDrag)
#SET(%ValueConstruct,'OleDataSource' & %ActiveTemplateInstance)
%20ValueConstruct CLASS(OleDataSourceClass)
GetOleData              PROCEDURE(),LONG,DERIVED
OnLButtonDown           PROCEDURE(),LONG,DERIVED
                     END

  #ENDIF
  #IF(%IncludeDrop)
#SET(%ValueConstruct,'OleDropTarget' & %ActiveTemplateInstance)
%20ValueConstruct CLASS(IDropTargetClass)
DropEffect              PROCEDURE(DWORD grfKeyState, *POINT pt, DWORD dwAllowed),DWORD,DERIVED
QueryDataObject         PROCEDURE(*IDataObject pDataObject),BOOL,DERIVED
DragEnter               PROCEDURE(*IDataObject pDataObject, DWORD grfKeyState, POINT pt, *DWORD pdwEffect),HRESULT,PROC,DERIVED
DragOver                PROCEDURE(DWORD grfKeyState, POINT pt, *DWORD pdwEffect),HRESULT,PROC,DERIVED
DragLeave               PROCEDURE(),HRESULT,PROC,DERIVED
Drop                    PROCEDURE(*IDataObject pDataObject, DWORD grfKeyState, POINT pt, *DWORD pdwEffect),HRESULT,PROC,DERIVED
#SET(%EmbedNumber,1)
#EMBED(%OleDropTargetClassDeclaration,'Ole Drag and Drop: DropTarget Class Declaration'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Class Declaration)')
                     END
  #ENDIF
#ENDAT
#!
#!
#AT(%WindowEventHandling,'OpenWindow'),PRIORITY(5000),DESCRIPTION('Ole Drag and Drop'),WHERE((NOT %DoNotGenerate) AND %IncludeDrop)
#EMBED(%OleDropTargetInitBeforeInit,'Ole Drag and Drop Before DropTarget.Init'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%DropTarget & '.Init','Before ' & %DropTarget & '.Init')
#FIX(%Control,%DragDropControl)
  #IF(%GenerateLegacy)
_hr = %DropTarget.Init(%DragDropControl{PROP:Handle},%EventOleDrop)
  #ELSE
hr = %DropTarget.Init(%DragDropControl{PROP:Handle},%EventOleDrop)
  #ENDIF
#SUSPEND
  #IF(%GenerateLegacy)
#?IF _hr = S_OK
  #ELSE
#?IF hr = S_OK
  #ENDIF
     #EMBED(%OleDropTargetInitSuccess,'Ole Drag and Drop DropTarget.Init succeeded'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%DropTarget & '.Init',%DropTarget & '.Init (success)')
#?ELSE
     #EMBED(%OleDropTargetInitFail,'Ole Drag and Drop DropTarget.Init failed'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%DropTarget & '.Init',%DropTarget & '.Init (fail)')
#?END
#RESUME
#ENDAT
#!
#!
#AT(%WindowEventHandling,'CloseWindow'),PRIORITY(5000),DESCRIPTION('Ole Drag and Drop'),WHERE(NOT %DoNotGenerate)
  #IF(%IncludeDrag)
kcr_RemoveProp(%DragDropControl{PROP:Handle},szPropFieldEquate)
kcr_RemoveProp(%DragDropControl{PROP:Handle},szPropDragData)
  #ENDIF
  #IF(%IncludeDrag OR %IncludeDrop)
kcr_RemoveProp(%DragDropControl{PROP:Handle},szPropDataObject)
  #ENDIF
#!
  #IF(%IncludeDrop)
kcr_RemoveProp(%DragDropControl{PROP:Handle},szPropDropTarget)
  #ENDIF
#ENDAT
#!
#!
#AT(%WindowManagerMethodCodeSection,'Init','(),BYTE'),PRIORITY(9950),DESCRIPTION('Ole Drag and Drop'),WHERE(NOT %DoNotGenerate AND NOT %GenerateLegacy)
#IF(%IncludeDrag)
%DragDropControl{PROP:DRAGID} = ''
  #SET(%ValueConstruct,'OleDataSource' & %ActiveTemplateInstance)
kcr_SetProp(%DragDropControl{PROP:Handle},szPropDragData,ADDRESS(%ValueConstruct))
kcr_SetProp(%DragDropControl{PROP:Handle},szPropFieldEquate,%DragDropControl)
REGISTER(EVENT:MOUSEDOWN,ADDRESS(%ValueConstruct.OnLButtonDown),ADDRESS(%ValueConstruct),%Window,%DragDropControl)
#ENDIF
#IF(%IncludeDrag OR %IncludeDrop)
  #SET(%ValueConstruct,'OleDataObject' & %ActiveTemplateInstance)
kcr_SetProp(%DragDropControl{PROP:Handle},szPropDataObject,ADDRESS(%ValueConstruct))
#ENDIF
#!
#IF(%IncludeDrop)
  #SET(%ValueConstruct,'OleDropTarget' & %ActiveTemplateInstance)
kcr_SetProp(%DragDropControl{PROP:Handle},szPropDropTarget,ADDRESS(%ValueConstruct))
#ENDIF
#!
#IF(%IncludeDrag)
  #SET(%ValueConstruct,SUB(%DragDropControl,2,LEN(%DragDropControl)-1) & ':origWndProc' & %ActiveTemplateInstance)
%ValueConstruct = %DragDropControl{PROP:WndProc}
  #SET(%ValueConstruct,SUB(%DragDropControl,2,LEN(%DragDropControl)-1) & ':WndProc' & %ActiveTemplateInstance)
%DragDropControl{PROP:WndProc} = ADDRESS(%ValueConstruct)

#ENDIF
#ENDAT
#!
#!
#AT(%AfterWindowOpening),PRIORITY(9950),DESCRIPTION('Ole Drag and Drop Sub-Class WndProc'),WHERE(NOT %DoNotGenerate AND %GenerateLegacy)
#IF(%IncludeDrag)
%DragDropControl{PROP:DRAGID} = ''
  #SET(%ValueConstruct,'OleDataSource' & %ActiveTemplateInstance)
kcr_SetProp(%DragDropControl{PROP:Handle},szPropDragData,ADDRESS(%ValueConstruct))
kcr_SetProp(%DragDropControl{PROP:Handle},szPropFieldEquate,%DragDropControl)
REGISTER(EVENT:MOUSEDOWN,ADDRESS(%ValueConstruct.OnLButtonDown),ADDRESS(%ValueConstruct),%Window,%DragDropControl)
#ENDIF
#IF(%IncludeDrag OR %IncludeDrop)
  #SET(%ValueConstruct,'OleDataObject' & %ActiveTemplateInstance)
kcr_SetProp(%DragDropControl{PROP:Handle},szPropDataObject,ADDRESS(%ValueConstruct))
#ENDIF
#!
#IF(%IncludeDrop)
  #SET(%ValueConstruct,'OleDropTarget' & %ActiveTemplateInstance)
kcr_SetProp(%DragDropControl{PROP:Handle},szPropDropTarget,ADDRESS(%ValueConstruct))
#ENDIF
#!
#IF(%IncludeDrag)
  #SET(%ValueConstruct,SUB(%DragDropControl,2,LEN(%DragDropControl)-1) & ':origWndProc' & %ActiveTemplateInstance)
%ValueConstruct = %DragDropControl{PROP:WndProc}
  #SET(%ValueConstruct,SUB(%DragDropControl,2,LEN(%DragDropControl)-1) & ':WndProc' & %ActiveTemplateInstance)
%DragDropControl{PROP:WndProc} = ADDRESS(%ValueConstruct)

#ENDIF
#ENDAT
#!
#!
#AT(%DataSectionBeforeWindow),PRIORITY(200),DESCRIPTION('Ole Drag and Drop'),WHERE(NOT %DoNotGenerate AND %ActiveTemplateInstance = %FirstInstance AND %GenerateLegacy)
_hr                  HRESULT
#ENDAT
#!
#!
#AT(%DataSectionBeforeWindow),PRIORITY(200),DESCRIPTION('Ole Drag and Drop'),WHERE((NOT %DoNotGenerate) AND %IncludeDrop AND %GenerateLegacy AND %ActiveTemplateInstance = %FirstInstance)
OleCursorPos         LONG,AUTO
#ENDAT
#!
#!
#AT(%WindowManagerMethodDataSection,'TakeWindowEvent','(),BYTE'),PRIORITY(200),DESCRIPTION('Ole Drag and Drop'),WHERE(NOT %DoNotGenerate AND %ActiveTemplateInstance = %FirstInstance AND NOT %GenerateLegacy)
hr                   HRESULT
#ENDAT
#!
#!
#AT(%WindowManagerMethodDataSection,'TakeEvent','(),BYTE'),PRIORITY(200),DESCRIPTION('Ole Drag and Drop'),WHERE((NOT %DoNotGenerate) AND %IncludeDrop AND %ActiveTemplateInstance = %FirstInstance AND NOT %GenerateLegacy)
OleCursorPos         LONG,AUTO
#ENDAT
#!
#!
#AT(%AcceptLoopBeforeEventHandling),WHERE((NOT %DoNotGenerate) AND %GenerateLegacy AND %IncludeDrop)
  IF EVENT() = %DropTargetEventName
     #EMBED(%OnOleDrop,'Ole Drag and Drop on EVENT:OleDrop'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,'Window Events',%DropTargetEventName & ' (After Generated Code)')
  #IF(%GenerateDropDefault)
    #FIX(%Control,%DragDropControl)
    #CASE(%ControlType)
      #OF('TEXT')
      #OROF('ENTRY')
        #CASE(%DefaultDropBehaviour)
        #OF('Append')
     %ControlUse = CLIP(%ControlUse) & %OleDataBufferOverride
     DISPLAY(%DragDropControl)
     OleCursorPos = LEN(%ControlUse)
     %Control{PROP:Selected} = OleCursorPos
     %Control{PROP:SelEnd} = 0
        #OF('Replace All')
     %ControlUse = %OleDataBufferOverride
     DISPLAY(%DragDropControl)
     OleCursorPos = LEN(%ControlUse)
     %Control{PROP:Selected} = 1
     %Control{PROP:SelEnd} = OleCursorPos
        #OF('Replace Selection')
     IF %Control{PROP:SelEnd} > 0
        OleCursorPos = %Control{PROP:SelStart}
        %ControlUse = %ControlUse[1 : %Control{PROP:SelStart} - 1] & %OleDataBufferOverride & %ControlUse[%Control{PROP:SelEnd} + 1 : LEN(%ControlUse)]
        DISPLAY(%DragDropControl)
        %Control{PROP:Selected} = OleCursorPos
        %Control{PROP:SelEnd} = OleCursorPos + LEN(%OleDataBufferOverride) - 1
     ELSE
        %ControlUse = CLIP(%ControlUse) & %OleDataBufferOverride
        DISPLAY(%DragDropControl)
        %Control{PROP:Selected} = 1
        %Control{PROP:SelEnd} = LEN(%ControlUse)
     END
        #ENDCASE
    #ENDCASE
  #ENDIF
     SELECT(%DragDropControl)
  END
#ENDAT
#!
#!
#AT(%WindowManagerMethodCodeSection,'TakeEvent','(),BYTE'),PRIORITY(6300),DESCRIPTION('Ole Drag and Drop'),WHERE((NOT %DoNotGenerate) AND %IncludeDrop AND NOT %GenerateLegacy)
  IF EVENT() = %DropTargetEventName
     #EMBED(%OnOleDrop,'Ole Drag and Drop on EVENT:OleDrop'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,'Window Events',%DropTargetEventName & ' (After Generated Code)')
  #IF(%GenerateDropDefault)
    #FIX(%Control,%DragDropControl)
    #CASE(%ControlType)
      #OF('TEXT')
      #OROF('ENTRY')
        #CASE(%DefaultDropBehaviour)
        #OF('Append')
     %ControlUse = CLIP(%ControlUse) & %OleDataBufferOverride
     DISPLAY(%DragDropControl)
     OleCursorPos = LEN(%ControlUse)
     %Control{PROP:Selected} = OleCursorPos
     %Control{PROP:SelEnd} = 0
        #OF('Replace All')
     %ControlUse = %OleDataBufferOverride
     DISPLAY(%DragDropControl)
     OleCursorPos = LEN(%ControlUse)
     %Control{PROP:Selected} = 1
     %Control{PROP:SelEnd} = OleCursorPos
        #OF('Replace Selection')
     IF %Control{PROP:SelEnd} > 0
        OleCursorPos = %Control{PROP:SelStart}
        %ControlUse = %ControlUse[1 : %Control{PROP:SelStart} - 1] & %OleDataBufferOverride & %ControlUse[%Control{PROP:SelEnd} + 1 : LEN(%ControlUse)]
        DISPLAY(%DragDropControl)
        %Control{PROP:Selected} = OleCursorPos
        %Control{PROP:SelEnd} = OleCursorPos + LEN(%OleDataBufferOverride) - 1
     ELSE
        %ControlUse = CLIP(%ControlUse) & %OleDataBufferOverride
        DISPLAY(%DragDropControl)
        %Control{PROP:Selected} = 1
        %Control{PROP:SelEnd} = LEN(%ControlUse)
     END
        #ENDCASE
    #ENDCASE
     SELECT(%DragDropControl)
  #ENDIF
  END
#ENDAT
#!
#!
#AT(%LocalProcedures),WHERE(NOT %DoNotGenerate),PRIORITY(7800)
  #IF(%IncludeDrag OR %IncludeDrop)
#SET(%ValueConstruct,%DataObject & '.Init')
%20ValueConstruct PROCEDURE(*CSTRING szText) !,HRESULT,PROC,VIRTUAL
hr              HRESULT(E_FAIL)
fmtetc          LIKE(FORMATETC)
stgmed          LIKE(STGMEDIUM)
pMem            LONG
#SET(%EmbedNumber,1)
#EMBED(%OleDataObjectInitDataSection,'Ole Drag and Drop:  OleDataObjectInit Data Section'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Data Section)')

  CODE
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataObjectInitCodeStart,'Ole Drag and Drop:  OleDataObjectInit Code Start'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code Start)')

    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataObjectInitCodeBeforeParent,'Ole Drag and Drop:  OleDataObjectInit Code Before Parent Call'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code Before Parent Call)')
    PARENT.Init()
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataObjectInitCodeAfterParent,'Ole Drag and Drop:  OleDataObjectInit Code After Parent Call'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code Before After Call)')
  #IF(%GenerateDragDefault)
    kcr_ZeroMemory(ADDRESS(fmtetc),SIZE(FORMATETC))
    fmtetc.cfFormat  = CF_TEXT
    fmtetc.ptd       = 0
    fmtetc.dwAspect  = DVASPECT_CONTENT
    fmtetc.lindex    = -1
    fmtetc.tymed     = TYMED_HGLOBAL

    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataObjectInitCodeAfter,'Ole Drag and Drop:  OleDataObjectInit Code After Setting FORMATETC'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code After Setting FORMATETC)')

    kcr_ZeroMemory(ADDRESS(stgmed),SIZE(STGMEDIUM))
    stgmed.tymed     = TYMED_HGLOBAL
    stgmed.u.hGlobal = kcr_GlobalAlloc(GHND, LEN(szText)+1)
    pMem = kcr_GlobalLock(stgmed.u.hGlobal)
    kcr_memcpy(pMem, ADDRESS(szText), LEN(szText))
    kcr_GlobalUnlock(stgmed.u.hGlobal)
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataObjectInitCodeBeforeInit,'Ole Drag and Drop:  OleDataObjectInit Code Before Init call'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code Before Init Call)')
    hr = SELF.SetData(fmtetc,stgmed)
  #ENDIF
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataObjectInitCodeEnd,'Ole Drag and Drop:  OleDataObjectInit Code End'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code End)')
    RETURN hr
!------------------------------------------------------------------------------
#SET(%ValueConstruct,%DataObject & '.Kill')
%20ValueConstruct PROCEDURE() !,HRESULT,PROC,DERIVED
hr           HRESULT(E_FAIL)
#SET(%EmbedNumber,1)
#EMBED(%OleDataObjectKillDataSection,'Ole Drag and Drop:  OleDataObjectKill Data Section'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Data Section)')

  CODE
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataObjectKillCodeStart,'Ole Drag and Drop:  OleDataObjectKill Code Start'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code Start)')
    hr = PARENT.Kill()
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataObjectKillCodeAfterParentCall,'Ole Drag and Drop:  OleDataObjectKill Code After Parent Call'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code After Parent Call)')
    RETURN hr
    #EMBED(%OleDataObjectKillCodeEnd,'Ole Drag and Drop:  OleDataObjectKill Code End'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code End)')
!------------------------------------------------------------------------------
#SET(%ValueConstruct,%DataObject & '._ShowErrorMessage')
%20ValueConstruct PROCEDURE(STRING pMethodName,HRESULT pHR,BYTE pUseDebugMode=true) !,HRESULT,PROC,DERIVED
hr           HRESULT(E_FAIL)
#SET(%EmbedNumber,1)
#EMBED(%OleDataObjectShowErrorMessageDataSection,'Ole Drag and Drop:  OleDataObjectShowErrorMessage Data Section'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Data Section)')

  CODE
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataObjectShowErrorMessageCodeStart,'Ole Drag and Drop:  OleDataObjectShowErrorMessage Code Start'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code Start)')
    hr = PARENT._ShowErrorMessage(pMethodName, pHR, pUseDebugMode)
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataObjectShowErrorMessageCodeAfterParentCall,'Ole Drag and Drop:  OleDataObjectShowErrorMessage Code After Parent Call'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code After Parent Call)')
    RETURN hr
    #EMBED(%OleDataObjectShowErrorMessageCodeEnd,'Ole Drag and Drop:  OleDataObjectShowErrorMessage Code End'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code End)')
!------------------------------------------------------------------------------
#SET(%ValueConstruct,%DataObject & '.QueryInterface')
%20ValueConstruct PROCEDURE(LONG riid, *LONG pvObject) !,HRESULT,DERIVED
hr           HRESULT(E_NOINTERFACE)
#SET(%EmbedNumber,1)
#EMBED(%OleDataObjectQueryInterfaceDataSection,'Ole Drag and Drop:  OleDataObjectQueryInterface Data Section'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Data Section)')

  CODE
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataObjectQueryInterfaceCodeStart,'Ole Drag and Drop:  OleDataObjectQueryInterface Code Start'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code Start)')
    hr = PARENT.QueryInterface(riid, pvObject)
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataObjectQueryInterfaceCodeAfterParentCall,'Ole Drag and Drop:  OleDataObjectQueryInterface Code After Parent Call'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code After Parent Call)')
    RETURN hr
    #EMBED(%OleDataObjectQueryInterfaceCodeEnd,'Ole Drag and Drop:  OleDataObjectQueryInterface Code End'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code End)')
!------------------------------------------------------------------------------
#SET(%ValueConstruct,%DataObject & '.AddRef')
%20ValueConstruct PROCEDURE()   !,LONG,PROC,DERIVED
count        LONG
#SET(%EmbedNumber,1)
#EMBED(%OleDataObjectAddRefDataSection,'Ole Drag and Drop:  OleDataObjectAddRef Data Section'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Data Section)')

  CODE
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataObjectAddRefCodeStart,'Ole Drag and Drop:  OleDataObjectAddRef Code Start'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code Start)')
    count = PARENT.AddRef()
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataObjectAddRefCodeAfterParentCall,'Ole Drag and Drop:  OleDataObjectAddRef Code After Parent Call'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code After Parent Call)')
    RETURN count
    #EMBED(%OleDataObjectAddRefCodeEnd,'Ole Drag and Drop:  OleDataObjectAddRef Code End'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code End)')
!------------------------------------------------------------------------------
#SET(%ValueConstruct,%DataObject & '.Release')
%20ValueConstruct PROCEDURE()   !,LONG,PROC,DERIVED
count        LONG
#SET(%EmbedNumber,1)
#EMBED(%OleDataObjectReleaseDataSection,'Ole Drag and Drop:  OleDataObjectRelease Data Section'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Data Section)')

  CODE
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataObjectReleaseCodeStart,'Ole Drag and Drop:  OleDataObjectRelease Code Start'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code Start)')
    count = PARENT.Release()
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataObjectReleaseCodeAfterParentCall,'Ole Drag and Drop:  OleDataObjectRelease Code After Parent Call'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code After Parent Call)')
    RETURN count
    #EMBED(%OleDataObjectReleaseCodeEnd,'Ole Drag and Drop:  OleDataObjectRelease Code End'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code End)')
!------------------------------------------------------------------------------
#SET(%ValueConstruct,%DataObject & '.GetData')
%20ValueConstruct PROCEDURE(*FORMATETC pformatetcIn,*STGMEDIUM pRemoteMedium) !,HRESULT,PROC,DERIVED
hr           HRESULT(E_FAIL)
#SET(%EmbedNumber,1)
#EMBED(%OleDataObjectGetDataDataSection,'Ole Drag and Drop:  OleDataObjectGetData Data Section'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Data Section)')

  CODE
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataObjectGetDataCodeStart,'Ole Drag and Drop:  OleDataObjectGetData Code Start'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code Start)')
    hr = PARENT.GetData(pformatetcIn, pRemoteMedium)
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataObjectGetDataCodeAfterParentCall,'Ole Drag and Drop:  OleDataObjectGetData Code After Parent Call'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code After Parent Call)')
    RETURN hr
    #EMBED(%OleDataObjectGetDataCodeEnd,'Ole Drag and Drop:  OleDataObjectGetData Code End'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code End)')
!------------------------------------------------------------------------------
#SET(%ValueConstruct,%DataObject & '.GetDataHere')
%20ValueConstruct PROCEDURE(*FORMATETC pformatetc,*STGMEDIUM pRemoteMedium) !,HRESULT,PROC,DERIVED
hr           HRESULT(E_FAIL)
#SET(%EmbedNumber,1)
#EMBED(%OleDataObjectGetDataHereDataSection,'Ole Drag and Drop:  OleDataObjectGetDataHere Data Section'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Data Section)')

  CODE
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataObjectGetDataHereCodeStart,'Ole Drag and Drop:  OleDataObjectGetDataHere Code Start'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code Start)')
    hr = PARENT.GetDataHere(pformatetc, pRemoteMedium)
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataObjectGetDataHereCodeAfterParentCall,'Ole Drag and Drop:  OleDataObjectGetDataHere Code After Parent Call'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code After Parent Call)')
    RETURN hr
    #EMBED(%OleDataObjectGetDataHereCodeEnd,'Ole Drag and Drop:  OleDataObjectGetDataHere Code End'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code End)')
!------------------------------------------------------------------------------
#SET(%ValueConstruct,%DataObject & '.QueryGetData')
%20ValueConstruct PROCEDURE(*FORMATETC pformatetc) !,HRESULT,PROC,DERIVED
hr           HRESULT(E_FAIL)
#SET(%EmbedNumber,1)
#EMBED(%OleDataObjectQueryGetDataDataSection,'Ole Drag and Drop:  OleDataObjectQueryGetData Data Section'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Data Section)')

  CODE
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataObjectQueryGetDataCodeStart,'Ole Drag and Drop:  OleDataObjectQueryGetData Code Start'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code Start)')
    hr = PARENT.QueryGetData(pformatetc)
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataObjectQueryGetDataCodeAfterParentCall,'Ole Drag and Drop:  OleDataObjectQueryGetData Code After Parent Call'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code After Parent Call)')
    RETURN hr
    #EMBED(%OleDataObjectQueryGetDataCodeEnd,'Ole Drag and Drop:  OleDataObjectQueryGetData Code End'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code End)')
!------------------------------------------------------------------------------
#SET(%ValueConstruct,%DataObject & '.GetCanonicalFormatEtc')
%20ValueConstruct PROCEDURE(*FORMATETC pformatectIn,*FORMATETC pformatetcOut) !,HRESULT,PROC,DERIVED
hr           HRESULT(E_FAIL)
#SET(%EmbedNumber,1)
#EMBED(%OleDataObjectGetCanonicalFormatEtcDataSection,'Ole Drag and Drop:  OleDataObjectGetCanonicalFormatEtc Data Section'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Data Section)')

  CODE
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataObjectGetCanonicalFormatEtcCodeStart,'Ole Drag and Drop:  OleDataObjectGetCanonicalFormatEtc Code Start'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code Start)')
    hr = PARENT.GetCanonicalFormatEtc(pformatectIn,pformatetcOut)
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataObjectGetCanonicalFormatEtcCodeAfterParentCall,'Ole Drag and Drop:  OleDataObjectGetCanonicalFormatEtc Code After Parent Call'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code After Parent Call)')
    RETURN hr
    #EMBED(%OleDataObjectGetCanonicalFormatEtcCodeEnd,'Ole Drag and Drop:  OleDataObjectGetCanonicalFormatEtc Code End'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code End)')
!------------------------------------------------------------------------------
#SET(%ValueConstruct,%DataObject & '.SetData')
%20ValueConstruct PROCEDURE(*FORMATETC pformatetc,*STGMEDIUM pmedium,LONG fRelease) !,HRESULT,PROC,DERIVED
hr           HRESULT(E_FAIL)
#SET(%EmbedNumber,1)
#EMBED(%OleDataObjectSetDataDataSection,'Ole Drag and Drop:  OleDataObjectSetData Data Section'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Data Section)')

  CODE
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataObjectSetDataCodeStart,'Ole Drag and Drop:  OleDataObjectSetData Code Start'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code Start)')
    hr = PARENT.SetData(pformatetc, pmedium, fRelease)
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataObjectSetDataCodeAfterParentCall,'Ole Drag and Drop:  OleDataObjectSetData Code After Parent Call'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code After Parent Call)')
    RETURN hr
    #EMBED(%OleDataObjectSetDataCodeEnd,'Ole Drag and Drop:  OleDataObjectSetData Code End'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code End)')
!------------------------------------------------------------------------------
#SET(%ValueConstruct,%DataObject & '.EnumFormatEtc')
%20ValueConstruct PROCEDURE(ULONG dwDirection,LONG ppenumFormatEtc) !,HRESULT,PROC,DERIVED
hr           HRESULT(E_FAIL)
#SET(%EmbedNumber,1)
#EMBED(%OleDataObjectEnumFormatEtcDataSection,'Ole Drag and Drop:  OleDataObjectEnumFormatEtc Data Section'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Data Section)')

  CODE
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataObjectEnumFormatEtcCodeStart,'Ole Drag and Drop:  OleDataObjectEnumFormatEtc Code Start'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code Start)')
    hr = PARENT.EnumFormatEtc(dwDirection,ppenumFormatEtc)
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataObjectEnumFormatEtcCodeAfterParentCall,'Ole Drag and Drop:  OleDataObjectEnumFormatEtc Code After Parent Call'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code After Parent Call)')
    RETURN hr
    #EMBED(%OleDataObjectEnumFormatEtcCodeEnd,'Ole Drag and Drop:  OleDataObjectEnumFormatEtc Code End'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code End)')
!------------------------------------------------------------------------------
#SET(%ValueConstruct,%DataObject & '.DAdvise')
%20ValueConstruct PROCEDURE(*FORMATETC pformatetc,ULONG advf,*IAdviseSink pAdvSink,*ULONG pdwConnection) !,HRESULT,PROC,DERIVED
hr           HRESULT(E_FAIL)
#SET(%EmbedNumber,1)
#EMBED(%OleDataObjectDAdviseDataSection,'Ole Drag and Drop:  OleDataObjectDAdvise Data Section'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Data Section)')

  CODE
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataObjectDAdviseCodeStart,'Ole Drag and Drop:  OleDataObjectDAdvise Code Start'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code Start)')
    hr = PARENT.DAdvise(pformatetc, advf, pAdvSink, pdwConnection)
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataObjectDAdviseCodeAfterParentCall,'Ole Drag and Drop:  OleDataObjectDAdvise Code After Parent Call'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code After Parent Call)')
    RETURN hr
    #EMBED(%OleDataObjectDAdviseCodeEnd,'Ole Drag and Drop:  OleDataObjectDAdvise Code End'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code End)')
!------------------------------------------------------------------------------
#SET(%ValueConstruct,%DataObject & '.DUnadvise')
%20ValueConstruct PROCEDURE(ULONG dwConnection) !,HRESULT,PROC,DERIVED
hr           HRESULT(E_FAIL)
#SET(%EmbedNumber,1)
#EMBED(%OleDataObjectDUnadviseDataSection,'Ole Drag and Drop:  OleDataObjectDUnadvise Data Section'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Data Section)')

  CODE
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataObjectDUnadviseCodeStart,'Ole Drag and Drop:  OleDataObjectDUnadvise Code Start'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code Start)')
    hr = PARENT.DUnadvise(dwConnection)
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataObjectDUnadviseCodeAfterParentCall,'Ole Drag and Drop:  OleDataObjectDUnadvise Code After Parent Call'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code After Parent Call)')
    RETURN hr
    #EMBED(%OleDataObjectDUnadviseCodeEnd,'Ole Drag and Drop:  OleDataObjectDUnadvise Code End'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code End)')
!------------------------------------------------------------------------------
#SET(%ValueConstruct,%DataObject & '.EnumDAdvise')
%20ValueConstruct PROCEDURE(LONG ppenumAdvise) !,HRESULT,PROC,DERIVED
hr           HRESULT(E_FAIL)
#SET(%EmbedNumber,1)
#EMBED(%OleDataObjectEnumDAdviseDataSection,'Ole Drag and Drop:  OleDataObjectEnumDAdvise Data Section'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Data Section)')

  CODE
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataObjectEnumDAdviseCodeStart,'Ole Drag and Drop:  OleDataObjectEnumDAdvise Code Start'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code Start)')
    hr = PARENT.EnumDAdvise(ppenumAdvise)
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataObjectEnumDAdviseCodeAfterParentCall,'Ole Drag and Drop:  OleDataObjectEnumDAdvise Code After Parent Call'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code After Parent Call)')
    RETURN hr
    #EMBED(%OleDataObjectEnumDAdviseCodeEnd,'Ole Drag and Drop:  OleDataObjectEnumDAdvise Code End'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code End)')
!------------------------------------------------------------------------------
  #ENDIF
  #IF(%IncludeDrop)
#!
#!
#!
#!
#FIX(%Control,%DragDropControl)
#SET(%ValueConstruct,%DropTarget & '.DropEffect')
%20ValueConstruct PROCEDURE(DWORD grfKeyState, *POINT pt, DWORD dwAllowed)
dwDropEffect        DWORD(DROPEFFECT_NONE)
#SET(%EmbedNumber,1)
#EMBED(%OleDropTargetDropEffectData,'Ole Drag and Drop: DropTarget.DropEffect Data Section'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Data Section)')
  CODE
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDropTargetDropEffectCodeBeforeParent,'Ole Drag and Drop: DropTarget.DropEffect Code Before Parent Call'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Before Parent Call)')
    dwDropEffect = PARENT.DropEffect(grfKeyState, pt, dwAllowed)
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDropTargetDropEffectCodeAfterParent,'Ole Drag and Drop: DropTarget.DropEffect Code After Parent Call'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (After Parent Call)')
    RETURN dwDropEffect
!------------------------------------------------------------------------------

#SET(%ValueConstruct,%DropTarget & '.QueryDataObject')
%20ValueConstruct PROCEDURE(*IDataObject pDataObject)
bResult         BOOL
#SET(%EmbedNumber,1)
#EMBED(%OleDropTargetQueryDataObjectData,'Ole Drag and Drop: DropTarget.QueryDataObject Data Section'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Data Section)')

  CODE
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDropTargetQueryDataObjectCodeBeforeParent,'Ole Drag and Drop: DropTarget.QueryDataObject Code Before Parent Call'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Before Parent Call)')
    bResult = PARENT.QueryDataObject(pDataObject)
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDropTargetQueryDataObjectCodeAfterParent,'Ole Drag and Drop: DropTarget.QueryDataObject Code After Parent Call'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (After Parent Call)')
    RETURN bResult
!------------------------------------------------------------------------------

#SET(%ValueConstruct,%DropTarget & '.DragEnter')
%20ValueConstruct PROCEDURE(*IDataObject pDataObject, DWORD grfKeyState, POINT pt, *DWORD pdwEffect)
hr              HRESULT
#SET(%EmbedNumber,1)
#EMBED(%OleDropTargetDragEnterData,'Ole Drag and Drop: DropTarget.DragEnter Data Section'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Data Section)')

  CODE
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDropTargetDragEnterCodeBeforeParent,'Ole Drag and Drop: DropTarget.DragEnter Code Before Parent Call'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Before Parent Call)')
    hr = PARENT.DragEnter(pDataObject,grfKeyState,pt,pdwEffect)
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDropTargetDragEnterCodeAfterParent,'Ole Drag and Drop: DropTarget.DragEnter Code After Parent Call'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (After Parent Call)')
    RETURN hr
!------------------------------------------------------------------------------

#SET(%ValueConstruct,%DropTarget & '.DragOver')
%20ValueConstruct PROCEDURE(DWORD grfKeyState, POINT pt, *DWORD pdwEffect)
hr              HRESULT
#SET(%EmbedNumber,1)
#EMBED(%OleDropTargetDragOverData,'Ole Drag and Drop: DropTarget.DragOver Data Section'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Data Section)')

  CODE
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDropTargetDragOverCodeBeforeParent,'Ole Drag and Drop: DropTarget.DragOver Code Before Parent Call'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Before Parent Call)')
    hr = PARENT.DragOver(grfKeyState,pt,pdwEffect)
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDropTargetDragOverCodeAfterParent,'Ole Drag and Drop: DropTarget.DragOver Code After Parent Call'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (After Parent Call)')
    RETURN hr
!------------------------------------------------------------------------------

#SET(%ValueConstruct,%DropTarget & '.DragLeave')
%20ValueConstruct PROCEDURE()
hr              HRESULT
#SET(%EmbedNumber,1)
#EMBED(%OleDropTargetDragLeaveData,'Ole Drag and Drop: DropTarget.DragLeave Data Section'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Data Section)')

  CODE
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDropTargetDragLeaveCodeBeforeParent,'Ole Drag and Drop: DropTarget.DragLeave Code Before Parent Call'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Before Parent Call)')
    hr = PARENT.DragLeave()
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDropTargetDragLeaveCodeAfterParent,'Ole Drag and Drop: DropTarget.DragLeave Code After Parent Call'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (After Parent Call)')
    RETURN hr
!------------------------------------------------------------------------------

#SET(%ValueConstruct,%DropTarget & '.Drop')
%20ValueConstruct PROCEDURE(*IDataObject pDataObject, LONG grfKeyState, POINT pt, *LONG pdwEffect)
hr              HRESULT
pNewString      &STRING
fmtetc          &FORMATETC
stgmed          &STGMEDIUM
pData           LONG
szData          &CSTRING
iUnk            &IUnknown
#SET(%EmbedNumber,1)
#EMBED(%OleDropTargetDropData,'Ole Drag and Drop: DropTarget.Drop Data Section'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Data Section)')

  CODE
#FIX(%Control,%DragDropControl)
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDropTargetDropCodeBeforeParent,'Ole Drag and Drop: DropTarget.Drop Code Before Parent Call'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Before Parent Call)')
    hr = PARENT.Drop(pDataObject, grfKeyState, pt, pdwEffect)
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDropTargetDropCodeAfterParent,'Ole Drag and Drop: DropTarget.Drop Code After Parent Call'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (After Parent Call)')


    pNewString &= NEW(STRING(SIZE(FORMATETC)))
    fmtetc &= ADDRESS(pNewString)
    IF ~fmtetc &= NULL
       kcr_memset(ADDRESS(fmtetc),0,SIZE(FORMATETC))
       fmtetc.cfFormat = CF_TEXT
       fmtetc.ptd      = 0
       fmtetc.dwAspect = DVASPECT_CONTENT
       fmtetc.lindex   = -1
       fmtetc.tymed    = TYMED_HGLOBAL
       IF pDataObject.QueryGetData(ADDRESS(fmtetc)) = S_OK
          !ask the IDataObject for some CF_TEXT data, stored as a HGLOBAL
          pNewString &= NEW(STRING(SIZE(STGMEDIUM)))
          stgmed &= ADDRESS(pNewString)
          IF ~stgmed &= NULL
             kcr_ZeroMemory(ADDRESS(stgmed),SIZE(STGMEDIUM))
             hr = pDataObject.GetData(ADDRESS(fmtetc), ADDRESS(stgmed))
             IF hr = S_OK
                !We need to lock the HGLOBAL handle because we can't
                !be sure if this is GMEM_FIXED (i.e. normal heap) data or not
                pData = kcr_globalLock(stgmed.u.hGlobal)
                szData &= (pData)
                IF LEN(szData) >  %OleDataBufferOverrideSize
                   %OleDataBufferOverride = szData[1 : %OleDataBufferOverrideSize] & '<0>'
                ELSE
                   %OleDataBufferOverride = szData & '<0>'
                END
                !cleanup
                kcr_globalUnlock(stgmed.u.hGlobal)
                IF stgmed.pUnkForRelease
                   iUnk &= (stgmed.pUnkForRelease)
                   iUnk.Release()
                   iUnk &= NULL
                ELSE
                   kcr_ReleaseStgMedium(stgmed)
                END
                #SET(%EmbedNumber,%EmbedNumber+1)
                #EMBED(%OleDropTargetDropCodeSuccess,'Ole Drag and Drop: DropTarget.Drop Code Success'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (After Successful Drop)')
                DISPOSE(stgmed)
                stgmed &= NULL
                pNewString &= NULL
                POST(SELF.m_OleDropEvent)
             ELSE
                DISPOSE(stgmed)
                stgmed &= NULL
                pNewString &= NULL
             END
          END
          DISPOSE(fmtetc)
          fmtetc &= NULL
          pNewString &= NULL
       END
    END
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDropTargetDropCodeEnd,'Ole Drag and Drop: DropTarget.Drop Code End'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code End)')
    RETURN hr
!------------------------------------------------------------------------------

  #ENDIF
#!
#!
  #IF(%IncludeDrag)
#SET(%ValueConstruct,'OleDataSource' & %ActiveTemplateInstance & '.GetOleData')
%20ValueConstruct PROCEDURE()
!This class method is called from a subclassed winproc
!just prior to initiating the drag operation to allow
!you to complete preparation of the data for transfer.
!
!Default Behaviour:
#FIX(%Control,%DragDropControl)
#CASE(%ControlType)
  #OF('LIST')
!%OleDataBufferOverride will contain the contents of the currently selected queue entry.
!The fields will be separated by %FieldDelimiter.
  #OF('TEXT')
  #OROF('ENTRY')
!If the control has a selected portion, then that becomes the
!drag text otherwise the entire contents are transferred.
#ENDCASE
!
!WARNING:
!Make sure the address returned points to global, module, static, or thread data
!otherwise you will most certinly GPF.
!
#SET(%EmbedNumber,1)
#EMBED(%OleDataClassGetAddressDataSection,'Ole Drag and Drop:  Prepare Ole Drag Text Data Section'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Data Section)')

  CODE
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataClassGetAddressCodeBefore,'Ole Drag and Drop:  Prepare Ole Drag Text Code Start'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code Start)')
#IF(%GenerateDragDefault)
    %OleDataBufferOverride = ''
  #CASE(%ControlType)
  #OF('LIST')
    #FIX(%Control,%DragDropControl)
    GET(%ControlFrom, CHOICE(%Control))
      #FOR(%ControlField)
        #SET(%TempConstruct,%ControlFrom & '.' & CLIP(%ControlField))
        #IF(%IncludeColumnHeading)
    %OleDataBufferOverride = %OleDataBufferOverride & '%ControlFieldHeader: ' & %TempConstruct & %Delimiter
        #ELSE
    %OleDataBufferOverride = %OleDataBufferOverride & %TempConstruct & %Delimiter
        #ENDIF
      #ENDFOR
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataClassGetAddressCodeListAfter,'Ole Drag and Drop:  Prepare Ole Drag Text Code End'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code End)')
    RETURN ADDRESS(%OleDataBufferOverride)
  #OF('TEXT')
  #OROF('ENTRY')
    UPDATE(%DragDropControl)
    IF %DragDropControl{PROP:SelEnd} > 0
       %OleDataBufferOverride = %ControlUse[%DragDropControl{PROP:SelStart} : %DragDropControl{PROP:SelEnd}]
    ELSE
       %OleDataBufferOverride = %ControlUse
    END
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataClassGetAddressCodeTextAfter,'Ole Drag and Drop:  Prepare Ole Drag Text Code End'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code End)')
    RETURN ADDRESS(%OleDataBufferOverride)
  #ENDCASE
#ELSE
    #SET(%EmbedNumber,%EmbedNumber+1)
    #EMBED(%OleDataClassGetAddressCodeAfter,'Ole Drag and Drop:  Prepare Ole Drag Text Code End'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Code End)')
    RETURN ADDRESS(%OleDataBufferOverride)
#ENDIF
!------------------------------------------------------------------------------

#SET(%ValueConstruct,'OleDataSource' & %ActiveTemplateInstance & '.OnLButtonDown')
%20ValueConstruct PROCEDURE()
hwnd            HWND
pt1             LIKE(POINT)
pTextDataObject &TextDataObjectClass
pDragText       &CSTRING
pIDropSource    &IDropSourceClass
pIDataObject    &IDataObject
dwResult        DWORD
dwEffect        DWORD
dwOKEffect      DWORD
      #CASE(%ControlType)
      #OF('TEXT')
      #OROF('ENTRY')
selStart        LONG
      #ENDCASE
      #IF(%IncludeDrop)
hr              HRESULT
pOleDropTarget  &IDropTargetClass
      #ENDIF
#SET(%EmbedNumber,1)
#EMBED(%OleDataSourceOnLButtonDownDataSection,'Ole Data Source OnLButtonDown:  Data Section'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Data Section)')

    CODE
      #SET(%EmbedNumber,%EmbedNumber+1)
      #EMBED(%OleDataSourceOnLButtonDownBeginCodeSection,'Ole Data Source OnLButtonDown:  Start of Code'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Start of Code)')
#!
      #IF(%CopyEffect)
      dwOKEffect = BOR(dwOKEffect,DROPEFFECT_COPY)
      #ENDIF
      #IF(%MoveEffect)
      dwOKEffect = BOR(dwOKEffect,DROPEFFECT_MOVE)
      #ENDIF
      #IF(%LinkEffect)
      dwOKEffect = BOR(dwOKEffect,DROPEFFECT_LINK)
      #ENDIF
#!
      hwnd = %DragDropControl{PROP:Handle}
      #CASE(%ControlType)
      #OF('TEXT')
      #OROF('ENTRY')
      kcr_SendMessage(hwnd,WM_LBUTTONUP,0,0)
      selStart = %DragDropControl{PROP:SelStart}
      IF INRANGE(SelStart,SELF.m_SelStart,SELF.m_SelEnd)
         %DragDropControl{PROP:SelStart} = SELF.m_SelStart
         %DragDropControl{PROP:SelEnd}   = SELF.m_SelEnd
      END
      #ENDCASE
      kcr_GetCursorPos(pt1)
      IF kcr_DragDetect(hwnd,pt1.x,pt1.y)
         pDragText &= SELF.GetOleData()
         IF pDragText <> ''
      #IF(%IncludeDrop)
            pOleDropTarget &= (kcr_GetProp(hwnd,szPropDropTarget))
            IF ~pOleDropTarget &= NULL
               hr = pOleDropTarget.Kill()
            END
      #ENDIF
            pTextDataObject &= (kcr_GetProp(hwnd,szPropDataObject))
            IF ~pTextDataObject &= NULL
               pTextDataObject.AddRef()
               IF pTextDataObject.Init(pDragText) = S_OK
                  pIDataObject &= (pTextDataObject.GetInterfaceObject())
                  kcr_OleSetClipBoard(pIDataObject)
                  kcr_OleFlushClipboard()
                  pIDropSource &= NEW IDropSourceClass
                  pIDropSource.Init()
                  dwResult = kcr_DoDragDrop(pTextDataObject.GetInterfaceObject(),pIDropSource.GetInterfaceObject(), dwOKEffect, dwEffect)
                  IF dwResult = DRAGDROP_S_DROP
                     CASE dwEffect
                     #IF(%CopyEffect)
                     OF   DROPEFFECT_COPY
                          #SET(%EmbedNumber,%EmbedNumber+1)
                          #EMBED(%OleDataSourceOnLButtonDownDropSuccessCopy,'Ole Data Source OnLButtonDown: Drag Success Copy'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Copy After Successful Drag)')
                     #ENDIF
                     #IF(%MoveEffect)
                     OF   DROPEFFECT_MOVE
                          #SET(%EmbedNumber,%EmbedNumber+1)
                          #EMBED(%OleDataSourceOnLButtonDownDropSuccessMove,'Ole Data Source OnLButtonDown: Drag Success Move'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Move After Successful Drag)')
                     #ENDIF
                     #IF(%LinkEffect)
                     OF   DROPEFFECT_LINK
                          #SET(%EmbedNumber,%EmbedNumber+1)
                          #EMBED(%OleDataSourceOnLButtonDownDropSuccessLink,'Ole Data Source OnLButtonDown: Drag Success Link'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Link After Successful Drag)')
                     #ENDIF
                     END
                  ELSE
                     #SET(%EmbedNumber,%EmbedNumber+1)
                     #EMBED(%OleDataSourceOnLButtonDownDropFail,'Ole Data Source OnLButtonDown: Drag Failure'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (After Failed Drag)')
                  END
                  pTextDataObject.Kill()
                  kcr_OleSetClipBoard(0)
                  pIDropSource.Kill()
                  DISPOSE(pIDropSource)
               END
               pTextDataObject.Release()
               pTextDataObject &= NULL
            END
      #IF(%IncludeDrop)
            hr = pOleDropTarget.Init(hWnd,%EventOleDrop)
            pOleDropTarget &= NULL
      #ENDIF
            pDragText &= NULL
      #CASE(%ControlType)
      #OF('TEXT')
      #OROF('ENTRY')
            kcr_SetFocus(hwnd)
      #ELSE
            kcr_SetFocus(%Window{PROP:ClientHandle})
      #ENDCASE
            RETURN Level:Benign
         ELSE
            pDragText &= NULL
      #CASE(%ControlType)
      #OF('TEXT')
      #OROF('ENTRY')
            kcr_SetFocus(hwnd)
      #OF('LIST')
            kcr_SetFocus(%Window{PROP:ClientHandle})
      #ENDCASE
            RETURN Level:Fatal
         END
      ELSE    !No Drag Detected
      #CASE(%ControlType)
      #OF('TEXT')
      #OROF('ENTRY')
         %DragDropControl{PROP:SelEnd}   = 0
         %DragDropControl{PROP:SelStart} = SelStart
      #ENDCASE
      END
      #IF(%ControlType = 'LIST')
      kcr_SetFocus(%Window{PROP:ClientHandle})
      #ENDIF
      #SET(%EmbedNumber,%EmbedNumber+1)
      #EMBED(%OleDataSourceOnLButtonDownEndCodeSection,'Ole Data Source OnLButtonDown:  End of Code'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (End of Code)')
      RETURN Level:Benign
!------------------------------------------------------------------------------

  #ENDIF
#!
#PRIORITY(8000)
#!
#FIX(%Control,%DragDropControl)
  #IF(%IncludeDrag)
#SET(%ValueConstruct,SUB(%DragDropControl,2,LEN(%DragDropControl)-1) & ':WndProc' & %ActiveTemplateInstance)
%20ValueConstruct PROCEDURE(UNSIGNED hWnd, UNSIGNED wMsg, UNSIGNED wParam, LONG lParam)
EventPosted     BOOL(FALSE),STATIC
pOleDataSource  LONG
OleDataSource   &OleDataSourceClass
feq             LONG
SelStart        LONG
SelEnd          LONG
#EMBED(%OleDragDropSubClassWindowDataSection,'Ole Drag and Drop (SubClassWindow):  Data Section'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Data Section)')

    CODE
      #SET(%EmbedNumber,%EmbedNumber+1)
      #EMBED(%OleDragDropSubClassWindowBeginCodeSection,'Ole Drag and Drop (SubClassWindow):  Start of Code'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Start of Code)')
      CASE wMsg
        #SET(%EmbedNumber,%EmbedNumber+1)
        #EMBED(%OleDragDropSubClassWindowCaseSection,'Ole Drag and Drop (SubClassWindow):  Inside CASE wMsg statement'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (Inside CASE wMsg)')
        OF WM_LBUTTONDOWN
           #IF(%ControlType = 'LIST')
           IF ~EventPosted
              POST(EVENT:MOUSEDOWN,kcr_GetProp(hwnd,szPropFieldEquate))
           END
           #ELSE
              feq = kcr_GetProp(hwnd,szPropFieldEquate)
              pOleDataSource = kcr_GetProp(hwnd,szPropDragData)
              OleDataSource &= (pOleDataSource)
              OleDataSource.m_SelStart = feq{PROP:SelStart}
              OleDataSource.m_SelEnd = feq{PROP:SelEnd}
              IF OleDataSource.m_SelEnd and OleDataSource.m_SelStart <> OleDataSource.m_SelEnd
                 IF ~EventPosted
                    POST(EVENT:MOUSEDOWN,kcr_GetProp(hwnd,szPropFieldEquate))
                 END
              END
           #END
        OF WM_LBUTTONUP
           EventPosted = FALSE
      END
      #SET(%EmbedNumber,%EmbedNumber+1)
      #EMBED(%OleDragDropSubClassWindowEndCodeSection,'Ole Drag and Drop (SubClassWindow):  End of Code'),%ActiveTemplateInstance,TREE('Ole Drag and Drop',%DragDropControl,%ValueConstruct,%EmbedNumber & '. ' & %ValueConstruct & ' (End of Code)')
      #SET(%ValueConstruct,SUB(%DragDropControl,2,LEN(%DragDropControl)-1) & ':origWndProc' & %ActiveTemplateInstance)
      RETURN kcr_CallWindowProc(%ValueConstruct,hWnd,wMsg,wParam,lParam)
!------------------------------------------------------------------------------

  #ENDIF
#ENDAT
#!
#!
#GROUP(%LoadCSTRINGs),PRESERVE
#IF(%OleDataBufferOption = 'Create')
  #ADD(%AllCSTRINGs, %OleDataBufferCreate)
  #SET(%StringSize, %OleDataBufferCreateSize)
#ENDIF
#FOR(%GlobalData),WHERE(EXTRACT(%GlobalDataStatement, 'CSTRING') AND EXTRACT(%GlobalDataStatement, 'THREAD'))
  #ADD(%AllCSTRINGs, %GlobalData)
  #SET(%StringSize, (EXTRACT(%GlobalDataStatement, 'CSTRING', 1) - 1))
#ENDFOR
#!---
#FIND(%ModuleProcedure, %Procedure, %Module)
#FOR(%ModuleData),WHERE(EXTRACT(%ModuleDataStatement, 'CSTRING') AND EXTRACT(%ModuleDataStatement, 'THREAD'))
  #ADD(%AllCSTRINGs, %ModuleData)
  #SET(%StringSize, (EXTRACT(%ModuleDataStatement, 'CSTRING', 1) - 1))
#ENDFOR
#!---
#FOR(%ActiveTemplate)
  #FOR(%ActiveTemplateInstance)
    #FIX(%File,%Primary)
    #FOR(%Field),WHERE(%FieldType = 'CSTRING')
      #ADD(%AllCSTRINGs, %Field)
      #SET(%StringSize, (EXTRACT(%FieldStatement, 'CSTRING', 1) - 1))
    #ENDFOR
    #FOR(%Secondary)
      #FIX(%File,%Secondary)
      #FOR(%Field),WHERE(%FieldType = 'CSTRING')
        #ADD(%AllCSTRINGs, %Field)
        #SET(%StringSize, (EXTRACT(%FieldStatement, 'CSTRING', 1) - 1))
      #ENDFOR
    #ENDFOR
    #FOR(%OtherFiles)
      #FIX(%File,%OtherFiles)
      #FOR(%Field),WHERE(%FieldType = 'CSTRING')
        #ADD(%AllCSTRINGs, %Field)
        #SET(%StringSize, (EXTRACT(%FieldStatement, 'CSTRING', 1) - 1))
      #ENDFOR
    #ENDFOR
  #ENDFOR
#ENDFOR
#!
#!
#GROUP(%SetOleDataOverrideSize,%DataBuffer)
#FIX(%AllCSTRINGs,%DataBuffer)
#SET(%OleDataBufferOverrideSize,%StringSize)
#!
#!
#!
#!
#!-----------------------------------------------------------------------------------------------------------
#!The following template code is derived from "The Clarion Handy Tools" created by Gus Creces.
#!Used with permission.  Look for "The Clarion Handy Tools" at www.cwhandy.com
#!-----------------------------------------------------------------------------------------------------------
#SYSTEM
 #TAB('Devuna OLE Drag and Drop Templates')
  #INSERT  (%SysHead)
  #BOXED   ('About Devuna OLE Drag and Drop Templates'),AT(5)
    #DISPLAY (''),AT(15)
    #DISPLAY ('Warning: These templates are protected by copyright   '),AT(15)
    #DISPLAY ('law and international treaties.   Unauthorized use,   '),AT(15)
    #DISPLAY ('reproduction or distribution of these templates, or   '),AT(15)
    #DISPLAY ('any part of  them, without the expressed written      '),AT(15)
    #DISPLAY ('consent of Devuna Inc., may      '),AT(15)
    #DISPLAY ('result in severe civil and criminal penalties, and    '),AT(15)
    #DISPLAY ('will be prosecuted to the maximum extent possible     '),AT(15)
    #DISPLAY ('under law.'),AT(15)
    #DISPLAY (''),AT(15)
    #DISPLAY ('For more information, contact the author at:'),AT(15)
    #DISPLAY ('rrogers@devuna.com'),AT(15)
    #DISPLAY ('(C)2006-2017 Devuna'),AT(15)
  #ENDBOXED
 #ENDTAB

#!-----------------------------------------------------------------------------------------------------------
#GROUP (%MakeHeadHiddenPrompts)
  #PROMPT('',@S50),%TplName
  #PROMPT('',@S100),%TplDescription
  #PROMPT('',@S10),%TplVersion
#!-----------------------------------------------------------------------------------------------------------
#GROUP   (%MakeHead,%xTplName,%xTplDescription,%xTplVersion)
  #SET (%TplName,%xTplName)
  #SET (%TplDescription,%xTplDescription)
  #SET (%TplVersion,%xTplVersion)
#!
#!-----------------------------------------------------------------------------------------------------------
#GROUP   (%Head)
  #IMAGE   ('DragDrop.ico'), AT(,,175,26)
  #DISPLAY (%TplName),AT(40,4)
  #DISPLAY ('Version ' & %TplVersion),AT(40,12)
  #DISPLAY ('(C)2006-2017 Devuna'),AT(40,20)
  #DISPLAY ('')
#!
#!-----------------------------------------------------------------------------------------------------------
#GROUP   (%SysHead)
  #IMAGE   ('DragDrop.ico'), AT(,4,175,26)
  #DISPLAY ('OleDragDrop.tpl'),AT(40,4)
  #DISPLAY ('Devuna OLE Drag and Drop ABC Templates'),AT(40,14)
  #DISPLAY ('for Clarion ABC Template Applications'),AT(40,24)
  #DISPLAY ('')
#!
#!-----------------------------------------------------------------------------------------------------------
#GROUP(%EmbedStart)
#?!-----------------------------------------------------------------------------------------------------------
#?! OleDragDrop.tpl   (C)2006-2017 Devuna
#?! Template: (%TplName - %TplDescription)
#IF (%EmbedID)
#?! Embed:    (%EmbedID) (%EmbedDescription) (%EmbedParameters)
#ENDIF
#?!-----------------------------------------------------------------------------------------------------------
#!
#!----------------------------------------------------------------------------------------------------------
#GROUP(%EmbedEnd)
#?!-----------------------------------------------------------------------------------------------------------
#!End of derived work.  Thanks Gus!
#!-----------------------------------------------------------------------------------------------------------
