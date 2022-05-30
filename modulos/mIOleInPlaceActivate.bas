Attribute VB_Name = "mIOleInPlaceActivate"
'========================================================================================
' Filename:    mIOleInPlaceActivate.bas
' Author:      Mike Gainer, Matt Curland and Bill Storage
' Date:        09 January 1999
'
' Requires:    Nothing ;-)
'
' Description:
' Allows you to replace the standard IOLEInPlaceActiveObject interface for a
' UserControl with a customisable one.  This allows you to take control
' of focus in VB controls.
'
' The code could be adapted to replace other UserControl OLE interfaces.
'
' ---------------------------------------------------------------------------------------
' Visit vbAccelerator, advanced, free source for VB programmers
' http://vbaccelerator.com
'========================================================================================
' Changed: No tlb needed. Frank Schüler ActiveVB.de
'========================================================================================

Option Explicit

'========================================================================================
' Constants
'========================================================================================
Private Const S_OK As Long = &H0&
Private Const S_FALSE As Long = &H1&
Private Const WM_KEYUP As Long = &H101&
Private Const WM_KEYDOWN As Long = &H100&
Private Const CC_STDCALL As Long = &H4&
Private Const IID_Release As Long = &H8&
Private Const E_NOINTERFACE As Long = &H80004002
Private Const OLEIVERB_UIACTIVATE As Long = &HFFFFFFFC

Private Const IID_IOleObject As String = "{00000112-0000-0000-c000-000000000046}"
Private Const IID_IOleInPlaceSite As String = "{00000119-0000-0000-c000-000000000046}"
Private Const IID_IOleInPlaceActiveObject As String = "{00000117-0000-0000-c000-000000000046}"

Private Enum vtb_Interfaces

    ' IUnknown
    QueryInterface = 0
    AddRef = 1
    Release = 2
    
    ' IOleObject
    GetClientSite = 4
    DoVerb = 11
    
    ' IOleInPlaceSite
    GetWindowContext = 8
    
    ' IOleInPlaceFrame / IOleInPlaceUIWindow
    SetActiveObject = 8
    
    ' IOleInPlaceActiveObject
    GetWindow = 3
    ContextSensitiveHelp = 4
    TranslateAccelerator = 5
    OnFrameWindowActivate = 6
    OnDocWindowActivate = 7
    ResizeBorder = 8
    EnableModeless = 9

End Enum

'========================================================================================
' Lightweight object definition
'========================================================================================
Public Type IPAOHookStruct
    lpVTable    As Long     'VTable pointer
    IPAOReal    As Long     'Un-AddRefed pointer for forwarding calls
    Ctrl         As Object   'Un-AddRefed native class pointer for making Friend calls
    ThisPointer As Long
End Type

'========================================================================================
' Types
'========================================================================================
Private Type POINT
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type Msg
    hWnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINT
    lPrivate As Long
End Type

Private Type OLEINPLACEFRAMEINFO
    cb As Long
    fMDIApp As Long
    hwndFrame As Long
    haccel As Long
    cAccelEntries As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

'========================================================================================
' API
'========================================================================================
Private Declare Function SendMessageLong Lib "user32.dll" _
                         Alias "SendMessageA" ( _
                         ByVal hWnd As Long, _
                         ByVal wMsg As Long, _
                         ByVal wParam As Long, _
                         ByRef lParam As Long) As Long
                         
Private Declare Sub CopyMemory Lib "kernel32" _
                    Alias "RtlMoveMemory" ( _
                    Destination As Any, _
                    Source As Any, _
                    ByVal Length As Long)
                    
Private Declare Function IsEqualGUID Lib "ole32" ( _
                         iid1 As GUID, _
                         iid2 As GUID) As Long
                         
Private Declare Sub CoTaskMemFree Lib "Ole32.dll" ( _
                    ByVal hMem As Long)

Private Declare Function StringFromCLSID Lib "Ole32.dll" ( _
                         ByRef pCLSID As GUID, _
                         ByRef lpszProgID As Long) As Long

Private Declare Function CLSIDFromString Lib "Ole32.dll" ( _
                         ByVal pstring As Long, _
                         ByRef pCLSID As GUID) As Long

Private Declare Function lstrlenW Lib "Kernel32.dll" ( _
                         ByVal lpString As Long) As Long

Private Declare Function DispCallFunc Lib "Oleaut32.dll" ( _
                         ByVal pvInstance As Long, _
                         ByVal oVft As Long, _
                         ByVal cc As Long, _
                         ByVal vtReturn As VbVarType, _
                         ByVal cActuals As Long, _
                         ByVal prgvt As Long, _
                         ByVal prgpvarg As Long, _
                         ByRef pvargResult As Variant) As Long


'========================================================================================
' Member variables
'========================================================================================
Private m_IID_IOleInPlaceActiveObject As GUID
Private m_IPAOVTable(9)             As Long

'========================================================================================
' Functions
'========================================================================================
Public Sub InitIPAO(IPAOHookStruct As IPAOHookStruct, Ctrl As Object)

    Dim pIOleInPlaceActiveObject As Long
    
    If Invoke(ObjPtr(Ctrl), QueryInterface, VarPtr(Str2Guid(IID_IOleInPlaceActiveObject)), _
        VarPtr(pIOleInPlaceActiveObject)) = S_OK Then
        
        With IPAOHookStruct
            Call CopyMemory(.IPAOReal, pIOleInPlaceActiveObject, 4)
            Call CopyMemory(.Ctrl, Ctrl, 4)
            .lpVTable = GetVTable
            .ThisPointer = VarPtr(IPAOHookStruct)
        End With
        
        Call ReleaseInterface(pIOleInPlaceActiveObject)
        
    End If

End Sub

Public Sub TerminateIPAO(IPAOHookStruct As IPAOHookStruct)
    With IPAOHookStruct
        Call CopyMemory(.IPAOReal, 0&, 4)
        Call CopyMemory(.Ctrl, 0&, 4)
    End With
End Sub

'========================================================================================
' Private
'========================================================================================
Private Function GetVTable() As Long

    ' Set up the vTable for the interface and return a pointer to it
    If (m_IPAOVTable(0) = 0) Then
        m_IPAOVTable(0) = AddressOfFunction(AddressOf QueryInterface_)
        m_IPAOVTable(1) = AddressOfFunction(AddressOf AddRef_)
        m_IPAOVTable(2) = AddressOfFunction(AddressOf Release_)
        m_IPAOVTable(3) = AddressOfFunction(AddressOf GetWindow_)
        m_IPAOVTable(4) = AddressOfFunction(AddressOf ContextSensitiveHelp_)
        m_IPAOVTable(5) = AddressOfFunction(AddressOf TranslateAccelerator_)
        m_IPAOVTable(6) = AddressOfFunction(AddressOf OnFrameWindowActivate_)
        m_IPAOVTable(7) = AddressOfFunction(AddressOf OnDocWindowActivate_)
        m_IPAOVTable(8) = AddressOfFunction(AddressOf ResizeBorder_)
        m_IPAOVTable(9) = AddressOfFunction(AddressOf EnableModeless_)
        '--- init guid
         m_IID_IOleInPlaceActiveObject = Str2Guid(IID_IOleInPlaceActiveObject)
    End If
    GetVTable = VarPtr(m_IPAOVTable(0))
End Function

Private Function AddressOfFunction(lpfn As Long) As Long
    ' Work around, VB thinks lPtr = AddressOf Method is an error
    AddressOfFunction = lpfn
End Function

'========================================================================================
' Interface implemenattion
'========================================================================================
Private Function AddRef_(This As IPAOHookStruct) As Long
    AddRef_ = Invoke(This.IPAOReal, AddRef)
End Function

Private Function Release_(This As IPAOHookStruct) As Long
    Release_ = Invoke(This.IPAOReal, Release)
End Function

Private Function QueryInterface_(This As IPAOHookStruct, riid As GUID, pvObj As Long) As Long
    ' Install the interface if required
    If (IsEqualGUID(riid, m_IID_IOleInPlaceActiveObject)) Then
        ' Install alternative IOleInPlaceActiveObject interface implemented here
        pvObj = This.ThisPointer
        AddRef_ This
        QueryInterface_ = 0
      Else
        ' Use the default support for the interface:
        QueryInterface_ = Invoke(This.IPAOReal, QueryInterface, VarPtr(riid), pvObj)
    End If
End Function

Private Function GetWindow_(This As IPAOHookStruct, phwnd As Long) As Long
    GetWindow_ = Invoke(This.IPAOReal, GetWindow, phwnd)
End Function

Private Function ContextSensitiveHelp_(This As IPAOHookStruct, ByVal fEnterMode As Long) As Long
    ContextSensitiveHelp_ = Invoke(This.IPAOReal, ContextSensitiveHelp, fEnterMode)
End Function

Private Function TranslateAccelerator_(This As IPAOHookStruct, lpMsg As Msg) As Long

    On Error Resume Next
    
    Select Case lpMsg.message
    
        Case WM_KEYDOWN, WM_KEYUP
        
            Select Case lpMsg.wParam
                Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyHome, vbKeyEnd, vbKeyPageDown, vbKeyPageUp
                    SendMessageLong lpMsg.hWnd, lpMsg.message, lpMsg.wParam, lpMsg.lParam
                    TranslateAccelerator_ = S_OK
                Case Else
                    TranslateAccelerator_ = Invoke(This.IPAOReal, TranslateAccelerator, VarPtr(lpMsg))
            End Select
        Case Else
            TranslateAccelerator_ = Invoke(This.IPAOReal, TranslateAccelerator, VarPtr(lpMsg))
    End Select
    
    On Error GoTo 0

End Function

Private Function OnFrameWindowActivate_(This As IPAOHookStruct, ByVal fActivate As Long) As Long
    OnFrameWindowActivate_ = Invoke(This.IPAOReal, OnFrameWindowActivate, fActivate)
End Function

Private Function OnDocWindowActivate_(This As IPAOHookStruct, ByVal fActivate As Long) As Long
    OnDocWindowActivate_ = Invoke(This.IPAOReal, OnDocWindowActivate, fActivate)
End Function

Private Function ResizeBorder_(This As IPAOHookStruct, prcBorder As RECT, ByVal puiWindow As Long, ByVal fFrameWindow As Long) As Long
    ResizeBorder_ = Invoke(This.IPAOReal, ResizeBorder, VarPtr(prcBorder), puiWindow, fFrameWindow)
End Function

Private Function EnableModeless_(This As IPAOHookStruct, ByVal fEnable As Long) As Long
    EnableModeless_ = Invoke(This.IPAOReal, EnableModeless, fEnable)
End Function


Public Sub SetIPAO(uIPAO As IPAOHookStruct, Ctrl As Object)

    Dim pIOleObject As Long
    Dim pIOleClientSite As Long
    Dim pIOleInPlaceSite As Long
    Dim pIOleInPlaceFrame As Long
    Dim pIOleInPlaceUIWindow As Long
    Dim rcPos As RECT
    Dim rcClip As RECT
    Dim uFrameInfo As OLEINPLACEFRAMEINFO
       
    On Error Resume Next
    
        If Invoke(ObjPtr(Ctrl), QueryInterface, VarPtr(Str2Guid(IID_IOleObject)), _
            VarPtr(pIOleObject)) = S_OK Then
        
        If Invoke(pIOleObject, GetClientSite, VarPtr(pIOleClientSite)) = S_OK Then
        
            If Invoke(pIOleClientSite, QueryInterface, VarPtr(Str2Guid(IID_IOleInPlaceSite)), _
                VarPtr(pIOleInPlaceSite)) = S_OK Then
                
                If Invoke(pIOleInPlaceSite, GetWindowContext, _
                    VarPtr(pIOleInPlaceFrame), VarPtr(pIOleInPlaceUIWindow), _
                    VarPtr(rcPos), VarPtr(rcClip), VarPtr(uFrameInfo)) = S_OK Then
                
                    If pIOleInPlaceFrame <> 0& Then
                    
                        If Invoke(pIOleInPlaceFrame, SetActiveObject, _
                            uIPAO.ThisPointer, 0&) = S_OK Then

                        End If
                        
                        Call ReleaseInterface(pIOleInPlaceFrame)
                        
                    End If
            
                    If pIOleInPlaceUIWindow <> 0& Then  '-- And Not m_bMouseActivate
                    
                        If Invoke(pIOleInPlaceUIWindow, SetActiveObject, _
                            uIPAO.ThisPointer, 0&) = S_OK Then
    
                        End If
                        
                        Call ReleaseInterface(pIOleInPlaceUIWindow)
                        
                    Else
                
                        If Invoke(pIOleObject, DoVerb, OLEIVERB_UIACTIVATE, 0&, _
                            pIOleInPlaceSite, 0&, 0&, VarPtr(rcPos)) = S_OK Then
                    
                        End If
                    
                    End If
            
                End If
            
                Call ReleaseInterface(pIOleInPlaceSite)
            
            End If

            Call ReleaseInterface(pIOleClientSite)

        End If
        
        Call ReleaseInterface(pIOleObject)
        
    End If

    On Error GoTo 0
End Sub

' ===================================================================================
Private Function Invoke(ByVal pInterface As Long, _
    ByVal eInterfaceFunction As vtb_Interfaces, _
    ParamArray arrParam()) As Variant

    If pInterface <> 0& Then
        
        Invoke = OleInvoke(pInterface, eInterfaceFunction, arrParam)

    End If

End Function

Private Sub ReleaseInterface(ByRef pInterface As Long)

    Dim varRet As Variant

    If pInterface <> 0& Then

        If DispCallFunc(pInterface, IID_Release, CC_STDCALL, _
            vbLong, 0&, 0&, 0&, varRet) = S_OK Then

            pInterface = 0&

        End If

    End If

End Sub

Private Function OleInvoke(ByVal pInterface As Long, _
    ByVal lngCmd As Long, _
    ParamArray arrParam()) As Variant
    
    Dim lngRet As Long
    Dim lngItem As Long
    Dim lngCount As Long
    Dim varRet As Variant
    Dim varParam As Variant
    Dim olePtr(10) As Long
    Dim oleTyp(10) As Integer

    If pInterface <> 0& Then

        If UBound(arrParam) >= 0 Then

            varParam = arrParam

            If IsArray(varParam) Then varParam = varParam(0)

            lngCount = UBound(varParam)

            For lngItem = 0 To lngCount

                oleTyp(lngItem) = VarType(varParam(lngItem))
                olePtr(lngItem) = VarPtr(varParam(lngItem))

            Next

        End If

        lngRet = DispCallFunc(pInterface, lngCmd * 4, CC_STDCALL, vbLong, _
            lngItem, VarPtr(oleTyp(0)), VarPtr(olePtr(0)), varRet)
            
        If lngRet <> S_OK Then
            
            'Debug.Print "Error: 0x" & Hex$(lngRet)
            
            varRet = S_FALSE
            
        End If
        
    End If

    OleInvoke = varRet

End Function

Private Function GetStringFromPointer(ByVal lpStrPointer As Long) As String

    Dim lLen As Long
    Dim bBuffer() As Byte

    lLen = lstrlenW(lpStrPointer) * 2

    If lLen > 0 Then

        ReDim bBuffer(lLen - 1)

        Call CopyMemory(bBuffer(0), ByVal lpStrPointer, lLen)

        Call CoTaskMemFree(lpStrPointer)

        GetStringFromPointer = bBuffer

    End If

End Function

Private Function Str2Guid(ByVal str As String) As GUID

    Call CLSIDFromString(StrPtr(str), Str2Guid)
    
End Function

Private Function Guid2Str(ByRef tguid As GUID) As String

    Dim lGuid As Long
    
    Call StringFromCLSID(tguid, lGuid)

    Guid2Str = GetStringFromPointer(lGuid)

End Function




