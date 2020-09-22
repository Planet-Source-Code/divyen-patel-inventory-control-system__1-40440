VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ======================================================================
' See the module modMenus:ReadMe for rules on using this class
' ======================================================================

Private ImageLister As Control
' name of menus imagelist, can't use handle 'cause it can change per MSDN
Public MainMenuID As Long  ' handle to form's main menu
Public OldWinProc As Long  ' handle to form's window message processor
Public ChildStatus As Byte ' 1 indicates a child
Public ParentForm As Long  ' for MDI children, this is the MDI parent -- for others it is it's own handle
Private MyMI() As MenuDataInformation       ' collection of menuitems
Private menuIDs As Collection               ' index to myMI array
Private mIDcurrent As Long                  ' current menu item
Private MyPanels() As PanelDataInformation
Private PanelData As Collection
Private pIDcurrent As Long
Private MDIchildren As Collection

Property Get TotalIcons() As Long
    On Error Resume Next
    TotalIcons = ImageLister.ListImages.Count
End Property

Property Let Icon(lValue As Long)
    MyMI(menuIDs(mIDcurrent)).Icon = lValue
End Property
Property Get Icon() As Long
    Icon = MyMI(menuIDs(mIDcurrent)).Icon
End Property

Property Let ItemHeight(lValue As Long)
    MyMI(menuIDs(mIDcurrent)).ItemHeight = lValue
End Property
Property Get ItemHeight() As Long
    ItemHeight = MyMI(menuIDs(mIDcurrent)).ItemHeight
End Property
Property Let ItemWidth(lValue As Long)
    MyMI(menuIDs(mIDcurrent)).ItemWidth = lValue
End Property
Property Get ItemWidth() As Long
    ItemWidth = MyMI(menuIDs(mIDcurrent)).ItemWidth
End Property
Property Let HotKeyPos(lValue As Long)
    MyMI(menuIDs(mIDcurrent)).HotKeyPos = lValue
End Property
Property Get HotKeyPos() As Long
    HotKeyPos = MyMI(menuIDs(mIDcurrent)).HotKeyPos
End Property

Property Let Status(lValue As Long)
    MyMI(menuIDs(mIDcurrent)).Status = lValue
End Property
Property Get Status() As Long
    Status = MyMI(menuIDs(mIDcurrent)).Status
End Property

Property Let Caption(sValue As String)
    MyMI(menuIDs(mIDcurrent)).Caption = sValue
End Property
Property Get Caption() As String
    Caption = MyMI(menuIDs(mIDcurrent)).Caption
End Property

Property Get ImageViewerObj() As Control
    On Error Resume Next
    Set ImageViewerObj = ImageLister
End Property
Property Get ImageViewer() As Long
   On Error Resume Next
   ImageViewer = ImageLister.hImageList
End Property
Public Sub SetImageViewer(vObject As Control)
    Set ImageLister = vObject
End Sub

Property Get OriginalCaption() As String
    OriginalCaption = MyMI(menuIDs(mIDcurrent)).OriginalCaption
End Property
Property Let OriginalCaption(sValue As String)
    MyMI(menuIDs(mIDcurrent)).OriginalCaption = sValue
End Property

Property Get SideBarIsText() As Boolean
    On Error Resume Next
    SideBarIsText = (MyPanels(CStr(pIDcurrent)).Status And 4) = 4
End Property
Property Get SideBarItem() As Long
    SideBarItem = MyPanels(CStr(pIDcurrent)).SBarIcon
End Property
Property Get SideBarWidth() As Long
    SideBarWidth = MyPanels(CStr(pIDcurrent)).SideBar
End Property
Property Get PanelWidth() As Long
    PanelWidth = MyPanels(CStr(pIDcurrent)).Width
End Property
Property Get PanelIDcount() As Long
    On Error Resume Next
    PanelIDcount = PanelData.Count
End Property
Property Get PanelHeight() As Long
    On Error Resume Next
    PanelHeight = MyPanels(CStr(pIDcurrent)).Height
End Property
Property Get HotKeyEdge() As Integer
    HotKeyEdge = CInt(MyPanels(CStr(pIDcurrent)).HKeyPos)
End Property

Public Function GetSetMDIchildSysMenu(lValue As Long, bSet As Boolean) As Boolean
On Error Resume Next
Dim lHwnd As Long
If bSet = True Then
    If MDIchildren Is Nothing Then Set MDIchildren = New Collection
    lHwnd = MDIchildren(CStr(lValue))
    If lHwnd = 0 Then MDIchildren.Add MDIchildren.Count + 1, CStr(lValue)
Else
    lHwnd = MDIchildren(CStr(lValue))
    GetSetMDIchildSysMenu = (lHwnd <> 0)
End If
Err.Clear
End Function

Property Get MenuIDcount() As Integer
' =====================================================================
' Simply returns the number of menu items processed
' =====================================================================
    On Error Resume Next
    MenuIDcount = menuIDs.Count
End Property

Public Sub UpdatePanelID(vData() As Long, sText As String, bPartial As Boolean)
    On Error Resume Next
    With MyPanels(CStr(pIDcurrent))
        .Width = vData(0)
        .Height = vData(1)
        .HKeyPos = vData(2)
        .PanelIcon = vData(3)
        If bPartial = False Then
            'Debug.Print "full update on paneldata"
            .SideBar = vData(4)
            .SideBarXY = vData(5)
            .BColor = vData(6)
            .FColor = vData(7)
            .Caption = sText
            .Status = vData(9)
            .SBarIcon = vData(10)
        End If
    End With
End Sub

Public Sub GetPanelInformation(vData() As Long, sText As String)
On Error Resume Next
ReDim vData(0 To 10)
With MyPanels(PanelData(CStr(MyMI(menuIDs(mIDcurrent)).Parent)))
    vData(0) = .Width + 16
    vData(1) = .Height
    vData(2) = .HKeyPos
    vData(3) = .PanelIcon
    vData(4) = .SideBar
    vData(5) = .SideBarXY
    vData(6) = .BColor
    vData(7) = .FColor
    sText = .Caption
    vData(9) = .Status
    vData(10) = .SBarIcon
End With
End Sub

Public Function SetMenuID(iID As Long, hSubMenu As Long, byPosition As Boolean, Optional bAlwaysCreate As Boolean = True) As Boolean
' =====================================================================
' Used to create a new reference to a menu item or point to
' an existing reference
' =====================================================================

    On Error Resume Next
    ' we reference passed menu item, if we don't have a reference
    ' an error occurs which triggers a new reference if the
    ' bAlwaysCreate boolean is set to true
    If byPosition Then
        ' menu item is positional (i.e., 1,2,3)
        mIDcurrent = iID
    Else
        ' menu item is by ID vs position
        mIDcurrent = menuIDs(CStr(iID) & "." & CStr(hSubMenu))
    End If
    If Err Then ' new reference
        If bAlwaysCreate = True Then
            ' let's add a new reference & use the menu ID as a key
            menuIDs.Add menuIDs.Count + 1, CStr(iID) & "." & CStr(hSubMenu)
            mIDcurrent = menuIDs.Count
            ' now we will add an MyMI array
            ReDim Preserve MyMI(1 To menuIDs.Count)
            MyMI(menuIDs(mIDcurrent)).ID = iID
            MyMI(menuIDs(mIDcurrent)).Parent = hSubMenu
            ' return a value indicating this is a new add
            SetMenuID = True
            Err.Clear
            pIDcurrent = PanelData(CStr(hSubMenu))
            If Err Then
                Err.Clear
                PanelData.Add PanelData.Count + 1, CStr(hSubMenu)
                ReDim Preserve MyPanels(1 To PanelData.Count)
                MyPanels(PanelData.Count).ID = hSubMenu
                'Debug.Print "new panel created-count="; hSubMenu; PanelData.Count
            End If
        End If
    Else    ' reference already exists
        ' if the following flag wasn't set, then the drawing/measuring
        ' routine wants to know if we have a reference
        ' so we set return to true if so
        ' otherwise, the menu metrics is calling this and we
        ' need to return false indicating this is not a new add
        If bAlwaysCreate = False Then SetMenuID = True
    End If
    pIDcurrent = PanelData(CStr(hSubMenu))
End Function

Public Sub GetIconData(vData() As Long, IconIndex As Long)
' =====================================================================
' Returns image handle, type and icon index/transparency option
' when drawing routine requests it
' =====================================================================
    On Error Resume Next
    ReDim vData(0 To 2)
    If Not ImageLister Is Nothing Then
        vData(0) = ImageLister.ListImages(IconIndex).Picture.Handle
        vData(1) = ImageLister.ListImages(IconIndex).Picture.Type
        If (MyMI(menuIDs(mIDcurrent)).Status And 4) = 4 Then
            vData(2) = 1
        Else
            If (MyMI(menuIDs(mIDcurrent)).Status And 8) = 8 Then vData(2) = 2
        End If
    End If
End Sub

Public Function GetPanelID(iID As Long) As Long
' =====================================================================
' Returns then actual menuID and related submenu item for
' a stored menuitem -- used in preparation for the DeleteMenuItem sub
' =====================================================================
On Error Resume Next
GetPanelID = MyPanels(PanelData.Item(iID)).ID
End Function

Public Sub PurgeObsoleteMenus(hSubMenu As Long)
Dim newMyMI() As MenuDataInformation, newMyPanels() As PanelDataInformation
Dim Looper As Long, lCounter As Long
On Error GoTo ExitSub
If menuIDs.Count Then
    ReDim newMyMI(1 To menuIDs.Count)
    lCounter = 1
    For Looper = menuIDs.Count To 1 Step -1
        If MyMI(menuIDs.Item(Looper)).Parent <> hSubMenu Then
            newMyMI(lCounter) = MyMI(menuIDs.Item(Looper))
            lCounter = lCounter + 1
        End If
    Next
    If lCounter - 1 Then
        Erase MyMI
        ReDim MyMI(1 To lCounter - 1)
        Set menuIDs = Nothing
        Set menuIDs = New Collection
        For Looper = 1 To lCounter - 1
            MyMI(Looper) = newMyMI(Looper)
            menuIDs.Add Looper, CStr(newMyMI(Looper).ID) & "." & CStr(newMyMI(Looper).Parent)
        Next
        'Debug.Print "Finished indexing menuitems"
    End If
    Erase newMyMI
End If
If PanelData.Count Then
    ReDim newMyPanels(1 To PanelData.Count)
    lCounter = 1
    For Looper = PanelData.Count To 1 Step -1
        pIDcurrent = Looper
        If MyPanels(PanelData.Item(Looper)).ID <> hSubMenu Then
            newMyPanels(lCounter) = MyPanels(PanelData.Item(Looper))
            If SideBarIsText = True Then DeleteObject MyPanels(PanelData.Item(Looper)).SBarIcon
            lCounter = lCounter + 1
        End If
    Next
    If lCounter - 1 Then
        Erase MyPanels
        ReDim MyPanels(1 To lCounter - 1)
        Set PanelData = Nothing
        Set PanelData = New Collection
        For Looper = 1 To lCounter - 1
            MyPanels(Looper) = newMyPanels(Looper)
            PanelData.Add Looper, CStr(newMyPanels(Looper).ID)
        Next
        'Debug.Print "Finished Indexing panels"
    End If
    Erase newMyPanels
End If
Looper = 0
Looper = MDIchildren(CStr(hSubMenu))
If Looper Then MDIchildren.Remove Looper
ExitSub:
End Sub

Private Sub Class_Terminate()
' =====================================================================
' Clean up variables, collections, etc for form closure
' =====================================================================
On Error Resume Next
Dim Looper As Long
For Looper = 1 To PanelData.Count
    pIDcurrent = Looper
    If SideBarIsText = True Then DeleteObject MyPanels(PanelData.Item(Looper)).SBarIcon
Next
Set PanelData = Nothing
Set MDIchildren = Nothing
Set ImageLister = Nothing
MainMenuID = 0
OldWinProc = 0
Set menuIDs = Nothing
Erase MyMI
End Sub

Private Sub Class_Initialize()
Set menuIDs = New Collection
Set PanelData = New Collection
End Sub


