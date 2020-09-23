VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form frmBrowser 
   Caption         =   "Auto Borwser"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10245
   Icon            =   "frmBrowser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   10245
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrWorking 
      Interval        =   50
      Left            =   6360
      Top             =   3960
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   5790
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12435
            Key             =   "URL"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "Information"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "7:58 AM"
            Key             =   "Clock"
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstCommands 
      Height          =   645
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1920
      Width           =   4695
   End
   Begin MSComctlLib.ImageList ilButtonsDisabled 
      Left            =   8280
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":0E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":139E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":18FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":1E56
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":23B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":290E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":2E10
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":336C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":39E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":3EE8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilButtonsHot 
      Left            =   8280
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":473A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":4C96
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":51F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":574E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":5CAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":6206
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":6708
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":6C64
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":72DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":77E0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilAddressEnabled 
      Left            =   7560
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   17
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":8432
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilAddressDisabled 
      Left            =   7560
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   17
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":87B4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin PicClip.PictureClip pcWorking 
      Left            =   6360
      Top             =   1200
      _ExtentX        =   4445
      _ExtentY        =   4763
      _Version        =   393216
      Rows            =   6
      Cols            =   6
      Picture         =   "frmBrowser.frx":8B36
   End
   Begin ComCtl3.CoolBar cbToolBar 
      Align           =   1  'Align Top
      Height          =   825
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   1455
      _CBWidth        =   10245
      _CBHeight       =   825
      _Version        =   "6.7.8988"
      Child1          =   "tbaNavigate"
      MinHeight1      =   390
      Width1          =   4605
      Key1            =   "Navigate"
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      BandBackColor2  =   -2147483630
      Child2          =   "pbWorking"
      MinHeight2      =   390
      UseCoolbarColors2=   0   'False
      UseCoolbarPicture2=   0   'False
      Key2            =   "Working"
      NewRow2         =   0   'False
      BandStyle2      =   1
      AllowVertical2  =   0   'False
      Child3          =   "pbAddress"
      MinHeight3      =   345
      Width3          =   1920
      Key3            =   "Address"
      NewRow3         =   -1  'True
      AllowVertical3  =   0   'False
      BandTag3        =   "Address"
      Begin MSComctlLib.Toolbar tbaNavigate 
         Height          =   390
         Left            =   165
         TabIndex        =   7
         Top             =   30
         Width           =   9960
         _ExtentX        =   17568
         _ExtentY        =   688
         ButtonWidth     =   714
         ButtonHeight    =   688
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilButtonsDisabled"
         DisabledImageList=   "ilButtonsDisabled"
         HotImageList    =   "ilButtonsHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   11
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Back"
               Description     =   "Back"
               Object.ToolTipText     =   "Back"
               ImageIndex      =   1
               Style           =   5
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Forward"
               Description     =   "Forward"
               Object.ToolTipText     =   "Forward"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "NavigateSeparator1"
               Style           =   3
               Object.Width           =   500
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Stop"
               Description     =   "Stop"
               Object.ToolTipText     =   "Stop"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Refresh"
               Description     =   "Refresh"
               Object.ToolTipText     =   "Refresh"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Home"
               Description     =   "Home"
               Object.ToolTipText     =   "Home"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "NavigateSeparator2"
               Style           =   3
               Object.Width           =   500
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Search"
               Description     =   "Search"
               Object.ToolTipText     =   "Search Internet"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Print"
               Description     =   "Print"
               Object.ToolTipText     =   "Print"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "NavigateSeparator3"
               Style           =   3
               Object.Width           =   500
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "RunScript"
               Description     =   "Run Scipt file"
               Object.ToolTipText     =   "Run Scipt file (""Script.txt"")"
               ImageIndex      =   10
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox pbWorking 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   10215
         ScaleHeight     =   390
         ScaleWidth      =   15
         TabIndex        =   6
         Top             =   30
         Width           =   15
      End
      Begin VB.PictureBox pbAddress 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   165
         ScaleHeight     =   345
         ScaleWidth      =   9990
         TabIndex        =   3
         Top             =   450
         Width           =   9990
         Begin VB.ComboBox cboAddress 
            Height          =   315
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   9495
         End
         Begin MSComctlLib.Toolbar tbAddress 
            Height          =   345
            Left            =   9600
            TabIndex        =   4
            Top             =   0
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   609
            ButtonWidth     =   1111
            ButtonHeight    =   609
            AllowCustomize  =   0   'False
            Wrappable       =   0   'False
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ilAddressDisabled"
            DisabledImageList=   "ilAddressDisabled"
            HotImageList    =   "ilAddressEnabled"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   1
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Go "
                  Key             =   "Go"
                  Object.ToolTipText     =   "Go to URL address"
                  ImageIndex      =   1
               EndProperty
            EndProperty
         End
      End
   End
   Begin SHDocVwCtl.WebBrowser wbBrowser 
      Height          =   2295
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   4695
      ExtentX         =   8281
      ExtentY         =   4048
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbDontNavigateNow As Boolean
Private mbNeedRefresh As Boolean

Dim WithEvents HTMLDoc As HTMLDocument
Attribute HTMLDoc.VB_VarHelpID = -1

Private Sub cboAddress_Change()

    'Set Go button tool tip to address
    tbAddress.Buttons("Go").ToolTipText = "Go To " & cboAddress.Text
    
End Sub

Private Sub cboAddress_Click()
    
    'Wait until browser control ready
    If mbDontNavigateNow Then Exit Sub
    'When ready navigate
    wbBrowser.Navigate cboAddress.Text
    
End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    
    'Wait until enter key is hit
    If KeyAscii = vbKeyReturn Then cboAddress_Click
    
End Sub


Private Sub cbToolBar_HeightChanged(ByVal NewHeight As Single)

    On Error Resume Next
    
    'Resize browser control
    If wbBrowser.StatusBar Then
        wbBrowser.Height = Me.ScaleHeight - NewHeight - sbStatus.Height
    Else
        wbBrowser.Height = Me.ScaleHeight - NewHeight
    End If

End Sub

Private Sub cbToolBar_Resize()

    'Resize address box and move go button
    tbAddress.Width = tbAddress.Buttons("Go").Width
    tbAddress.Left = pbAddress.Width - tbAddress.Width
    cboAddress.Width = pbAddress.Width - tbAddress.Width

End Sub

Private Sub Form_Load()

    Dim iCount As Long
    
    'Hide script commands window
    lstCommands.Visible = False
    
    'Set up working animaiton
    pbWorking.ScaleMode = vbPixels
    pbWorking.Width = cbToolBar.Bands("Working").Width
    pbWorking.Height = cbToolBar.Bands("Working").Height
    pbWorking.Picture = pcWorking.GraphicCell(0)
    cbToolBar.Bands("Working").MinWidth = 735

End Sub

Private Sub Form_Resize()
    
    Dim iSizing As Integer
    
    On Error Resume Next
    
    'Init sizing
    iSizing = 0
    
    'Add Cool bar
    If cbToolBar.Visible Then
        iSizing = iSizing + cbToolBar.Height
    End If
    
    If lstCommands.Visible Then
        'Posistion command list box
        lstCommands.Top = iSizing
        lstCommands.Left = 0
        lstCommands.Width = Me.ScaleWidth
        'Add command list box to sizing
        iSizing = iSizing + lstCommands.Height
    End If
    
    'Posistion browser
    wbBrowser.Top = iSizing
    wbBrowser.Left = 0
    wbBrowser.Width = Me.ScaleWidth
    
    'Add status bar
    If sbStatus.Visible Then
        iSizing = iSizing + sbStatus.Height
    End If
    
    'Set browser height
    wbBrowser.Height = Me.ScaleHeight - iSizing
    
    'Refresh finished
    mbNeedRefresh = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Dim iCount As Integer
    'unload sub forms
    For iCount = Forms.Count - 1 To 1 Step -1
        Unload Forms(iCount)
    Next

End Sub

Private Sub tbAddress_ButtonClick(ByVal Button As MSComctlLib.Button)

    On Error Resume Next
    
    'Do command on button
    Select Case Button.Key
    Case "Go"
        wbBrowser.Navigate cboAddress.Text
    End Select

End Sub

Private Sub tbaNavigate_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    On Error Resume Next
    
    'Do command on button
    Select Case Button.Key
    Case "Back"
        wbBrowser.GoBack
    Case "Forward"
        wbBrowser.GoForward
    Case "Stop"
        wbBrowser.Stop
    Case "Refresh"
        wbBrowser.Refresh
    Case "Home"
        'Navigate home
        wbBrowser.GoHome
    Case "Search"
        wbBrowser.GoSearch
    Case "Print"
        PrintWebPage False
    Case "RunScript"
        'Clear command list box
        lstCommands.Clear
        'Run script
        RunScript "Script.txt"
    End Select

End Sub

Private Sub tmrWorking_Timer()

    'Static variable rotate counter
    Static iRotate As Integer
    
    'Check for clip control cell
    If iRotate = 35 Then
        iRotate = 0
    Else
        iRotate = iRotate + 1
    End If
    
    'Set clip control cell as picture
    pbWorking.Picture = pcWorking.GraphicCell(iRotate)
    
End Sub

Private Sub wbBrowser_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)

    'Reset HTMLDoc
    Set HTMLDoc = Nothing
    
End Sub

Private Sub wbBrowser_DocumentComplete(ByVal pDisp As Object, URL As Variant)

    Dim htm As IHTMLDocument2
    
    'Set htm to browser document
    On Error Resume Next
    Set htm = wbBrowser.Document
    Set HTMLDoc = htm
    
End Sub

Private Sub wbBrowser_DownloadBegin()

    'Start working animation
    tmrWorking.Enabled = True

End Sub

Private Sub wbBrowser_DownloadComplete()

    'Stop working animation
    tmrWorking.Enabled = False

End Sub

Private Sub wbBrowser_NavigateComplete2(ByVal pDisp As Object, URL As Variant)

    Dim i As Integer
    Dim bFound As Boolean

    'Wait to renavigate
    mbDontNavigateNow = True

    'Add URL to address list
    For i = 0 To cboAddress.ListCount - 1
        If cboAddress.List(i) = wbBrowser.LocationURL Then
            bFound = True
            Exit For
        End If
    Next i
        
    'If the address is on list remove
    If bFound Then
        cboAddress.RemoveItem i
    End If
    
    'Add URL to top of list
    cboAddress.AddItem wbBrowser.LocationURL, 0
    cboAddress.ListIndex = 0

    'Renavigate OK
    mbDontNavigateNow = False
    'Set focus to browser window
    wbBrowser.SetFocus
    
    'This resize the form just in case the web page called SetTop or SetLeft
    If mbNeedRefresh Then Form_Resize
    
    DoEvents

End Sub

Private Sub wbBrowser_NewWindow2(ppDisp As Object, Cancel As Boolean)
    
    'Dim new form
    Dim frmWB As frmBrowser
    Set frmWB = New frmBrowser
        
    'Register as top level browser
    frmWB.wbBrowser.RegisterAsBrowser = True
        
    'Set form as new browser
    Set ppDisp = frmWB.wbBrowser.Object
        
    'Set window to normal
    'frmWB.WindowState = vbNormal
        
    'Show new form
    frmWB.Visible = True
    
End Sub

Private Sub wbBrowser_OnStatusBar(ByVal StatusBar As Boolean)

    'Hide status bar
    sbStatus.Visible = StatusBar
    Form_Resize
    
End Sub

Private Sub wbBrowser_OnToolBar(ByVal ToolBar As Boolean)
    
    'Hide toolbar
    cbToolBar.Visible = ToolBar
    Form_Resize
    
End Sub

Private Sub wbBrowser_TitleChange(ByVal Text As String)

    'Add URL to title
    Me.Caption = Text & " - " & Me.Caption

End Sub

Private Sub wbBrowser_WindowClosing(ByVal IsChildWindow As Boolean, Cancel As Boolean)

    'Check if child window and close
    If IsChildWindow Then
        Cancel = True
        Unload Me
    'Else warn closing
    Else
        Cancel = True
        If (MsgBox("The web page you are viewing is trying to close the window." & vbCrLf & vbCrLf & "Do you want to close this window?", vbYesNo, "Public Web Browser")) = vbYes Then
            Unload Me
        End If
    End If

End Sub

Private Sub wbBrowser_WindowSetHeight(ByVal Height As Long)
    
    Dim t_Height As Integer
    
    'Get differnces between heights
    t_Height = Me.Height - Me.ScaleHeight
    'Set new height
    Me.Height = (Height * Screen.TwipsPerPixelY) + t_Height

End Sub

Private Sub wbBrowser_WindowSetLeft(ByVal Left As Long)

    'Set window left
    Me.Left = Left
    'Refresh the screen
    mbNeedRefresh = True

End Sub

Private Sub wbBrowser_WindowSetResizable(ByVal Resizable As Boolean)

    If Resizable Then
        Me.BorderStyle = 2
    Else
        Me.BorderStyle = 3
    End If

End Sub

Private Sub wbBrowser_WindowSetTop(ByVal Top As Long)

    'Set window top
    Me.Top = Top
    'Refresh the screen
    mbNeedRefresh = True
    
End Sub

Private Sub wbBrowser_StatusTextChange(ByVal Text As String)

    'Mouse over status bar change
    If Len(Text) Then
        sbStatus.Panels("URL").Text = Text
    Else
        sbStatus.Panels("URL").Text = wbBrowser.LocationName
    End If
    
End Sub

Private Sub wbBrowser_WindowSetWidth(ByVal Width As Long)

    Dim t_Width As Integer
    
    'Get differnces between Widths
    t_Width = Me.Width - Me.ScaleWidth
    'Set new Width
    Me.Width = (Width * Screen.TwipsPerPixelY) + t_Width

End Sub

Public Function PrintWebPage(ShowDialog As Boolean) As Boolean

    On Error Resume Next
    Dim eQuery As OLECMDF
           
    'Get print command status
    eQuery = wbBrowser.QueryStatusWB(OLECMDID_PRINT)
    If Err.Number = 0 Then
        If eQuery And OLECMDF_ENABLED Then
            'Check for dialog then Print
            If ShowDialog Then
                wbBrowser.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER
            Else
                wbBrowser.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER
            End If
            PrintWebPage = True
        Else
            'Notify printing is disabled
            MsgBox "The Print command is currently disabled."
            PrintWebPage = False
        End If
    End If
    If Err.Number <> 0 Then MsgBox "Print command Error: " & Err.Description
    
End Function

Private Function RunScript(FileName As String) As Boolean

    Dim iOpenFile As Integer
    Dim sCommand As String
    
    'Show commands list
    lstCommands.Visible = True
    Form_Resize
    
    'Open file
    iOpenFile = FreeFile
    Open FileName For Input Shared As #iOpenFile
    
    'Get line from command file
    Do While Not EOF(iOpenFile)
    
        Line Input #iOpenFile, sCommand
        If Len(sCommand) Then
            'Add to command list box
            lstCommands.AddItem sCommand
            'Do command
            ParseCode sCommand
        End If
        'Wait until result are loaded
        Do
            DoEvents
            sbStatus.Panels("Information").Text = "Waiting"
        Loop Until Not wbBrowser.Busy
        sbStatus.Panels("Information").Text = "Running"
    Loop
    
    'Close file
    Close iOpenFile
    'Show ready status
    sbStatus.Panels("Information").Text = "Ready"
    'Hide commands list
    lstCommands.Visible = False
    Form_Resize
    'Success return true
    RunScript = True
    Exit Function
    
ErrorTrap:
    Select Case Err.Number
        Case 53 'File not found
            MsgBox "File not found. Please check the Ini file for appropiate path.", vbOKOnly, FileName
        Case Else
            'do NOT use the stop statement except for testing purposes.
            MsgBox "An Error has occured. Please verify all files and paths specified in Ini file.", vbOKOnly, FileName
    End Select
    
End Function
