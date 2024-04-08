VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frmLogViewer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Request Log"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15000
   Icon            =   "frmLogViewer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   15000
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab 
      Height          =   8295
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   14631
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Query for Requests"
      TabPicture(0)   =   "frmLogViewer.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdCopySdgToClipboard"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "CmbStatus"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "CmbAppCodeNot"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "CmbAppCodePassed"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cldrFrom"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "imlsSdg"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "btnSdgGo"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cldrTo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lvSdg"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "btnToDate"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "edtToDate"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "btnFromDate"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "edtFromDate"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label5"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label4"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label3"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label2"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Logs for a Request / Work Station"
      TabPicture(1)   =   "frmLogViewer.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "LblComputerName"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label7"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label8"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "edtSdgName"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "btnLogGo"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lvLog"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Option1"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Option2"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "CmbComputerName"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "CmdToDate"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "TxtToDate"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "CmdFromDate"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "TxtFromDate"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "CalendarToDate"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "CalendarFromDate"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "cmdCopyToClipboard"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).ControlCount=   17
      Begin VB.CommandButton cmdCopySdgToClipboard 
         Caption         =   "Copy To Clipboard"
         Height          =   456
         Left            =   -61320
         TabIndex        =   34
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdCopyToClipboard 
         Caption         =   "Copy To Clipboard"
         Height          =   285
         Left            =   13080
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   480
         Width           =   1455
      End
      Begin MSACAL.Calendar CalendarFromDate 
         Height          =   2535
         Left            =   7320
         TabIndex        =   31
         Top             =   1080
         Width           =   3615
         _Version        =   524288
         _ExtentX        =   6376
         _ExtentY        =   4471
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2005
         Month           =   9
         Day             =   6
         DayLength       =   1
         MonthLength     =   1
         DayFontColor    =   0
         FirstDay        =   7
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSACAL.Calendar CalendarToDate 
         Height          =   2535
         Left            =   10920
         TabIndex        =   32
         Top             =   1080
         Width           =   3615
         _Version        =   524288
         _ExtentX        =   6376
         _ExtentY        =   4471
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2005
         Month           =   9
         Day             =   6
         DayLength       =   1
         MonthLength     =   1
         DayFontColor    =   0
         FirstDay        =   7
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox TxtFromDate 
         Height          =   360
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton CmdFromDate 
         Caption         =   "..."
         Height          =   285
         Left            =   9720
         TabIndex        =   27
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox TxtToDate 
         Height          =   360
         Left            =   10560
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton CmdToDate 
         Caption         =   "..."
         Height          =   285
         Left            =   11760
         TabIndex        =   25
         Top             =   480
         Width           =   375
      End
      Begin VB.ComboBox CmbComputerName 
         Height          =   360
         Left            =   5040
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   480
         Width           =   2175
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option1"
         Height          =   255
         Left            =   3480
         TabIndex        =   22
         Top             =   510
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   510
         Width           =   255
      End
      Begin VB.ComboBox CmbStatus 
         Height          =   360
         Left            =   -68520
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   795
         Width           =   1335
      End
      Begin VB.ComboBox CmbAppCodeNot 
         Height          =   360
         Left            =   -72600
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   975
         Width           =   3375
      End
      Begin VB.ComboBox CmbAppCodePassed 
         Height          =   360
         Left            =   -72600
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   555
         Width           =   3375
      End
      Begin MSComctlLib.ListView lvLog 
         Height          =   7215
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   14500
         _ExtentX        =   25585
         _ExtentY        =   12726
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Time"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Request Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "App. Code"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Operator"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Description"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Computer Name"
            Object.Width           =   4851
         EndProperty
      End
      Begin VB.CommandButton btnLogGo 
         BackColor       =   &H80000018&
         Caption         =   "Go"
         Height          =   285
         Left            =   12240
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   480
         Width           =   750
      End
      Begin VB.TextBox edtSdgName 
         Height          =   360
         Left            =   1680
         TabIndex        =   0
         Top             =   480
         Width           =   1695
      End
      Begin MSACAL.Calendar cldrFrom 
         Height          =   2415
         Left            =   -67560
         TabIndex        =   13
         Top             =   1560
         Width           =   3495
         _Version        =   524288
         _ExtentX        =   6165
         _ExtentY        =   4260
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2005
         Month           =   9
         Day             =   6
         DayLength       =   1
         MonthLength     =   1
         DayFontColor    =   0
         FirstDay        =   7
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ImageList imlsSdg 
         Left            =   -74640
         Top             =   7440
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin VB.CommandButton btnSdgGo 
         Caption         =   "Go"
         Default         =   -1  'True
         Height          =   285
         Left            =   -62160
         TabIndex        =   15
         Top             =   840
         Width           =   750
      End
      Begin MSACAL.Calendar cldrTo 
         Height          =   2415
         Left            =   -63960
         TabIndex        =   14
         Top             =   1560
         Width           =   3495
         _Version        =   524288
         _ExtentX        =   6165
         _ExtentY        =   4260
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2005
         Month           =   9
         Day             =   6
         DayLength       =   1
         MonthLength     =   1
         DayFontColor    =   0
         FirstDay        =   7
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ListView lvSdg 
         Height          =   6735
         Left            =   -74880
         TabIndex        =   12
         Top             =   1440
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   11880
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imlsSdg"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "name"
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton btnToDate 
         Caption         =   "..."
         Height          =   285
         Left            =   -62640
         TabIndex        =   9
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox edtToDate 
         Height          =   360
         Left            =   -63840
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton btnFromDate 
         Caption         =   "..."
         Height          =   285
         Left            =   -64680
         TabIndex        =   7
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox edtFromDate 
         Height          =   360
         Left            =   -65880
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Between Dates:"
         Height          =   195
         Left            =   7320
         TabIndex        =   30
         Top             =   510
         Width           =   1140
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "and:"
         Height          =   195
         Left            =   10185
         TabIndex        =   29
         Top             =   510
         Width           =   315
      End
      Begin VB.Label LblComputerName 
         AutoSize        =   -1  'True
         Caption         =   "Computer Name:"
         Height          =   195
         Left            =   3720
         TabIndex        =   24
         Top             =   510
         Width           =   1185
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Internal Number:"
         Height          =   195
         Left            =   390
         TabIndex        =   16
         Top             =   510
         Width           =   1170
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Didn't Pass Application Code:"
         Height          =   435
         Left            =   -74760
         TabIndex        =   11
         Top             =   1005
         Width           =   2085
      End
      Begin VB.Label Label4 
         Caption         =   "and:"
         Height          =   255
         Left            =   -64200
         TabIndex        =   10
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Between Dates:"
         Height          =   255
         Left            =   -67080
         TabIndex        =   5
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Passed Application Code:"
         Height          =   435
         Left            =   -74760
         TabIndex        =   4
         Top             =   645
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Status:"
         Height          =   195
         Left            =   -69120
         TabIndex        =   3
         Top             =   840
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmLogViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public con As ADODB.Connection
Public strOperatorId As String

Private StatusCodes As Scripting.Dictionary
Private AppCodes As Scripting.Dictionary
Private dicOperators As New Dictionary
Private dicLabManagers As New Dictionary

Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, _
   ByVal dwBytes As Long) As Long

Private Declare Function CloseClipboard Lib "User32" () As Long

Private Declare Function OpenClipboard Lib "User32" (ByVal hwnd As Long) As Long

Private Declare Function EmptyClipboard Lib "User32" () As Long

Private Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, _
   ByVal lpString2 As Any) As Long

Private Declare Function SetClipboardData Lib "User32" (ByVal wFormat _
   As Long, ByVal hMem As Long) As Long

Private Const GHND = &H42
Private Const CF_TEXT = 1
Private Const MAXSIZE = 4096


Private Sub btnFromDate_Click()
220       If Not cldrFrom.Visible Then
230           cldrFrom.Visible = True
240           cldrFrom.Value = Now
250       Else
260           cldrFrom.Visible = False
270       End If
End Sub

Private Sub CalendarFromDate_Click()
280       TxtFromDate.Text = CalendarFromDate.Value
290       CalendarFromDate.Visible = False
End Sub

Private Sub CalendarToDate_Click()
300       TxtToDate.Text = CalendarToDate.Value
310       CalendarToDate.Visible = False
End Sub

Private Sub cldrFrom_Click()
320       edtFromDate.Text = cldrFrom.Value
330       cldrFrom.Visible = False
End Sub

Private Sub GetLogs4Workstation()
340       On Error GoTo ErrEnd
          Dim rst As ADODB.Recordset
          Dim sqlstr As String
          Dim li As ListItem

350       If CmbComputerName.Text = "None" Then
360           CmbComputerName.BackColor = vbRed
370           MsgBox "Computer Name must be selected from list !", vbCritical, "Nautilus - Request Log"
380           CmbComputerName.BackColor = vbWhite
390           Call CmbComputerName.SetFocus
400           Exit Sub
410       End If

420       If Trim(TxtFromDate.Text) = "" Then
430           TxtFromDate.BackColor = vbRed
440           MsgBox "From date must be choosed !", vbCritical, "Nautilus - Request Log"
450           TxtFromDate.BackColor = vbWhite
460           Call TxtFromDate.SetFocus
470           Exit Sub
480       End If

490       If Trim(TxtToDate.Text) = "" Then
500           TxtToDate.BackColor = vbRed
510           MsgBox "To date must be choosed !", vbCritical, "Nautilus - Request Log"
520           TxtToDate.BackColor = vbWhite
530           Call TxtToDate.SetFocus
540           Exit Sub
550       End If

          
560       MousePointer = vbHourglass

570       sqlstr = "select to_char(l.time, 'dd/mm/yy hh24:mi') time, " & _
                      "l.application_code, " & _
                      "'(' || o.operator_id || ') ' || o.full_name op, " & _
                      "l.description, " & _
                      "d.name request_name, " & _
                      "s.terminal_name " & _
                   "from lims_sys.sdg_log l, " & _
                      "lims_sys.sdg d, " & _
                      "lims_sys.lims_session s, " & _
                      "lims_sys.operator o " & _
                   "where s.terminal_name = '" & CmbComputerName.Text & "' " & _
                      "and l.session_id = s.session_id(+) " & _
                      "and l.application_code not like 'BOX%' " & _
                      "and s.operator_id = o.operator_id(+) " & _
                      "and l.sdg_id = d.sdg_id " & _
                      "and trunc(l.time,'ddd') >= to_date('" & TxtFromDate.Text & "','dd/mm/yyyy') " & _
                      "and trunc(l.time,'ddd') <= to_date('" & TxtToDate.Text & "','dd/mm/yyyy')+1 " & _
                   "order by l.time"
                      
          'if the current user is the lab manager,
          'cut by the selected operator's id:
580       If dicLabManagers.Exists(strOperatorId) Then
          
                           
590           sqlstr = " SELECT   TO_CHAR (l.TIME, 'dd/mm/yy hh24:mi') TIME, l.application_code,"
600           sqlstr = sqlstr & "          '(' || o.operator_id || ') ' || o.full_name op, l.description,"
610           sqlstr = sqlstr & "          d.NAME request_name, s.terminal_name"
620           sqlstr = sqlstr & "     FROM  lims_sys.sdg_log l,  lims_sys.lims_session s,  lims_sys.OPERATOR o, lims_sys.sdg d"
630           sqlstr = sqlstr & "    WHERE l.TIME BETWEEN TO_DATE ('" & TxtFromDate.Text & "','dd/mm/yyyy') "
640           sqlstr = sqlstr & "                     AND TO_DATE ('" & TxtToDate.Text & "','dd/mm/yyyy') +1"
650           sqlstr = sqlstr & "      AND o.operator_id = " & dicOperators(CmbComputerName.Text)
660           sqlstr = sqlstr & "      AND s.operator_id = o.operator_id(+)"
670           sqlstr = sqlstr & "      AND l.session_id = s.session_id(+)"
680           sqlstr = sqlstr & "        AND l.application_code NOT LIKE 'BOX%'"
690           sqlstr = sqlstr & "      AND l.sdg_id = d.sdg_id"
700           sqlstr = sqlstr & " ORDER BY l.TIME"
                               
                       
710       End If
          
720       Set rst = con.Execute(sqlstr)
730       lvLog.ListItems.Clear
740       While Not rst.EOF
750           Set li = lvLog.ListItems.Add(, , nte(rst("TIME")))
760           li.SubItems(1) = nte(rst("request_name"))
770           li.SubItems(2) = TranslateAppCode(nte(rst("APPLICATION_CODE")))
780           li.SubItems(3) = nte(rst("OP"))
790           li.SubItems(4) = nte(rst("DESCRIPTION"))
800           li.SubItems(5) = nte(rst("TERMINAL_NAME"))
810           rst.MoveNext
820       Wend
          
830       MousePointer = vbDefault
          
840       Exit Sub

ErrEnd:
850       MousePointer = vbDefault
          
860       MsgBox "GetLogs4Workstation... " & vbCrLf & _
                  "line # " & Erl & vbCrLf & Err.Description
End Sub

Private Sub GetLogs4Sdg()
870       On Error GoTo ErrEnd
          Dim rst As ADODB.Recordset
          Dim sqlstr As String
          Dim li As ListItem
          Dim GetAllSdgs As ADODB.Recordset
          Dim getSDgSql As String
          Dim sdgName As String

880       If Trim(edtSdgName.Text) = "" Then
890           edtSdgName.BackColor = vbRed
900           MsgBox "Request Name must be entered !", vbCritical, "Nautilus - Request Log"
910           edtSdgName.BackColor = vbWhite
920           Call edtSdgName.SetFocus
930           Exit Sub
940       End If

          
          
          'ashi - 954

950       If InStr(edtSdgName.Text, ".") > 0 Then
960           edtSdgName.Text = Left(edtSdgName.Text, InStr(edtSdgName.Text, ".") - 1)
970       End If
          
980       getSDgSql = "select name from lims_sys.sdg ,lims_sys.sdg_user where " _
                & " ( sdg.name='" & edtSdgName.Text & "' " _
                & " or sdg_user.U_PATHOLAB_NUMBER= '" & edtSdgName.Text & "' ) " _
                & " and sdg.SDG_ID=sdg_user.SDG_ID "
990       Set GetAllSdgs = con.Execute(getSDgSql)
1000      If Not GetAllSdgs.EOF Then
1010          If Trim(nte(GetAllSdgs(0))) <> "" Then
1020              sdgName = nte(GetAllSdgs(0))
1030          End If
1040      End If
          
1050      GetAllSdgs.Close
         '----
1060  MousePointer = vbHourglass
1070         sqlstr = "select to_char(l.time, 'dd/mm/yy hh24:mi') time, " & _
                      "l.application_code, " & _
                      "'(' || o.operator_id || ') ' || o.full_name op, " & _
                      "l.description, " & _
                      "d.name request_name, " & _
                      "s.terminal_name " & _
                   "from lims_sys.sdg_log l, " & _
                      "lims_sys.sdg d, " & _
                      "lims_sys.lims_session s, " & _
                      "lims_sys.operator o " & _
                   "where l.sdg_id = d.sdg_id and d.name = '" & sdgName & "'" & _
                      "and l.session_id = s.session_id(+) " & _
                      "and l.application_code not like 'BOX%' " & _
                      "and s.operator_id = o.operator_id(+) " & _
                   "order by l.time"
                
1080      Set rst = con.Execute(sqlstr)
1090      lvLog.ListItems.Clear
1100      While Not rst.EOF
1110          Set li = lvLog.ListItems.Add(, , nte(rst("TIME")))
1120          li.SubItems(1) = nte(rst("request_name"))
1130          li.SubItems(2) = TranslateAppCode(nte(rst("APPLICATION_CODE")))
1140          li.SubItems(3) = nte(rst("OP"))
1150          li.SubItems(4) = nte(rst("DESCRIPTION"))
1160          li.SubItems(5) = nte(rst("TERMINAL_NAME"))
1170          rst.MoveNext
1180      Wend
          
1190      MousePointer = vbDefault
          
1200      Exit Sub

ErrEnd:
1210      MousePointer = vbDefault
1220      MsgBox "GetLogs4Sdg... " & vbCrLf & _
                  "line # " & Erl & vbCrLf & Err.Description
End Sub

Private Function TranslateAppCode(strAppCode As String) As String
          Dim PhraseRs As ADODB.Recordset
          Dim strSQL As String

1230      TranslateAppCode = Trim(strAppCode)

1240      strSQL = "select phrase_description from lims_sys.phrase_entry " & _
              "where phrase_id = (select phrase_id from lims_sys.phrase_header where " & _
              "name = 'Sdg-log names') " & _
              "and phrase_name = '" & strAppCode & "'"

1250      Set PhraseRs = con.Execute(strSQL)
1260      If Not PhraseRs.EOF Then
1270          If Trim(nte(PhraseRs(0))) <> "" Then
1280              TranslateAppCode = nte(PhraseRs(0))
1290          End If
1300      End If
End Function

Private Sub btnLogGo_Click()
1310      If Option1.Value = True Then
1320          GetLogs4Sdg
1330      Else
1340          GetLogs4Workstation
1350      End If
End Sub

Private Sub btnToDate_Click()
1360      If Not cldrTo.Visible Then
1370          cldrTo.Visible = True
1380          cldrTo.Value = Now
1390      Else
1400          cldrTo.Visible = False
1410      End If
End Sub

Private Sub cldrTo_Click()
1420      edtToDate.Text = cldrTo.Value
1430      cldrTo.Visible = False
End Sub

Private Sub cmdCopySdgToClipboard_Click()
1440   On Error GoTo Err_cmdCopySdgToClipboard
          Dim strFields As String
          Dim tmpStr As String
          Dim strHeader As String
          Dim i, j  As Integer
          Dim li As ListItem
          
1450      strFields = ""
1460      strHeader = ""
          
1470      For i = 1 To lvSdg.ListItems.Count
          
1480          strFields = strFields & Trim(lvSdg.ListItems(i).Text) & vbCrLf
1490      Next i
          
1500      If Trim(strFields) <> "" Then
1510          strHeader = Trim(lvSdg.ColumnHeaders.Item(1))
1520          strFields = strHeader & vbCrLf & strFields
1530          Call ClipBoard_SetData(strFields)
1540          MsgBox "The information has been successfuly copied to the clipboard.", _
                      vbInformation + vbYes, "Nautilus - Copy To Clipboard"
1550      End If
          
              
1560      Exit Sub
Err_cmdCopySdgToClipboard:
1570     MsgBox " Error on cmdCopySdgToClipboard" & vbCrLf & "line # " & Erl & vbCrLf & Err.Description, vbOKOnly, "Error"

End Sub

Private Sub cmdCopyToClipboard_Click()
1580     On Error GoTo Err_cmdCopyToClipboard
          Dim strFields As String
          Dim tmpStr As String
          Dim strHeader As String
          Dim i, j  As Integer
          Dim li As ListItem
          
1590      strFields = ""
1600      strHeader = ""
          
1610      For i = 1 To lvLog.ListItems.Count
          
1620          strFields = strFields & Trim(lvLog.ListItems(i).Text)
1630          Set li = lvLog.ListItems(i)
              
1640          For j = 1 To lvLog.ColumnHeaders.Count - 1 'sub items
1650              tmpStr = Trim(li.SubItems(j))
1660              strFields = strFields & vbTab & tmpStr
1670          Next j
1680          strFields = strFields & vbCrLf
1690      Next i
          
          
         

1700      If Trim(strFields) <> "" Then
          
1710       For i = 1 To lvLog.ColumnHeaders.Count
1720          strHeader = strHeader & IIf(i = 1, "", vbTab) & Trim(lvLog.ColumnHeaders.Item(i))
1730      Next i
          
1740          strFields = strHeader & vbCrLf & strFields
1750          Call ClipBoard_SetData(strFields)
1760          MsgBox "The information has been successfuly copied to the clipboard.", _
                      vbInformation + vbYes, "Nautilus - Copy To Clipboard"
1770      End If
          
              
1780      Exit Sub
Err_cmdCopyToClipboard:
1790     MsgBox " Error on cmdCopyToClipboard" & vbCrLf & "line # " & Erl & vbCrLf & Err.Description, vbOKOnly, "Error"
End Sub

Function ClipBoard_SetData(MyString As String)
         Dim hGlobalMemory As Long, lpGlobalMemory As Long
         Dim hClipMemory As Long, X As Long

         ' Allocate moveable global memory.
         '-------------------------------------------
1800     hGlobalMemory = GlobalAlloc(GHND, Len(MyString) + 1)

         ' Lock the block to get a far pointer
         ' to this memory.
1810     lpGlobalMemory = GlobalLock(hGlobalMemory)

         ' Copy the string to this global memory.
1820     lpGlobalMemory = lstrcpy(lpGlobalMemory, MyString)

         ' Unlock the memory.
1830     If GlobalUnlock(hGlobalMemory) <> 0 Then
1840        MsgBox "Could not unlock memory location. Copy aborted."
1850        GoTo OutOfHere2
1860     End If

         ' Open the Clipboard to copy data to.
1870     If OpenClipboard(0&) = 0 Then
1880        MsgBox "Could not open the Clipboard. Copy aborted."
1890        Exit Function
1900     End If

         ' Clear the Clipboard.
1910     X = EmptyClipboard()

         ' Copy the data to the Clipboard.
1920     hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)

OutOfHere2:

1930     If CloseClipboard() = 0 Then
1940        MsgBox "Could not close Clipboard."
1950     End If
End Function

Private Sub CmdFromDate_Click()
1960      If Not CalendarFromDate.Visible Then
1970          CalendarFromDate.Visible = True
1980          CalendarFromDate.Value = Now
1990      Else
2000          CalendarFromDate.Visible = False
2010      End If
End Sub

Private Sub CmdToDate_Click()
2020      If Not CalendarToDate.Visible Then
2030          CalendarToDate.Visible = True
2040          CalendarToDate.Value = Now
2050      Else
2060          CalendarToDate.Visible = False
2070      End If
End Sub

Private Sub edtFromDate_KeyDown(KeyCode As Integer, Shift As Integer)
2080      If KeyCode = vbKeyDelete Then
2090          edtFromDate.Text = ""
2100      End If
End Sub

Private Sub edtSdgName_KeyDown(KeyCode As Integer, Shift As Integer)

2110      If Not KeyCode = vbKeyReturn Then Exit Sub
2120      If Trim(edtSdgName.Text) = "" Then Exit Sub
2130      GetLogs4Sdg
End Sub

Private Sub edtToDate_KeyDown(KeyCode As Integer, Shift As Integer)
2140      If KeyCode = vbKeyDelete Then
2150          edtToDate.Text = ""
2160      End If
End Sub

Private Sub Form_Activate()
2170      Option1.Value = True
          'SSTab.Tab = 1
2180      edtFromDate.Text = Format(Now - 30, "dd/mm/yyyy")
2190      edtToDate.Text = Format(Now, "dd/mm/yyyy")
2200      SSTab.Tab = 1
      '    Call edtSdgName.SetFocus
          
          'so ENTER will activate search by SDG / Computer name:
2210      btnLogGo.Default = True
2220      Call edtSdgName.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
          Dim strVer As String

2230      If KeyCode = vbKeyF10 And Shift = 1 Then
2240          strVer = "Name: " & App.EXEName & vbCrLf & vbCrLf & _
                       "Path: " & App.Path & vbCrLf & vbCrLf & _
                       "Version: " & "[" & App.Major & "." & App.Minor & "." & App.Revision & "]" & vbCrLf & vbCrLf & _
                       "Company: One Software Technologies (O.S.T) Ltd."
2250          MsgBox strVer, vbInformation, "Nautilus - Project Properties"
2260      End If

2270      If KeyCode = vbKeyEscape Then
2280          Unload Me
2290      End If
End Sub

Private Sub Form_Load()
2300  On Error GoTo ERR_Form_Load
          
2310      cldrFrom.Visible = False
2320      cldrTo.Visible = False
2330      CalendarFromDate.Visible = False
2340      CalendarToDate.Visible = False
       
2350      Call imlsSdg.ListImages.Add(, "U", LoadPicture("C:\Program Files (x86)\Thermo\Nautilus\Resource\sdgu.ico"))
2360      Call imlsSdg.ListImages.Add(, "V", LoadPicture("C:\Program Files (x86)\Thermo\Nautilus\Resource\sdgv.ico"))
2370      Call imlsSdg.ListImages.Add(, "P", LoadPicture("C:\Program Files (x86)\Thermo\Nautilus\Resource\sdgp.ico"))
2380      Call imlsSdg.ListImages.Add(, "C", LoadPicture("C:\Program Files (x86)\Thermo\Nautilus\Resource\sdgc.ico"))
2390      Call imlsSdg.ListImages.Add(, "A", LoadPicture("C:\Program Files (x86)\Thermo\Nautilus\Resource\sdga.ico"))
2400      Call imlsSdg.ListImages.Add(, "X", LoadPicture("C:\Program Files (x86)\Thermo\Nautilus\Resource\sdgx.ico"))
2410      Call imlsSdg.ListImages.Add(, "S", LoadPicture("C:\Program Files (x86)\Thermo\Nautilus\Resource\sdgs.ico"))
2420      Call imlsSdg.ListImages.Add(, "R", LoadPicture("C:\Program Files (x86)\Thermo\Nautilus\Resource\sdgr.ico"))
2430      Call imlsSdg.ListImages.Add(, "I", LoadPicture("C:\Program Files (x86)\Thermo\Nautilus\Resource\sdgi.ico"))
        
2440      Call InitApplicationList
2450      Call InitStatusList
2460      Call InitLabManagersList
          
2470      If dicLabManagers.Exists(strOperatorId) Then
        
2480          Call InitWorkersList
2490      Else
2500          Call InitWorkStationList
            
2510      End If
          
2520      Exit Sub
ERR_Form_Load:
2530  MsgBox "ERR_Form_Load" & vbCrLf & "line # " & Erl & vbCrLf & Err.Description
End Sub

'if the current user is NOT the lab manager:
Private Sub InitWorkStationList()
2540  On Error GoTo ERR_InitWorkStationList
          
          
          
          Dim WSRs As ADODB.Recordset

          'Init the work station combo
2550      Set WSRs = con.Execute("select name " & _
              "from lims_sys.workstation " & _
              "order by name")

2560      CmbComputerName.Clear
2570      CmbComputerName.List(0) = "None"
2580      While Not WSRs.EOF
2590          CmbComputerName.List(CmbComputerName.ListCount) = _
                  nte(WSRs("NAME"))
2600          WSRs.MoveNext
2610      Wend
2620      WSRs.Close
2630      CmbComputerName.ListIndex = 0
          
2640      Exit Sub
ERR_InitWorkStationList:
2650  MsgBox "ERR_InitWorkStationList" & vbCrLf & "line # " & Erl & vbCrLf & Err.Description
End Sub

'if the current user is the lab manager:
Private Sub InitWorkersList()
2660  On Error GoTo ERR_InitWorkersList

          Dim WSRs As ADODB.Recordset
          Dim sql As String

          'Init the worker combo

2670      sql = " select o.NAME, o.OPERATOR_ID"
2680      sql = sql & " from lims_sys.operator o,"
2690      sql = sql & "      lims_sys.operator_user ou"
2700      sql = sql & " where ou.OPERATOR_ID=o.OPERATOR_ID"
2710      sql = sql & " and   ou.U_MANAGER_ORDER is not null"
2720      sql = sql & " order by ou.U_MANAGER_ORDER"
          
2730      Set WSRs = con.Execute(sql)
          
      '    Set WSRs = con.Execute(" select FULL_NAME, operator_id " & _
              " from lims_sys.operator " & _
              " order by FULL_NAME")
      '    Set WSRs = con.Execute(" select NAME, operator_id " & _
              " from lims_sys.operator " & _
              " order by operator_id")

2740      CmbComputerName.Clear
2750      CmbComputerName.List(0) = "None"
2760      While Not WSRs.EOF
2770          CmbComputerName.List(CmbComputerName.ListCount) = _
                  nte(WSRs("NAME"))
                  
2780          If Not dicOperators.Exists(nte(WSRs("NAME"))) Then
2790              Call dicOperators.Add(nte(WSRs("NAME")), nte(WSRs("operator_id")))
2800          End If
                  
2810          WSRs.MoveNext
2820      Wend
2830      WSRs.Close
2840      CmbComputerName.ListIndex = 0
          
2850      LblComputerName.Caption = "Worker Name:"
          
2860      Exit Sub
ERR_InitWorkersList:
2870  MsgBox "ERR_InitWorkersList" & vbCrLf & "line # " & Erl & vbCrLf & Err.Description
End Sub

Private Sub InitApplicationList()
2880  On Error GoTo ERR_InitApplicationList
          
          Dim AppCodeRs As ADODB.Recordset

          'Init the Status combo
2890      Set AppCodeRs = con.Execute("select phrase_description, phrase_name " & _
              "from lims_sys.phrase_entry " & _
              "where phrase_id = (select phrase_id from lims_sys.phrase_header where " & _
              "name = 'Sdg-log names') " & _
              "order by order_number")

2900      CmbAppCodePassed.Clear
2910      CmbAppCodeNot.Clear
2920      CmbAppCodePassed.List(0) = "None"
2930      CmbAppCodeNot.List(0) = "None"
2940      Set AppCodes = New Scripting.Dictionary
2950      While Not AppCodeRs.EOF
2960          CmbAppCodePassed.List(CmbAppCodePassed.ListCount) = AppCodeRs("PHRASE_DESCRIPTION")
2970          CmbAppCodeNot.List(CmbAppCodeNot.ListCount) = AppCodeRs("PHRASE_DESCRIPTION")
2980          Call AppCodes.Add(CStr(AppCodeRs("PHRASE_DESCRIPTION").Value), CStr(AppCodeRs("PHRASE_NAME").Value))
2990          AppCodeRs.MoveNext
3000      Wend
         
3010      AppCodeRs.Close
3020      CmbAppCodePassed.ListIndex = 0
3030      CmbAppCodeNot.ListIndex = 0
          
3040      Exit Sub
ERR_InitApplicationList:
3050  MsgBox "ERR_InitApplicationList" & vbCrLf & "line # " & Erl & vbCrLf & Err.Description & _
    vbCrLf & IIf(AppCodeRs.EOF, "", "App Code Description is double, change the description in phrase 'Sdg-log names' " & _
    vbCrLf & "Description:'" & CStr(AppCodeRs("PHRASE_DESCRIPTION").Value) & _
    "' Name:'" & CStr(AppCodeRs("PHRASE_NAME").Value) & "'")
End Sub

Private Sub InitStatusList()
3060  On Error GoTo ERR_InitStatusList
          
          Dim Status As ADODB.Recordset

          'Init the Status combo
3070      Set Status = con.Execute("select phrase_description, phrase_name " & _
              "from lims_sys.phrase_entry " & _
              "where phrase_id = (select phrase_id from lims_sys.phrase_header where " & _
              "name = 'SDG Status') " & _
              "order by order_number")

3080      CmbStatus.Clear
3090      CmbStatus.List(0) = "All"
3100      Set StatusCodes = New Scripting.Dictionary
3110      While Not Status.EOF
3120          CmbStatus.List(CmbStatus.ListCount) = Status("PHRASE_DESCRIPTION")
3130          Call StatusCodes.Add(CStr(Status("PHRASE_DESCRIPTION").Value), CStr(Status("PHRASE_NAME").Value))
3140          Status.MoveNext
3150      Wend
3160      Status.Close
3170      CmbStatus.ListIndex = 0
          
3180      Exit Sub
ERR_InitStatusList:
3190  MsgBox "ERR_InitStatusList" & vbCrLf & "line # " & Erl & vbCrLf & Err.Description
End Sub

'get all codes of those defined as the
'pathology lab managers:
Private Sub InitLabManagersList()
3200  On Error GoTo ERR_InitLabManagersList

          Dim rs As Recordset
          
3210      Set rs = con.Execute("select phrase_description, phrase_name " & _
              "from lims_sys.phrase_entry " & _
              "where phrase_id = (select phrase_id from lims_sys.phrase_header where " & _
              "name = 'PathologyLabManagers') " & _
              "order by order_number")

3220      Call dicLabManagers.RemoveAll

3230      While Not rs.EOF
3240          Call dicLabManagers.Add(nte(rs("phrase_description")), _
                                      nte(rs("phrase_name")))
              
3250          rs.MoveNext
3260      Wend
          
3270      Exit Sub
ERR_InitLabManagersList:
3280  MsgBox "ERR_InitLabManagersList" & vbCrLf & "line # " & Erl & vbCrLf & Err.Description
End Sub

Private Sub btnSdgGo_Click()

          
3290      On Error GoTo ErrEnd
          Dim rst As ADODB.Recordset
          Dim Where As String
          Dim li As ListItem
          Dim sqlstr As String

3300      cldrFrom.Visible = False
3310      cldrTo.Visible = False

3320      Where = ""

3330      If Trim(CmbStatus.Text) <> "All" Then
3340          Where = Where & "and d.status = '" & StatusCodes(CmbStatus.Text) & "' "
3350      End If

3360      If Trim(CmbAppCodePassed.Text) <> "None" Then
3370          Where = Where & "and exists (select 1 from lims_sys.sdg_log l " & _
                  "where l.sdg_id = d.sdg_id and " & _
                      "application_code = '" & AppCodes(CmbAppCodePassed.Text) & "' "
3380          If Trim(edtFromDate.Text) <> "" Then _
                  Where = Where & _
                      "and trunc(l.time,'ddd') >= to_date('" & edtFromDate.Text & "','dd/mm/yyyy') "
3390          If Trim(edtToDate.Text) <> "" Then _
                  Where = Where & _
                      "and trunc(l.time,'ddd') < to_date('" & edtToDate.Text & "','dd/mm/yyyy') + 1 "
3400          Where = Where & ") "
3410      End If

3420      If Trim(CmbAppCodeNot.Text) <> "None" Then
3430          Where = Where & "and not exists (select 1 from lims_sys.sdg_log l " & _
                  "where l.sdg_id = d.sdg_id and " & _
                      "application_code = '" & AppCodes(CmbAppCodeNot.Text) & "') "
3440      End If

3450      Where = Mid(Where, 5)
3460      sqlstr = "select name, status " & _
                   "from lims_sys.sdg d " & _
                   "order by d.name"
3470      If Trim(Where) <> "" Then
3480          sqlstr = "select name, status " & _
                       "from lims_sys.sdg d " & _
                       "where " & Where & " order by d.name"
3490      End If
3500      MousePointer = vbHourglass
3510      Set rst = con.Execute(sqlstr)
3520      lvSdg.ListItems.Clear
3530      While Not rst.EOF
3540          Set li = lvSdg.ListItems.Add(, , rst("NAME"), , imlsSdg.ListImages(rst("STATUS").Value).Index)
3550          rst.MoveNext
3560      Wend
3570      MousePointer = vbDefault
3580      If rst.RecordCount = 0 Then
3590          MsgBox "לא נמצאו דרישות !", vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, "Nautilus - Request Log"
3600      ElseIf rst.RecordCount = 1 Then
3610          MsgBox "נמצאה דרישה אחת.", vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, "Nautilus - Request Log"
3620      Else
3630          MsgBox "נמצאו " & rst.RecordCount & " דרישות.", vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, "Nautilus - Request Log"
3640      End If
3650      Exit Sub

ErrEnd:
3660      MousePointer = vbDefault
3670      MsgBox "btnSdgGo_Click... " & vbCrLf & _
                  "line # " & Erl & vbCrLf & Err.Description
End Sub



Private Sub lvSdg_DblClick()
3680      SSTab.Tab = 1
3690      edtSdgName.Text = lvSdg.SelectedItem.Text
3700      GetLogs4Sdg
End Sub

Private Sub Option1_Click()
3710      On Error GoTo ErrEnd
3720      CalendarFromDate.Visible = False
3730      CalendarToDate.Visible = False
3740      TxtFromDate.Text = ""
3750      TxtFromDate.Enabled = False
3760      CmdFromDate.Enabled = False
3770      TxtToDate.Text = ""
3780      TxtToDate.Enabled = False
3790      CmdToDate.Enabled = False
3800      CmbComputerName.ListIndex = 0
3810      CmbComputerName.Enabled = False
3820      edtSdgName.Text = ""
3830      edtSdgName.Enabled = True
3840      Call edtSdgName.SetFocus
3850      Exit Sub
ErrEnd:
3860      MsgBox "Option1_Click... " & vbCrLf & _
                  "line # " & Erl & vbCrLf & Err.Description
End Sub

Private Sub Option2_Click()
3870      On Error GoTo ErrEnd
3880      CalendarFromDate.Visible = False
3890      CalendarToDate.Visible = False
3900      TxtFromDate.Enabled = True
3910      CmdFromDate.Enabled = True
3920      TxtToDate.Enabled = True
3930      CmdToDate.Enabled = True
3940      CmbComputerName.ListIndex = 0
3950      TxtFromDate.Text = Format(Now, "dd/mm/yyyy")
3960      TxtToDate.Text = Format(Now, "dd/mm/yyyy")
3970      edtSdgName.Text = ""
3980      edtSdgName.Enabled = False
3990      CmbComputerName.Enabled = True
4000      Call CmbComputerName.SetFocus
4010      Exit Sub
ErrEnd:
4020      MsgBox "Option2_Click... " & vbCrLf & _
                  "line # " & Erl & vbCrLf & Err.Description
End Sub

Private Sub SSTab_Click(PreviousTab As Integer)
4030  On Error GoTo ERR_SSTab_Click

4040      If SSTab.Tab = 0 Then
4050          btnSdgGo.Default = True
4060          Call CmbAppCodePassed.SetFocus
4070      End If
4080      If SSTab.Tab = 1 Then
4090          btnLogGo.Default = True
4100          If edtSdgName.Enabled = True Then
4110              Call edtSdgName.SetFocus
4120          End If
4130      End If

4140      Exit Sub
ERR_SSTab_Click:
4150  MsgBox "ERR_SSTab_Click" & vbCrLf & "line # " & Erl & vbCrLf & Err.Description
End Sub

Private Function nte(e As Variant) As String
4160      nte = IIf(IsNull(e), "", e)
End Function

Private Sub TxtFromDate_KeyDown(KeyCode As Integer, Shift As Integer)
4170      If KeyCode = vbKeyDelete Then
4180          TxtFromDate.Text = ""
4190      End If
End Sub

Private Sub TxtToDate_KeyDown(KeyCode As Integer, Shift As Integer)
4200      If KeyCode = vbKeyDelete Then
4210          TxtToDate.Text = ""
4220      End If
End Sub

