VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAppRun 
   Caption         =   "Select an Application"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   3240
      Width           =   975
   End
   Begin VB.PictureBox picLarge 
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   5280
      ScaleHeight     =   510
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtMax 
      Height          =   285
      Left            =   5280
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   840
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   3960
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      PrinterDefault  =   0   'False
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   975
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3000
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   "Test"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView AppRun 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5106
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmAppRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const ATTR_NORMAL = 0
Const ATTR_READONLY = 1
Const ATTR_HIDDEN = 2
Const ATTR_SYSTEM = 4
Const ATTR_VOLUME = 8
Const ATTR_DIRECTORY = 16
Const ATTR_ARCHIVE = 32

'Public Variables
Dim sINI_File_Location As String


Private Sub AppRun_DblClick()
Dim itemx As ListItem

'Get the Selected the Item
    For Each itemx In AppRun.ListItems
        If itemx.Selected = True Then
                'Run the Selected item
                Call ShellExecute(hwnd, "Open", itemx.Key, "", App.Path, 1)
                Exit Sub
        End If
    Next itemx
End Sub

Private Sub cmdAdd_Click()
On Error GoTo nosave

'Load CommonDialog
cdl.CancelError = True
cdl.Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt + cdlOFNPathMustExist
cdl.DialogTitle = "Select a EXE program"
cdl.Filter = "EXE File|*.exe"
cdl.DefaultExt = "exe"
cdl.ShowOpen

sExeName = cdl.Filename

Save_INI_File sExeName
Load_Filename Next_Available_Number ' Call Sub
 
nosave:

End Sub

Private Sub cmdRemove_Click()
'Remove the Application from the view

On Error GoTo ErrFound

'Remove from INI File
    Remove_File AppRun.ListItems.Item(AppRun.SelectedItem.Index).Key
    
'Remove from View
    AppRun.ListItems.Remove AppRun.SelectedItem.Index
        
ErrFound:
    
End Sub


Private Sub Load_Filename(Optional KeyNumber As String)
'Load the Icon into the List View

ReDim glLargeIcons(lIcons)
ReDim glSmallIcons(lIcons)

On Error GoTo ErrFound


Dim lIndex

lIndex = "0"

'Get Icon from the File
Call ExtractIconEx(sExeName, lIndex, glLargeIcons(lIndex), glSmallIcons(lIndex), 1)

With picLarge
    Set .Picture = LoadPicture("")
     .AutoRedraw = True
    Call DrawIconEx(.hdc, 0, 0, glLargeIcons(lIndex), LARGE_ICON, LARGE_ICON, 0, 0, DI_NORMAL)
     .Refresh
End With

Mykey = sExeName & "(" & "-" & KeyNumber & ")"

ImageList1.ListImages.Add , Mykey, picLarge.Image
txtMax = sExeName

' Add Icon to Listview
AppRun.ListItems.Add , txtMax, SplitName(txtMax), Mykey


ErrFound:

End Sub
Function SplitName(Filename) As String
'Trim the File Path down to Size
' i.e c:\winnt\notepad.exe
' Stage 1 Removes any '\'
' Stage 2 Removes .exe

Dim tmpInstr As Integer
Dim tmpstr As String
Dim x, n As Integer

tmpInstr = InStr(Filename, "\")

' Stage 1 Remove any \
While tmpInstr > 0
            
    x = Len(Filename)
    n = x - tmpInstr
    tmpstr = Right$(Filename, n)
    Filename = tmpstr
    tmpInstr = InStr(Filename, "\")

Wend

' Stage 2 Remove .exe
tmpInstr = InStr(Filename, ".")

While tmpInstr > 0
    tmpstr = Left$(Filename, tmpInstr - 1)
    'tmpInstr = InStr(Filename, ".")
    tmpInstr = "0"
    Filename = tmpstr
    
Wend
' When complete, you end up with just the program name 'notepad'
    SplitName = Filename
End Function

Private Sub Form_Load()
'Load Startup Varables when the form loads

    sINI_File_Location = App.Path & "\Save.ini" ' Store the Location of the INI File
    
    Load_INI_File ' Load any Saved Data
    
End Sub

Private Sub Load_INI_File()
'Load Information from the INI file.
Dim xFinish As Boolean
Dim xFile As String
Dim xCount As String

'Reset Counters
xCount = "1": xFile = "": xFinish = False


While xFinish = False
    'Load data into xFile
    xFile = sGetINI(sINI_File_Location, "Save", xCount, "")
        If Len(xFile) > 0 Then  ' No More information to Load
            If UCase(xFile) <> UCase("Complete") Then
                sExeName = xFile
                Load_Filename xCount
            End If
        Else
            'If the Length of xFile is 0 the exit while loop
            xFinish = True
        End If
    
        xCount = xCount + 1
    
Wend
    
End Sub
Private Sub Save_INI_File(dFilename As String)
'Write the Data to the INI Files
    writeINI sINI_File_Location, "Save", Next_Available_Number, dFilename
End Sub
Private Sub Remove_File(dKey As String)
'Remove any Information from the INI File
    writeINI sINI_File_Location, "Save", get_Key_INI_Number(dKey), "Complete"
End Sub
Private Function get_Key_INI_Number(dKey As String) As Integer
'Find the Key avaible key
Dim xFinish As Boolean
Dim xFile As String
Dim xCount As String

'Reset Counters
xCount = "1": xFile = "": xFinish = False

While xFinish = False
    'Load data into xFile
    xFile = sGetINI(sINI_File_Location, "Save", xCount, "")
        If UCase(xFile) = UCase(dKey) Then ' No More information to Load
            get_Key_INI_Number = xCount
            xFinish = True
        Else
        End If
    
        xCount = xCount + 1
    
Wend
End Function
Private Function Next_Available_Number() As Integer
'Get the next Available Number in the INI File
Dim xFinish As Boolean
Dim xFile As String
Dim xCount As String

'Reset Counters
xCount = "1": xFile = "": xFinish = False

While xFinish = False
    'Load data into xFile
    xFile = sGetINI(sINI_File_Location, "Save", xCount, "")
        If Len(xFile) > 0 And UCase(xFile) <> UCase("Complete") Then  ' No More information to Load
        
        ElseIf UCase(xFile) = UCase("Complete") Or Len(xFile) = 0 Then
            'If the Length of xFile is 0 the exit while loop
            xFinish = True
            Next_Available_Number = xCount
        End If
    
        xCount = xCount + 1
    
Wend
End Function

