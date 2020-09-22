VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmMain 
   Caption         =   "XLS2FLEXGRID"
   ClientHeight    =   6555
   ClientLeft      =   2985
   ClientTop       =   2265
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   8160
   Begin VB.OptionButton Option2 
      Caption         =   "Formulas"
      Height          =   315
      Left            =   1680
      TabIndex        =   8
      Top             =   960
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Values"
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load into Flexgrid"
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      Top             =   600
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   600
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   255
      Left            =   7200
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   5655
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   9128
      _Version        =   393216
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin VB.Label Label2 
      Caption         =   "Worksheet"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "XLS File"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Sub Command1_Click()
    Dim OFName As OPENFILENAME
    Dim XLS As Object
    Dim WRK As Object
    Dim SHT As Object
    
    OFName.lStructSize = Len(OFName)
    'Set the parent window
    OFName.hwndOwner = Me.hWnd
    'Set the application's instance
    OFName.hInstance = App.hInstance
    'Select a filter
    OFName.lpstrFilter = "Excel Files (*.xls)" + Chr$(0) + "*.xls" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    'create a buffer for the file
    OFName.lpstrFile = Space$(254)
    'set the maximum length of a returned file
    OFName.nMaxFile = 255
    'Create a buffer for the file title
    OFName.lpstrFileTitle = Space$(254)
    'Set the maximum length of a returned file title
    OFName.nMaxFileTitle = 255
    'Set the initial directory
    OFName.lpstrInitialDir = "C:\"
    'Set the title
    OFName.lpstrTitle = "Open XLS File"
    'No flags
    OFName.flags = 0

    'Show the 'Open File'-dialog
    If GetOpenFileName(OFName) Then
        Text1.Text = Trim$(OFName.lpstrFile)
        
        Combo1.Clear
        'Create a new instance of Excel
        Set XLS = CreateObject("Excel.Application")
        
        'Open the XLS file. The two parameters representes, UpdateLink = False and ReadOnly = True. These parameters have this setting to dont occur any error on broken links and allready opened XLS file.
        Set WRK = XLS.Workbooks.Open(Text1.Text, False, True)
        'Read all worksheets in xls file
        For Each SHT In WRK.Worksheets
            'Put the name of worksheet in combo
            Combo1.AddItem SHT.Name
        Next
        'Close the XLS file and dont save
        WRK.Close False
        'Quit the MS Excel
        XLS.Quit
        
        'Release variables
        Set XLS = Nothing
        Set WRK = Nothing
        Set SHT = Nothing
    Else
        MsgBox "Cancel was pressed"
    End If

End Sub


Private Sub Command2_Click()
On Error GoTo step_error
    Dim XLS As New Excel.Application
    Dim WRK As Excel.Workbook
    Dim SHT As Excel.Worksheet
    Dim RNG As Excel.Range
    
    Dim ArrayCells() As Variant
    
    If Combo1.ListIndex <> -1 Then
        'Create a new instance of Excel
        Set XLS = CreateObject("Excel.Application")
        'Open the XLS file. The two parameters representes, UpdateLink = False and ReadOnly = True. These parameters have this setting to dont occur any error on broken links and allready opened XLS file.
        Set WRK = XLS.Workbooks.Open(Text1.Text, False, True)
        'Set the SHT variable to selected worksheet
        Set SHT = WRK.Worksheets(Combo1.List(Combo1.ListIndex))
        
        'Get the used range of current worksheet
        Set RNG = SHT.UsedRange
        
        'Change the dimensions of array to fit the used range of worksheet
        ReDim ArrayCells(1 To RNG.Rows.Count, 1 To RNG.Columns.Count)
        
        'Transfer values of the used range to new array
        If Option1.Value Then
            ArrayCells = RNG.Value
        ElseIf Option2.Value Then
            ArrayCells = RNG.Formula
        End If
        
        'Close worksheet
        WRK.Close False
        'Quit the MS Excel
        XLS.Quit
        
        'Release variables
        Set XLS = Nothing
        Set WRK = Nothing
        Set SHT = Nothing
        Set RNG = Nothing
        
        'Configure the flexgrid to display data
        MSFlexGrid1.Redraw = False
        MSFlexGrid1.FixedCols = 0
        MSFlexGrid1.FixedRows = 0
        MSFlexGrid1.Rows = UBound(ArrayCells, 1)
        MSFlexGrid1.Cols = UBound(ArrayCells, 2)
        
        For r = 0 To UBound(ArrayCells, 1) - 1
            For c = 0 To UBound(ArrayCells, 2) - 1
                MSFlexGrid1.TextMatrix(r, c) = CStr(ArrayCells(r + 1, c + 1))
            Next
        Next
        MSFlexGrid1.Redraw = True
    Else
        MsgBox "Select the worksheet!", vbCritical
        Combo1.SetFocus
    End If
Exit Sub
step_error:
MsgBox Err.Number & " - " & Err.Description
Resume Next
End Sub


