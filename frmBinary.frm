VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBinary 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Binary File Maker"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3600
   Icon            =   "frmBinary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   3600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdES 
      Caption         =   "Extract Selected File"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton cmdCombine 
      Caption         =   "Combine Files"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add to new Bin"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extract All Files"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   1320
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox lstfile 
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   3375
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open a Bin"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmBinary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************
'*
'*  Wrote by Bryan Skopnik (wishy@shaw.ca)
'*      Feel free to use this code as you wish.
'*
'****************************************************


'This structure will describe our binary file's
'size and number of contained files
Private Type FILEHEADER
    intNumFiles As Integer      'How many files are inside?
    lngFileSize As Long         'How big is this file? (Used to check integrity)
End Type

'This structure will describe each file contained
'in our binary file
Private Type INFOHEADER
    lngFileSize As Long         'How big is this chunk of stored data?
    lngFileStart As Long        'Where does the chunk start?
    strFileName As String * 16  'What's the name of the file this data came from?
End Type

Private Type MultFiles
    Path As String
    file As String
    data() As Byte
End Type

Dim I As Integer

Dim myFiles() As MultFiles
Dim filecount As Integer
Dim myPath As String
Dim myStr As String
Dim curExtract As Integer
Dim curBin As String

Private Sub cmdAdd_Click()

On Error GoTo errh
With cd
    .InitDir = App.Path
    .Flags = &H1000
    .CancelError = False
    .DialogTitle = "Add a file.."
    .Filter = "All Files (*.*) | *.*"
    .ShowOpen
End With

filecount = filecount + 1

ReDim Preserve myFiles(filecount)

Open cd.FileName For Binary Access Read Lock Write As #1
    ReDim myFiles(filecount).data(LOF(1) - 1)
    Get #1, 1, myFiles(filecount).data
Close #1

For I = Len(cd.FileName) To 1 Step -1
    If Mid(cd.FileName, I, 1) = "\" Then
        myFiles(filecount).Path = Left(cd.FileName, I)
        myFiles(filecount).file = Right(cd.FileName, Len(cd.FileName) - I)
        Exit For
    End If
Next

'Open myFiles(filecount).Path & "\test.txt" For Binary As #1
'    Put #1, , myFiles(filecount).data
'Close #1

lstfile.Clear
For I = 0 To UBound(myFiles())
    'MsgBox I & vbCrLf & vbCrLf & myFiles(I).Path & myFiles(I).File
    lstfile.AddItem myFiles(I).file
Next
    
Exit Sub
errh:
End Sub

Private Sub cmdCombine_Click()
    
On Error GoTo errh
    
    Dim myBinFile As Integer
    Dim filehead As FILEHEADER
    Dim infohead() As INFOHEADER
    Dim lngFileStart As Long
    
    filehead.intNumFiles = UBound(myFiles()) + 1
    'MsgBox Filehead.intNumFiles
        
    For I = 0 To UBound(myFiles())
        filehead.lngFileSize = filehead.lngFileSize + UBound(myFiles(I).data) + 1
        myStr = myStr & myFiles(I).file & vbCrLf
    Next
    filehead.lngFileSize = filehead.lngFileSize + (6) + (filehead.intNumFiles * 24)
    
    ReDim infohead(UBound(myFiles()))
    
    'MsgBox UBound(myFiles()) & vbCrLf & UBound(InfoHead())
    
    lngFileStart = (6) + (filehead.intNumFiles * 24) + 1
    
    For I = 0 To UBound(myFiles())
        infohead(I).lngFileSize = UBound(myFiles(0).data) + 1
        'MsgBox InfoHead(I).lngFileSize & vbCrLf & UBound(myFiles(I).data)
        infohead(I).lngFileStart = lngFileStart
        lngFileStart = lngFileStart + infohead(I).lngFileSize
        infohead(I).strFileName = myFiles(I).file
    Next
    
    With cd
        .InitDir = App.Path
        .CancelError = False
        .DialogTitle = "Save the bin..."
        .Flags = &H2
        .Filter = "Bin Files (*.bin) | *.bin"
        .ShowSave
    End With
    
    
    myBinFile = FreeFile
    Open cd.FileName For Binary Access Write Lock Write As myBinFile
        Put myBinFile, 1, filehead
        Put myBinFile, , infohead
        
        For I = 0 To UBound(myFiles())
            Put myBinFile, , myFiles(I).data
        Next
    Close myBinFile
    
    lstfile.Clear
    
    MsgBox "The files:" & vbCrLf & vbCrLf & myStr & vbCrLf & vbCrLf & "Were added into:" & vbCrLf & vbCrLf & cd.FileName
    
Exit Sub

errh:
    MsgBox "There was an error in compiling the file..", vbCritical + vbOKOnly, "Error..."
End Sub

Private Sub cmdES_Click()

On Error GoTo errh
    Dim myFile As Integer
    Dim myBin As Integer
    Dim myData() As Byte
    Dim filehead As FILEHEADER
    Dim infohead() As INFOHEADER
    
    myBin = FreeFile
    
    Open curBin For Binary Access Read Lock Write As myBin
    
        Get myBin, 1, filehead
        
        ReDim infohead(filehead.intNumFiles - 1)
        
        Get myBin, , infohead
        
        For I = 0 To UBound(infohead)
            If I = curExtract Then
                ReDim myData(infohead(I).lngFileSize - 1)
                Get myBin, infohead(I).lngFileStart, myData
                myStr = infohead(I).strFileName
                myFile = FreeFile
                Open myPath & "\" & myStr For Binary Access Write Lock Write As myFile
                    Put myFile, 1, myData
                Close myFile
                MsgBox "The file:" & vbCrLf & vbCrLf & myStr & vbCrLf & vbCrLf & "Was extracted to:" & vbCrLf & vbCrLf & myPath
                Exit For
            End If
        Next
    Close myBin

Exit Sub
errh:
End Sub

Private Sub cmdExtract_Click()
    
On Error GoTo errh
    
    Dim myFile As Integer
    Dim myBin As Integer
    Dim myData() As Byte
    Dim filehead As FILEHEADER
    Dim infohead() As INFOHEADER
    
    myBin = FreeFile
    
    With cd
        .InitDir = App.Path
        .DialogTitle = "Open a bin file..."
        .Flags = &H1000
        .CancelError = False
        .Filter = "Custom Bin (*.bin) | *.bin"
        .ShowOpen
    End With
    
    Open cd.FileName For Binary Access Read Lock Write As myBin
    
        For I = Len(cd.FileName) To 1 Step -1
            If Mid(cd.FileName, I, 1) = "\" Then
                myPath = Left(cd.FileName, I)
                'MsgBox myPath
                Exit For
            End If
        Next
        
    Get myBin, 1, filehead
    
    If LOF(myBin) <> filehead.lngFileSize Then
        MsgBox "This is not a valid file format.", vbOKOnly, "Invalid File"
        Exit Sub
    End If
    
    ReDim infohead(filehead.intNumFiles - 1)
    
    Get myBin, , infohead
    
    For I = 0 To UBound(infohead)
        ReDim myData(infohead(I).lngFileSize - 1)
        Get myBin, infohead(I).lngFileStart, myData
        myStr = myStr & infohead(I).strFileName & vbCrLf
        myFile = FreeFile
        Open myPath & infohead(I).strFileName For Binary Access Write Lock Write As myFile
            Put myFile, 1, myData
        Close myFile
    Next
    
    Close myBin
    
    MsgBox "The Following files:" & vbCrLf & vbCrLf & myStr & vbCrLf & vbCrLf & "Where extracted to:" & vbCrLf & vbCrLf & myPath

Exit Sub
errh:
    MsgBox "There was an error in extraction.", vbCritical + vbOKOnly, "Error..."
End Sub

Private Sub cmdOpen_Click()
On Error GoTo errh
    
    Dim myBin As Integer
    Dim filehead As FILEHEADER
    Dim infohead() As INFOHEADER
    
    myBin = FreeFile
    
    With cd
        .InitDir = App.Path
        .DialogTitle = "Open a bin file..."
        .Flags = &H1000
        .CancelError = False
        .Filter = "Custom Bin (*.bin) | *.bin"
        .ShowOpen
    End With
    
    Open cd.FileName For Binary Access Read Lock Write As myBin
    curBin = cd.FileName
    
        For I = Len(cd.FileName) To 1 Step -1
            If Mid(cd.FileName, I, 1) = "\" Then
                myPath = Left(cd.FileName, I)
                'MsgBox myPath
                Exit For
            End If
        Next
        
    Get myBin, 1, filehead
    
    If LOF(myBin) <> filehead.lngFileSize Then
        MsgBox "This is not a valid file format.", vbOKOnly, "Invalid File"
        Exit Sub
    End If
    
    ReDim infohead(filehead.intNumFiles - 1)
    ReDim mySFile(filehead.intNumFiles - 1)
        
    Get myBin, , infohead
    
    lstfile.Clear
    For I = 0 To UBound(infohead)
        lstfile.AddItem infohead(I).strFileName, I
    Next
    Close myBin
    
Exit Sub
errh:
    MsgBox "There was an error in extraction.", vbCritical + vbOKOnly, "Error..."
End Sub

Private Sub Form_Load()
    filecount = -1
End Sub

Private Sub lstfile_Click()
    cmdES.Enabled = True
    curExtract = lstfile.ListIndex
End Sub
