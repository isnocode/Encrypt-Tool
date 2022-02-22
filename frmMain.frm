VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Encrypt/Decrypt Tool"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6360
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   6360
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPass 
      Height          =   315
      Left            =   1320
      TabIndex        =   4
      Top             =   960
      Width           =   4935
   End
   Begin VB.CommandButton cmdCom 
      Caption         =   "OK"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txtPath 
      Height          =   315
      Left            =   1200
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   480
      Width           =   5055
   End
   Begin VB.Label lblSta 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label lblAuthor 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © Free for All"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label lblPass 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblPath 
      BackStyle       =   0  'Transparent
      Caption         =   "File Path:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Drop your file into the ""File Path"" box to begin"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private EncryptCryptAPI As clsCryptAPI
Private Sub cmdCom_Click()
    'Complete : 0h15 26/11/2008 :)
    'On Error Resume Next
    If txtPass.Text = "" Then
        MsgBox "Password is required!"
    Else
    'Xu ly du lieu
        If FileExist(txtPath.Text) = True Then
        
            Dim buff As String
            Dim strMD5 As String
            
            buff = InputFile(txtPath)
            
            Dim strR As String
            strR = EncryptCryptAPI.DecryptString(Left(buff, 6), "dark")
            
            If strR <> "eNcOdE" Then
            'Neu du lieu chua duoc
                strMD5 = MD5String(buff)
                'buff = EncryptString(buff, txtPass.Text)
                'Cau truc
                'eNCoDe strMD5 Data
                
                buff = EncryptCryptAPI.EncryptString("eNcOdE", "dark") & _
                        EncryptCryptAPI.EncryptString(strMD5, "light" & CStr(Len(buff))) & _
                        EncryptCryptAPI.EncryptString(buff, txtPass.Text)
                
                CreatAFile buff, txtPath & ".encrypted"
                MsgBox "Successfully! File has been encrypted with the extension (.encrypted)."
            Else
                'Neu du lieu da duoc ma hoa
                
                Dim lgLen As Long
                lgLen = Len(buff) - 38
                
                strMD5 = Mid(buff, 7, 32)
                strMD5 = EncryptCryptAPI.DecryptString(strMD5, "light" & CStr(lgLen))
                
                buff = Right(buff, Len(buff) - 38)
                
                buff = EncryptCryptAPI.DecryptString(buff, txtPass.Text)
                
                If strMD5 = MD5String(buff) Then
                'Password dung
                    CreatAFile buff, txtPath & ".decrypted"
                    MsgBox "Successfully! File has been decrypted with the extension (.decrypted)." & vbCrLf & "To open the file you need to rename it by removing extensions (.encrypted & .decrypted)."
                Else
                'Password sai
                    MsgBox "Wrong Password!"
                End If
            End If
        Else
            MsgBox "Not Found!"
        End If
    End If
End Sub
Private Sub Form_Load()
    On Error Resume Next
    Dim strTmp As String
    strTmp = Command()
    DoEvents
    txtPath.Text = strTmp
    Set EncryptCryptAPI = New clsCryptAPI
End Sub
Private Sub txtPath_Change()
    On Error Resume Next
    Check
    DoEvents
End Sub
Private Sub txtPath_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'On Error Resume Next
    txtPath = Data.Files(1)
    Check
End Sub
Private Sub Check()
    If FileExist(txtPath.Text) = True Then
        cmdCom.Enabled = True
            Dim buff As String
            Dim strMD5 As String
            
            buff = InputFile(txtPath)
            
            Dim strR As String
            strR = EncryptCryptAPI.DecryptString(Left(buff, 6), "dark")
            
            If strR <> "eNcOdE" Then
                'Neu du lieu chua duoc
                cmdCom.Caption = "Encrypt"
                lblSta.Caption = "The file is not Encrypted"
            Else
                cmdCom.Caption = "Decrypt"
                lblSta.Caption = "The file is Encrypted"
            End If
    Else
        cmdCom.Enabled = False
        lblSta.Caption = "Not Found!"
    End If
End Sub
