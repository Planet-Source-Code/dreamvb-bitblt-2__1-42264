VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Bitblt 2"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   207
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   390
      Left            =   3360
      TabIndex        =   4
      Top             =   2550
      Width           =   1185
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imglst"
      DisabledImageList=   "imglst"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglst 
      Left            =   405
      Top             =   4005
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Buttons"
      Height          =   405
      Left            =   1875
      TabIndex        =   2
      Top             =   2550
      Width           =   1305
   End
   Begin VB.PictureBox picdst 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   60
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   3300
      Width           =   480
   End
   Begin VB.PictureBox picsrc 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   285
      Picture         =   "bitblt2.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   3840
      TabIndex        =   0
      Top             =   1335
      Width           =   3840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   1
      X2              =   62
      Y1              =   46
      Y2              =   46
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   1
      X2              =   62
      Y1              =   45
      Y2              =   45
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Dim Xcnt As Long ' Counter var

Private Sub Command1_Click()
Dim xHeight, xWidth As Long, FrameCnt As Long, iCnt As Long
    On Error Resume Next
    imglst.MaskColor = RGB(255, 0, 255) ' Mask colour for button
    
    xHeight = 32    ' Button Height
    xWidth = 32     ' Button Width
    
    FrameCnt = picdst.Width / 2 ' Find out how many images we have
    For iCnt = 1 To FrameCnt
        Xcnt = Xcnt + 32 ' Add 32 to our counter
        BitBlt picdst.hDC, 0, 0, xHeight, xWidth, picsrc.hDC, Xcnt - 32, 0, vbSrcCopy ' Copy image at position found
        imglst.ListImages.Add iCnt, "a" & iCnt, picdst.Image ' Add the new image to to listimage with key and index value
        Toolbar1.Buttons(iCnt).Image = imglst.ListImages(iCnt).Index ' Add the new images to the tool bar
        picdst.Refresh ' make sure our image shows
    Next
    
    ' Reset vars
    iCnt = 0
    Xcnt = 0
    xHeight = 0
    xWidth = 0
    FrameCnt = 0
    
End Sub

Private Sub Command2_Click()
    Unload Form1 ' unload the form
End Sub

Private Sub Form_Resize()
    ' Nothing special just add a 3D line under the toolbar
    Line1(0).X2 = Form1.Width
    Line1(1).X2 = Form1.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MsgBox "Hope you found this code usfull" & vbNewLine & "By Ben Jones", vbInformation, Form1.Caption ' think you know what this does
    Set Form1 = Nothing ' Release any memory
End Sub
