VERSION 5.00
Begin VB.Form frm_Main 
   BorderStyle     =   0  'None
   Caption         =   "GGJ_2020"
   ClientHeight    =   11250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15825
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   750
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cmd_Close 
      Caption         =   "Close"
      Height          =   480
      Left            =   13920
      TabIndex        =   1
      Top             =   360
      Width           =   810
   End
   Begin VB.Timer Tmr_Draw 
      Interval        =   100
      Left            =   11160
      Top             =   480
   End
   Begin VB.PictureBox PicMain 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   11175
      Left            =   120
      ScaleHeight     =   745
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1057
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   15855
   End
End
Attribute VB_Name = "frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmd_Close_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
     Dim i As Integer
    With this_Graphic.Buffer
        For i = 0 To 15
            SelectObject this_Graphic.Buffer.TileSetBmp(i), this_Graphic.Buffer.OldTilesetBmpDC(i)
            DeleteDC .TileSetBmp(i)
        Next i
        
        SelectObject .BackBuffer, .OldBackBufferDC
        DeleteDC .BackBuffer
        DeleteObject .BackBufferBmp
            
    End With

'    For i = 0 To Cmd_Object.UBound
'        Unload Cmd_Object(i)
'    Next i
    
    Unload Me
    End
End Sub

Private Sub Tmr_Draw_Timer()
    Call GameDraw(PicMain.hDC, this_Graphic, this_Switch)
End Sub
