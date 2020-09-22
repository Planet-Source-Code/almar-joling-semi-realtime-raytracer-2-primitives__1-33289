VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Raytrace II"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10140
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   10140
   StartUpPosition =   2  'CenterScreen
   Begin VB.VScrollBar scrFov 
      Height          =   2625
      Left            =   7005
      Max             =   1500
      TabIndex        =   7
      Top             =   2010
      Value           =   600
      Width           =   270
   End
   Begin VB.CheckBox chkRender 
      Caption         =   "Enable Cylinder"
      Height          =   225
      Index           =   3
      Left            =   7050
      TabIndex        =   6
      Top             =   1620
      Value           =   1  'Checked
      Width           =   3060
   End
   Begin VB.CheckBox chkRender 
      Caption         =   "Enable Plane"
      Height          =   225
      Index           =   2
      Left            =   7050
      TabIndex        =   5
      Top             =   1350
      Value           =   1  'Checked
      Width           =   3060
   End
   Begin VB.CheckBox chkRender 
      Caption         =   "Enable Vertices"
      Height          =   225
      Index           =   1
      Left            =   7050
      TabIndex        =   4
      Top             =   1080
      Value           =   1  'Checked
      Width           =   3060
   End
   Begin VB.CheckBox chkRender 
      Caption         =   "Enable Sphere"
      Height          =   225
      Index           =   0
      Left            =   7050
      TabIndex        =   3
      Top             =   810
      Value           =   1  'Checked
      Width           =   3060
   End
   Begin VB.CheckBox chkReflections 
      Caption         =   "Enable Reflections"
      Height          =   225
      Left            =   7050
      TabIndex        =   2
      Top             =   345
      Width           =   3060
   End
   Begin VB.CheckBox chkShadows 
      Caption         =   "Enable Shadows"
      Height          =   225
      Left            =   7050
      TabIndex        =   1
      Top             =   75
      Width           =   3060
   End
   Begin VB.PictureBox picRay 
      BorderStyle     =   0  'None
      Height          =   6360
      Left            =   0
      ScaleHeight     =   424
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   459
      TabIndex        =   0
      Top             =   0
      Width           =   6885
   End
   Begin VB.Label lblDistance 
      Caption         =   "FOV"
      Height          =   195
      Left            =   7020
      TabIndex        =   8
      Top             =   4740
      Width           =   540
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'//Realtime raytracer [version 2]
'//Original (c++) version and other nice
'//Raytrace versions (with shadows, cilinders, etc)
'//Can be found at http://www.2tothex.com/
'//VB port by Almar Joling / quadrantwars@quadrantwars.com
'//Websites: http://www.quadrantwars.com (my game)
'//          http://vbfibre.digitalrice.com (Many VB speed tricks with benchmarks)

'//This code is highly optimized. If you manage to gain some more FPS
'//I'm always interested =-)

'//Finished @ 01/04/2002
'//Feel free to post this code anywhere, but please leave the above info
'//and author info intact. Thank you.

Private Sub chkRender_Click(Index As Integer)
    '//Reset scene
    mdlScene.SetupScene
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    End
End Sub
