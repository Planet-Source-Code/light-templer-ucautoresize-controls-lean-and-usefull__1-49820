VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   " ucAutoResize Demo"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin AutoResizeDemo.ucAutoResize ucAutoResize 
      Left            =   210
      Top             =   3255
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.CommandButton btnReady 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   330
      Left            =   3930
      TabIndex        =   5
      Tag             =   "Tag value A - B - C  |LT"
      Top             =   3345
      Width           =   1125
   End
   Begin VB.Frame Frame 
      Height          =   3090
      Index           =   1
      Left            =   5025
      TabIndex        =   4
      Tag             =   "1,2 3,4 dif Tag values |LB"
      Top             =   -105
      Width           =   30
   End
   Begin VB.Frame Frame 
      Height          =   30
      Index           =   0
      Left            =   -60
      TabIndex        =   3
      Tag             =   "Tag values for other purposes|RT"
      Top             =   2955
      Width           =   5610
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   " A frame"
      Height          =   2205
      Left            =   300
      TabIndex        =   0
      Tag             =   "Some more Tag values |RB"
      Top             =   585
      Width           =   4395
      Begin VB.ListBox List 
         BackColor       =   &H00FFFFFF&
         Height          =   900
         IntegralHeight  =   0   'False
         ItemData        =   "frmMain.frx":0000
         Left            =   180
         List            =   "frmMain.frx":0007
         TabIndex        =   8
         Tag             =   "Val1, Val2 |-T"
         Top             =   1185
         Width           =   1485
      End
      Begin VB.CommandButton Command1 
         Caption         =   "A button"
         Height          =   315
         Left            =   3180
         TabIndex        =   2
         Tag             =   "Different Tag value |LT"
         Top             =   1740
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         Height          =   810
         Left            =   180
         MultiLine       =   -1  'True
         TabIndex        =   1
         Tag             =   "Some Tag values |RB"
         Text            =   "frmMain.frx":0017
         Top             =   300
         Width           =   4095
      End
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Two frames abused as 3D lines"
      Height          =   180
      Index           =   1
      Left            =   3180
      TabIndex        =   7
      Tag             =   "|LT"
      Top             =   2985
      Width           =   2340
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "<----- Thats all you need. No code."
      Height          =   225
      Index           =   2
      Left            =   795
      TabIndex        =   9
      Top             =   3465
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Simple example. Look at the TAG values, RUN this app and resize the form !"
      Height          =   435
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   75
      Width           =   4410
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
