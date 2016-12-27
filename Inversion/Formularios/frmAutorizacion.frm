VERSION 5.00
Begin VB.Form frmAutorizacion 
   Caption         =   "Autorización"
   ClientHeight    =   2235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   3135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAccion 
      Caption         =   "&Cancelar"
      Height          =   735
      Index           =   1
      Left            =   1680
      Picture         =   "frmAutorizacion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "&Aceptar"
      Height          =   735
      Index           =   0
      Left            =   360
      Picture         =   "frmAutorizacion.frx":0562
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtUsuarioAutoriza 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1245
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   3
      ToolTipText     =   "Clave de autorización"
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtClave 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1245
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   0
      ToolTipText     =   "Clave de autorización"
      Top             =   795
      Width           =   1575
   End
   Begin VB.Label lblDescrip 
      Caption         =   "Usuario"
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   750
   End
   Begin VB.Label lblDescrip 
      Caption         =   "Clave"
      Height          =   285
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   795
      Width           =   750
   End
End
Attribute VB_Name = "frmAutorizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
