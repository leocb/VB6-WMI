VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WMI - CPUID"
   ClientHeight    =   13665
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   21150
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   13665
   ScaleWidth      =   21150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdRefresh 
      Caption         =   "F5"
      Height          =   375
      Left            =   6120
      TabIndex        =   67
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton CmdSair 
      Caption         =   "Sair"
      Height          =   375
      Left            =   6720
      TabIndex        =   66
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton CmdCooler 
      Caption         =   "Coolers"
      Height          =   375
      Left            =   4800
      TabIndex        =   65
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CmdPlacamae 
      Caption         =   "Placa Mãe"
      Height          =   375
      Left            =   3240
      TabIndex        =   64
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton CmdMemoria 
      Caption         =   "Memória"
      Height          =   375
      Left            =   1680
      TabIndex        =   63
      Top             =   120
      Width           =   1455
   End
   Begin VB.Timer Refresh 
      Interval        =   1000
      Left            =   7800
      Top             =   120
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   57
      Left            =   15000
      Locked          =   -1  'True
      TabIndex        =   62
      Text            =   "Text1"
      Top             =   9720
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   56
      Left            =   15000
      Locked          =   -1  'True
      TabIndex        =   61
      Text            =   "Text1"
      Top             =   9375
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   55
      Left            =   15000
      Locked          =   -1  'True
      TabIndex        =   60
      Text            =   "Text1"
      Top             =   9045
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   54
      Left            =   15000
      Locked          =   -1  'True
      TabIndex        =   59
      Text            =   "Text1"
      Top             =   8715
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   53
      Left            =   15000
      Locked          =   -1  'True
      TabIndex        =   58
      Text            =   "Text1"
      Top             =   8370
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   52
      Left            =   15000
      Locked          =   -1  'True
      TabIndex        =   57
      Text            =   "Text1"
      Top             =   8040
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   51
      Left            =   15000
      Locked          =   -1  'True
      TabIndex        =   56
      Text            =   "Text1"
      Top             =   7710
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   50
      Left            =   15000
      Locked          =   -1  'True
      TabIndex        =   55
      Text            =   "Text1"
      Top             =   7380
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   49
      Left            =   15000
      Locked          =   -1  'True
      TabIndex        =   54
      Text            =   "Text1"
      Top             =   7035
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   48
      Left            =   15000
      Locked          =   -1  'True
      TabIndex        =   53
      Text            =   "Text1"
      Top             =   6705
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   47
      Left            =   15000
      Locked          =   -1  'True
      TabIndex        =   52
      Text            =   "Text1"
      Top             =   6375
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   46
      Left            =   15000
      Locked          =   -1  'True
      TabIndex        =   51
      Text            =   "Text1"
      Top             =   6045
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   45
      Left            =   15000
      Locked          =   -1  'True
      TabIndex        =   50
      Text            =   "Text1"
      Top             =   5700
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   44
      Left            =   15000
      Locked          =   -1  'True
      TabIndex        =   49
      Text            =   "Text1"
      Top             =   5370
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   43
      Left            =   15000
      Locked          =   -1  'True
      TabIndex        =   48
      Text            =   "Text1"
      Top             =   5040
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   42
      Left            =   15000
      Locked          =   -1  'True
      TabIndex        =   47
      Text            =   "Text1"
      Top             =   4695
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   41
      Left            =   15000
      Locked          =   -1  'True
      TabIndex        =   46
      Text            =   "Text1"
      Top             =   4365
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   40
      Left            =   15000
      Locked          =   -1  'True
      TabIndex        =   45
      Text            =   "Text1"
      Top             =   4035
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   39
      Left            =   15000
      Locked          =   -1  'True
      TabIndex        =   44
      Text            =   "Text1"
      Top             =   3705
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   38
      Left            =   15000
      Locked          =   -1  'True
      TabIndex        =   43
      Text            =   "Text1"
      Top             =   3360
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   37
      Left            =   15000
      Locked          =   -1  'True
      TabIndex        =   42
      Text            =   "Text1"
      Top             =   3030
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   36
      Left            =   15000
      Locked          =   -1  'True
      TabIndex        =   41
      Text            =   "Text1"
      Top             =   2700
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   35
      Left            =   15000
      Locked          =   -1  'True
      TabIndex        =   40
      Text            =   "Text1"
      Top             =   2370
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   34
      Left            =   15000
      Locked          =   -1  'True
      TabIndex        =   39
      Text            =   "Text1"
      Top             =   2025
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   33
      Left            =   15000
      Locked          =   -1  'True
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   1695
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   32
      Left            =   15000
      Locked          =   -1  'True
      TabIndex        =   37
      Text            =   "Text1"
      Top             =   1365
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   31
      Left            =   15000
      Locked          =   -1  'True
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   1035
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   30
      Left            =   15000
      Locked          =   -1  'True
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   690
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   29
      Left            =   15000
      Locked          =   -1  'True
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   360
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   28
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   9720
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   27
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   9375
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   26
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   9045
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   25
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   8715
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   24
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   8370
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   23
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   8040
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   22
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   7710
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   21
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   7380
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   20
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   7035
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   19
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   6705
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   18
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   6375
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   17
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   6045
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   16
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   5700
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   15
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   5370
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   14
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   5040
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   13
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   4695
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   12
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   4365
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   11
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   4035
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   10
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   3705
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   9
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   3360
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   8
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   3030
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   7
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   2700
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   6
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   2370
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   5
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   2025
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   4
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   1695
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   3
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1365
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   2
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1035
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   1
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   690
      Width           =   6000
   End
   Begin VB.TextBox InfoTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AA0000&
      Height          =   315
      Index           =   0
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   360
      Width           =   6000
   End
   Begin VB.CommandButton CmdProcessador 
      Caption         =   "Processador"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7215
      Left            =   120
      ScaleHeight     =   7185
      ScaleWidth      =   7545
      TabIndex        =   0
      Top             =   600
      Width           =   7575
      Begin VB.TextBox ProcInfo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00AA0000&
         Height          =   315
         Index           =   0
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   540
         Width           =   4335
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   7560
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Processador"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1185
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Nome"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1260
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdSair_Click()
End
End Sub

Private Sub Form_Load()

Dim Campos As Integer
Dim i As Integer
Campos = InfoTest.Count

For i = 0 To Campos - 1 Step 1
    InfoTest(i).Text = "----------------------------------------"
Next i

End Sub

Private Sub Refresh_Timer()

On Error Resume Next

Dim WMI
Dim wmiWin32Objects
Dim wmiWin32Object
Dim ComputerName As String

Set WMI = GetObject("WinMgmts://./root/cimv2")

Set wmiWin32Objects = WMI.ExecQuery("SELECT * FROM Win32_Processor", , 48)

With wmiWin32Object
For Each wmiWin32Object In wmiWin32Objects
InfoTest(0).Text = "(0 .AddressWidth)" & .AddressWidth
InfoTest(1).Text = "(1 .Architecture)" & .Architecture
InfoTest(2).Text = "(2 .AssetTag)" & .AssetTag
InfoTest(3).Text = "(3 .Availability)" & .Availability
InfoTest(4).Text = "(4 .Caption)" & .Caption
InfoTest(5).Text = "(5 .Characteristics)" & .Characteristics
InfoTest(6).Text = "(6 .ConfigManagerErrorCode)" & .ConfigManagerErrorCode
InfoTest(7).Text = "(7 .ConfigManagerUserConfig)" & .ConfigManagerUserConfig
InfoTest(8).Text = "(8 .CpuStatus)" & .CpuStatus
InfoTest(9).Text = "(9 .CreationClassName)" & .CreationClassName
InfoTest(10).Text = "(10 .CurrentClockSpeed) " & .CurrentClockSpeed
InfoTest(11).Text = "(11 .CurrentVoltage) " & .CurrentVoltage
InfoTest(12).Text = "(12 .DataWidth) " & .DataWidth
InfoTest(13).Text = "(13 .Description) " & .Description
InfoTest(14).Text = "(14 .DeviceID) " & .DeviceID
InfoTest(15).Text = "(15 .ErrorCleared) " & .ErrorCleared
InfoTest(16).Text = "(16 .ErrorDescription) " & .ErrorDescription
InfoTest(17).Text = "(17 .ExtClock) " & .ExtClock
InfoTest(18).Text = "(18 .Family) " & .Family
InfoTest(19).Text = "(19 .InstallDate) " & .InstallDate
InfoTest(20).Text = "(20 .L2CacheSize) " & .L2CacheSize
InfoTest(21).Text = "(21 .L2CacheSpeed) " & .L2CacheSpeed
InfoTest(22).Text = "(22 .L3CacheSize) " & .L3CacheSize
InfoTest(23).Text = "(23 .L3CacheSpeed) " & .L3CacheSpeed
InfoTest(24).Text = "(24 .LastErrorCode) " & .LastErrorCode
InfoTest(25).Text = "(25 .Level) " & .Level
InfoTest(26).Text = "(26 .LoadPercentage) " & .LoadPercentage
InfoTest(27).Text = "(27 .Manufacturer) " & .Manufacturer
InfoTest(28).Text = "(28 .MaxClockSpeed) " & .MaxClockSpeed
InfoTest(29).Text = "(29 .Name) " & .Name
InfoTest(30).Text = "(30 .NumberOfCores) " & .NumberOfCores
InfoTest(31).Text = "(31 .NumberOfEnabledCore) " & .NumberOfEnabledCore
InfoTest(32).Text = "(32 .NumberOfLogicalProcessors) " & .NumberOfLogicalProcessors
InfoTest(33).Text = "(33 .OtherFamilyDescription) " & .OtherFamilyDescription
InfoTest(34).Text = "(34 .PartNumber) " & .PartNumber
InfoTest(35).Text = "(35 .PNPDeviceID) " & .PNPDeviceID
InfoTest(36).Text = "(36 .PowerManagementCapabilities[]) " & .PowerManagementCapabilities(0)
InfoTest(37).Text = "(37 .PowerManagementSupported) " & .PowerManagementSupported
InfoTest(38).Text = "(38 .ProcessorId) " & .ProcessorId
InfoTest(39).Text = "(39 .ProcessorType) " & .ProcessorType
InfoTest(40).Text = "(40 .Revision) " & .Revision
InfoTest(41).Text = "(41 .Role) " & .Role
InfoTest(42).Text = "(42 .SecondLevelAddressTranslationExtensions) " & .SecondLevelAddressTranslationExtensions
InfoTest(43).Text = "(43 .SerialNumber) " & .SerialNumber
InfoTest(44).Text = "(44 .SocketDesignation) " & .SocketDesignation
InfoTest(45).Text = "(45 .Status) " & .Status
InfoTest(46).Text = "(46 .StatusInfo) " & .StatusInfo
InfoTest(47).Text = "(47 .Stepping) " & .Stepping
InfoTest(48).Text = "(48 .SystemCreationClassName) " & .SystemCreationClassName
InfoTest(49).Text = "(49 .SystemName) " & .SystemName
InfoTest(50).Text = "(50 .ThreadCount) " & .ThreadCount
InfoTest(51).Text = "(51 .UniqueId) " & .UniqueId
InfoTest(52).Text = "(52 .UpgradeMethod) " & .UpgradeMethod
InfoTest(53).Text = "(53 .Version) " & .Version
InfoTest(54).Text = "(54 .VirtualizationFirmwareEnabled) " & .VirtualizationFirmwareEnabled
InfoTest(55).Text = "(55 .VMMonitorModeExtensions) " & .VMMonitorModeExtensions
InfoTest(56).Text = "(56 .VoltageCaps) " & .VoltageCaps
Next
End With
End Sub
