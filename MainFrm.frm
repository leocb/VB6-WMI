VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WMI - CPUID"
   ClientHeight    =   5475
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   7560
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox ContainerInit 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      FillColor       =   &H8000000B&
      ForeColor       =   &H8000000B&
      Height          =   5055
      Left            =   0
      ScaleHeight     =   5055
      ScaleWidth      =   7575
      TabIndex        =   46
      Top             =   6000
      Width           =   7575
      Begin ComctlLib.ProgressBar Loading 
         Height          =   495
         Left            =   720
         TabIndex        =   131
         Top             =   2280
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   873
         _Version        =   327682
         Appearance      =   0
      End
      Begin VB.Label LoadingInfo 
         Caption         =   "Iniciando"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   720
         TabIndex        =   48
         Top             =   2760
         Width           =   6165
      End
      Begin VB.Label TituloLoading 
         Alignment       =   2  'Center
         Caption         =   "Carregando..."
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
         Left            =   705
         TabIndex        =   47
         Top             =   1920
         Width           =   6135
      End
   End
   Begin VB.PictureBox container 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      FillColor       =   &H8000000B&
      ForeColor       =   &H8000000B&
      Height          =   4815
      Index           =   0
      Left            =   0
      ScaleHeight     =   4815
      ScaleWidth      =   7575
      TabIndex        =   0
      Top             =   330
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
         Index           =   11
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   2480
         Width           =   1215
      End
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
         Index           =   15
         Left            =   5800
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   2480
         Width           =   1535
      End
      Begin VB.CheckBox ProcCheck 
         Caption         =   "RDF/RDG/WRF/WRG"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   17
         Left            =   5040
         TabIndex        =   44
         Top             =   4200
         Width           =   1935
      End
      Begin VB.CheckBox ProcCheck 
         Caption         =   "EXCHANGE 128"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   16
         Left            =   5040
         TabIndex        =   43
         Top             =   3960
         Width           =   1935
      End
      Begin VB.CheckBox ProcCheck 
         Caption         =   "EXCHANGE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   5040
         TabIndex        =   42
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CheckBox ProcCheck 
         Caption         =   "COMPARE 128"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   5040
         TabIndex        =   41
         Top             =   3480
         Width           =   1815
      End
      Begin VB.CheckBox ProcCheck 
         Caption         =   "COMPARE 64"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   5040
         TabIndex        =   40
         Top             =   3240
         Width           =   1335
      End
      Begin VB.CheckBox ProcCheck 
         Caption         =   "COMPARE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   5040
         TabIndex        =   39
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CheckBox ProcCheck 
         Caption         =   "AVX"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   2880
         TabIndex        =   38
         Top             =   4200
         Width           =   1935
      End
      Begin VB.CheckBox ProcCheck 
         Caption         =   "VIRTUALIZATION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   2880
         TabIndex        =   37
         Top             =   3960
         Width           =   1575
      End
      Begin VB.CheckBox ProcCheck 
         Caption         =   "3D-NOW"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   2880
         TabIndex        =   36
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CheckBox ProcCheck 
         Caption         =   "CHANNELS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   2880
         TabIndex        =   35
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CheckBox ProcCheck 
         Caption         =   "SLAT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   2880
         TabIndex        =   34
         Top             =   3240
         Width           =   1335
      End
      Begin VB.CheckBox ProcCheck 
         Caption         =   "RDTSC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   2880
         TabIndex        =   33
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CheckBox ProcCheck 
         Caption         =   "PAE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   1920
         TabIndex        =   32
         Top             =   4200
         Width           =   1335
      End
      Begin VB.CheckBox ProcCheck 
         Caption         =   "DEP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   1920
         TabIndex        =   31
         Top             =   3960
         Width           =   1335
      End
      Begin VB.CheckBox ProcCheck 
         Caption         =   "SSE 3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   1920
         TabIndex        =   30
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CheckBox ProcCheck 
         Caption         =   "SSE 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   1920
         TabIndex        =   29
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CheckBox ProcCheck 
         Caption         =   "SSE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1920
         TabIndex        =   28
         Top             =   3240
         Width           =   1335
      End
      Begin VB.CheckBox ProcCheck 
         Caption         =   "MMX"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1920
         TabIndex        =   27
         Top             =   3000
         Width           =   1335
      End
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
         Index           =   14
         Left            =   5800
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   2115
         Width           =   1535
      End
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
         Index           =   13
         Left            =   5800
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1755
         Width           =   1535
      End
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
         Index           =   12
         Left            =   5800
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1395
         Width           =   1535
      End
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
         Index           =   10
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   2115
         Width           =   1215
      End
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
         Index           =   9
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1755
         Width           =   1215
      End
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
         Index           =   8
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1395
         Width           =   1215
      End
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
         Index           =   7
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   2115
         Width           =   1335
      End
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
         Index           =   6
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1755
         Width           =   1335
      End
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
         Index           =   5
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1395
         Width           =   1335
      End
      Begin VB.TextBox ProcInfo 
         Alignment       =   2  'Center
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
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   915
         Width           =   735
      End
      Begin VB.TextBox ProcInfo 
         Alignment       =   2  'Center
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
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   915
         Width           =   495
      End
      Begin VB.TextBox ProcInfo 
         Alignment       =   2  'Center
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   915
         Width           =   735
      End
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
         Index           =   1
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   555
         Width           =   5415
      End
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   195
         Width           =   5415
      End
      Begin VB.Label ProcLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Cache L3"
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
         Index           =   12
         Left            =   2760
         TabIndex        =   54
         Top             =   2520
         Width           =   1740
      End
      Begin VB.Label ProcLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Instruções Suportadas"
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
         Index           =   11
         Left            =   120
         TabIndex        =   26
         Top             =   3000
         Width           =   1740
      End
      Begin VB.Label ProcLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Cache L2"
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
         Index           =   10
         Left            =   2760
         TabIndex        =   19
         Top             =   2160
         Width           =   1740
      End
      Begin VB.Label ProcLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Cache L1 Inst."
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
         Index           =   9
         Left            =   2760
         TabIndex        =   18
         Top             =   1800
         Width           =   1740
      End
      Begin VB.Label ProcLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Cache L1 Data"
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
         Index           =   8
         Left            =   2760
         TabIndex        =   17
         Top             =   1440
         Width           =   1740
      End
      Begin VB.Label ProcLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Bus Clock"
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
         Index           =   7
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   1740
      End
      Begin VB.Label ProcLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Clock Atual"
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
         Index           =   6
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   1740
      End
      Begin VB.Label ProcLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Clock Máximo"
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
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   1740
      End
      Begin VB.Label ProcLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Arquitetura"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   5640
         TabIndex        =   9
         Top             =   960
         Width           =   840
      End
      Begin VB.Label ProcLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Núcleos Físicos/Lógicos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   2640
         TabIndex        =   7
         Top             =   960
         Width           =   1860
      End
      Begin VB.Label ProcLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Largura de Banda"
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
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1740
      End
      Begin VB.Label ProcLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Especificação"
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
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1740
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000A&
         Index           =   0
         X1              =   120
         X2              =   7440
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label ProcLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Descrição"
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
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1740
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000A&
         X1              =   120
         X2              =   7440
         Y1              =   2880
         Y2              =   2880
      End
   End
   Begin VB.PictureBox container 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      FillColor       =   &H8000000B&
      ForeColor       =   &H8000000B&
      Height          =   4815
      Index           =   1
      Left            =   0
      ScaleHeight     =   4815
      ScaleWidth      =   7575
      TabIndex        =   50
      Top             =   330
      Visible         =   0   'False
      Width           =   7575
      Begin VB.TextBox memInfo 
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
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   100
         Top             =   3075
         Width           =   1335
      End
      Begin VB.TextBox memInfo 
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
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   99
         Top             =   3800
         Width           =   1335
      End
      Begin VB.TextBox memInfo 
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
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   98
         Top             =   3440
         Width           =   1335
      End
      Begin VB.TextBox memInfo 
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
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   97
         Top             =   4155
         Width           =   1335
      End
      Begin VB.TextBox memInfo 
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
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   96
         Top             =   3075
         Width           =   1335
      End
      Begin VB.TextBox memInfo 
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
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   95
         Top             =   3795
         Width           =   1335
      End
      Begin VB.TextBox memInfo 
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
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   94
         Top             =   3435
         Width           =   1335
      End
      Begin VB.TextBox memInfo 
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
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   93
         Top             =   4155
         Width           =   1335
      End
      Begin VB.TextBox memInfo 
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   92
         Top             =   3075
         Width           =   1335
      End
      Begin VB.TextBox memInfo 
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   91
         Top             =   3795
         Width           =   1335
      End
      Begin VB.TextBox memInfo 
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   90
         Top             =   3435
         Width           =   1335
      End
      Begin VB.TextBox memInfo 
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   89
         Top             =   4155
         Width           =   1335
      End
      Begin VB.TextBox memInfo 
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   84
         Top             =   3080
         Width           =   1335
      End
      Begin VB.TextBox memInfo 
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   83
         Top             =   3795
         Width           =   1335
      End
      Begin VB.TextBox memInfo 
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   82
         Top             =   3435
         Width           =   1335
      End
      Begin VB.TextBox memInfo 
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   81
         Top             =   4155
         Width           =   1335
      End
      Begin VB.TextBox memInfoGeral 
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
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   80
         Top             =   555
         Width           =   1335
      End
      Begin VB.TextBox memInfoGeral 
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   79
         Top             =   555
         Width           =   1335
      End
      Begin VB.TextBox memInfoGeral 
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   78
         Top             =   195
         Width           =   1335
      End
      Begin VB.TextBox memInfo 
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
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   74
         Top             =   1280
         Width           =   1335
      End
      Begin VB.TextBox memInfo 
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
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   73
         Top             =   2000
         Width           =   1335
      End
      Begin VB.TextBox memInfo 
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
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   72
         Top             =   1640
         Width           =   1335
      End
      Begin VB.TextBox memInfo 
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
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   71
         Top             =   2360
         Width           =   1335
      End
      Begin VB.TextBox memInfo 
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
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   70
         Top             =   1280
         Width           =   1335
      End
      Begin VB.TextBox memInfo 
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
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   69
         Top             =   2000
         Width           =   1335
      End
      Begin VB.TextBox memInfo 
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
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   1640
         Width           =   1335
      End
      Begin VB.TextBox memInfo 
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
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   2360
         Width           =   1335
      End
      Begin VB.TextBox memInfo 
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   1280
         Width           =   1335
      End
      Begin VB.TextBox memInfo 
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   2000
         Width           =   1335
      End
      Begin VB.TextBox memInfo 
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   1640
         Width           =   1335
      End
      Begin VB.TextBox memInfo 
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   2360
         Width           =   1335
      End
      Begin VB.TextBox memInfo 
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   2360
         Width           =   1335
      End
      Begin VB.TextBox memInfo 
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   1640
         Width           =   1335
      End
      Begin VB.TextBox memInfo 
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   2000
         Width           =   1335
      End
      Begin VB.TextBox memInfo 
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   1280
         Width           =   1335
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000A&
         Index           =   9
         X1              =   5640
         X2              =   5640
         Y1              =   3000
         Y2              =   4560
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000A&
         Index           =   8
         X1              =   2520
         X2              =   2520
         Y1              =   3000
         Y2              =   4560
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000A&
         Index           =   7
         X1              =   4080
         X2              =   4080
         Y1              =   3000
         Y2              =   4560
      End
      Begin VB.Label memLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Slot"
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
         Index           =   10
         Left            =   120
         TabIndex        =   88
         Top             =   3120
         Width           =   900
      End
      Begin VB.Label memLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Capacidade"
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
         Index           =   9
         Left            =   0
         TabIndex        =   87
         Top             =   3840
         Width           =   1020
      End
      Begin VB.Label memLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo"
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
         Index           =   8
         Left            =   480
         TabIndex        =   86
         Top             =   3480
         Width           =   540
      End
      Begin VB.Label memLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Velocidade"
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
         Index           =   7
         Left            =   120
         TabIndex        =   85
         Top             =   4200
         Width           =   900
      End
      Begin VB.Label memLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Utilizado"
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
         Index           =   6
         Left            =   120
         TabIndex        =   77
         Top             =   600
         Width           =   900
      End
      Begin VB.Label memLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Disponivel"
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
         Index           =   5
         Left            =   2640
         TabIndex        =   76
         Top             =   600
         Width           =   780
      End
      Begin VB.Label memLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Total"
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
         Index           =   4
         Left            =   120
         TabIndex        =   75
         Top             =   240
         Width           =   900
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000A&
         Index           =   6
         X1              =   5640
         X2              =   5640
         Y1              =   1200
         Y2              =   2760
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000A&
         Index           =   5
         X1              =   2520
         X2              =   2520
         Y1              =   1200
         Y2              =   2760
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000A&
         Index           =   3
         X1              =   4080
         X2              =   4080
         Y1              =   1200
         Y2              =   2760
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000A&
         Index           =   2
         X1              =   120
         X2              =   7440
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label memLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Velocidade"
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
         Index           =   3
         Left            =   120
         TabIndex        =   61
         Top             =   2400
         Width           =   900
      End
      Begin VB.Label memLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo"
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
         Index           =   2
         Left            =   480
         TabIndex        =   59
         Top             =   1680
         Width           =   540
      End
      Begin VB.Label memLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Capacidade"
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
         Index           =   1
         Left            =   0
         TabIndex        =   57
         Top             =   2040
         Width           =   1020
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000A&
         Index           =   1
         X1              =   120
         X2              =   7440
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label memLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Slot"
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
         Index           =   0
         Left            =   120
         TabIndex        =   55
         Top             =   1320
         Width           =   900
      End
   End
   Begin VB.PictureBox container 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      FillColor       =   &H8000000B&
      ForeColor       =   &H8000000B&
      Height          =   4815
      Index           =   2
      Left            =   0
      ScaleHeight     =   4815
      ScaleWidth      =   7575
      TabIndex        =   51
      Top             =   330
      Visible         =   0   'False
      Width           =   7575
      Begin VB.TextBox BIOSInfo 
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
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   130
         Top             =   1040
         Width           =   1815
      End
      Begin VB.TextBox SOInfo 
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
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   129
         Top             =   3440
         Width           =   1815
      End
      Begin VB.TextBox SOInfo 
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
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   128
         Top             =   3440
         Width           =   735
      End
      Begin VB.TextBox SOInfo 
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   127
         Top             =   3440
         Width           =   1215
      End
      Begin VB.TextBox SOInfo 
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   126
         Top             =   3080
         Width           =   6255
      End
      Begin VB.TextBox VidInfo 
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
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   125
         Top             =   2600
         Width           =   1815
      End
      Begin VB.TextBox VidInfo 
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
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   124
         Top             =   2600
         Width           =   1575
      End
      Begin VB.TextBox VidInfo 
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   123
         Top             =   2600
         Width           =   1575
      End
      Begin VB.TextBox VidInfo 
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   122
         Top             =   2240
         Width           =   3735
      End
      Begin VB.TextBox VidInfo 
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   121
         Top             =   1880
         Width           =   3735
      End
      Begin VB.TextBox BIOSInfo 
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   120
         Top             =   1400
         Width           =   3735
      End
      Begin VB.TextBox BIOSInfo 
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   119
         Top             =   1040
         Width           =   3735
      End
      Begin VB.TextBox MBInfo 
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   118
         Top             =   560
         Width           =   3735
      End
      Begin VB.TextBox MBInfo 
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
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   117
         Top             =   200
         Width           =   1815
      End
      Begin VB.TextBox MBInfo 
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   116
         Top             =   200
         Width           =   3735
      End
      Begin VB.Label geralLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Data de Instalação"
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
         Index           =   14
         Left            =   3960
         TabIndex        =   115
         Top             =   3480
         Width           =   1515
      End
      Begin VB.Label geralLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Arquitetura"
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
         Index           =   13
         Left            =   2160
         TabIndex        =   114
         Top             =   3480
         Width           =   1035
      End
      Begin VB.Label geralLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Versão"
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
         Index           =   12
         Left            =   0
         TabIndex        =   113
         Top             =   3480
         Width           =   1035
      End
      Begin VB.Label geralLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Nome do SO"
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
         Index           =   11
         Left            =   0
         TabIndex        =   112
         Top             =   3120
         Width           =   1035
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000A&
         Index           =   11
         X1              =   120
         X2              =   7440
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Label geralLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Fabricante"
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
         Index           =   10
         Left            =   0
         TabIndex        =   111
         Top             =   2280
         Width           =   1020
      End
      Begin VB.Label geralLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "VRAM"
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
         Index           =   9
         Left            =   4440
         TabIndex        =   110
         Top             =   2640
         Width           =   1020
      End
      Begin VB.Label geralLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Data"
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
         Index           =   8
         Left            =   2160
         TabIndex        =   109
         Top             =   2640
         Width           =   1020
      End
      Begin VB.Label geralLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Driver"
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
         Index           =   7
         Left            =   0
         TabIndex        =   108
         Top             =   2640
         Width           =   1020
      End
      Begin VB.Label geralLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Placa Gráfica"
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
         Index           =   6
         Left            =   0
         TabIndex        =   107
         Top             =   1920
         Width           =   1020
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000A&
         Index           =   10
         X1              =   120
         X2              =   7440
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label geralLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Fabricante"
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
         Index           =   5
         Left            =   0
         TabIndex        =   106
         Top             =   1440
         Width           =   1020
      End
      Begin VB.Label geralLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Data"
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
         Index           =   4
         Left            =   4440
         TabIndex        =   105
         Top             =   1080
         Width           =   1020
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000A&
         Index           =   4
         X1              =   120
         X2              =   7440
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label geralLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "BIOS"
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
         Index           =   3
         Left            =   0
         TabIndex        =   104
         Top             =   1080
         Width           =   1020
      End
      Begin VB.Label geralLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Versão"
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
         Index           =   2
         Left            =   4440
         TabIndex        =   103
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label geralLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Fabricante"
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
         Index           =   1
         Left            =   0
         TabIndex        =   102
         Top             =   600
         Width           =   1020
      End
      Begin VB.Label geralLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Placa mãe"
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
         Index           =   0
         Left            =   0
         TabIndex        =   101
         Top             =   240
         Width           =   1020
      End
   End
   Begin VB.Timer StartupTimer 
      Interval        =   500
      Left            =   6600
      Top             =   -120
   End
   Begin VB.Timer RefreshTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7080
      Top             =   -120
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   5175
      Left            =   0
      TabIndex        =   49
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   9128
      ShowTips        =   0   'False
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Processador"
            Key             =   ""
            Object.Tag             =   "Proc"
            Object.ToolTipText     =   "Informações do Processador"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Memória"
            Key             =   ""
            Object.Tag             =   "Mem"
            Object.ToolTipText     =   "Informações da memória"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Geral"
            Key             =   ""
            Object.Tag             =   "Geral"
            Object.ToolTipText     =   "Informações Gerais"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Copyright 2015 - Leonardo Bottaro"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   2700
      TabIndex        =   45
      Top             =   5200
      Width           =   2145
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim step As Integer
Dim SelectedTab As Integer
Dim procInstruct(18) As Integer
Dim MemTotal As Integer

Private Declare Function IsProcessorFeaturePresent Lib "kernel32" (ByVal ProcessorFeature As Long) As Long

Private Sub Form_Load()
SelectedTab = 1
ContainerInit.Top = 0
ContainerInit.Left = 0
End Sub

Sub CarregaProcessador()

On Error Resume Next

Dim WMI
Dim wmiWin32Objects
Dim wmiWin32Object
Dim ComputerName As String

Set WMI = GetObject("WinMgmts://./root/cimv2")

Set wmiWin32Objects = WMI.ExecQuery("SELECT * FROM Win32_Processor", , 48)

For Each wmiWin32Object In wmiWin32Objects
With wmiWin32Object
ProcInfo(0).Text = .Name
ProcInfo(1).Text = .Description
ProcInfo(2).Text = .DataWidth & " bits"
ProcInfo(3).Text = .NumberOfCores & "/" & .NumberOfLogicalProcessors
ProcInfo(5).Text = .MaxClockSpeed & " MHz"
If .ExtClock > 0 Then ProcInfo(7).Text = .ExtClock & " MHz"


'Arquitetura
Select Case .Architecture
Case 0
ProcInfo(4).Text = "x86"
Case 1
ProcInfo(4).Text = "MIPS"
Case 2
ProcInfo(4).Text = "Alpha"
Case 3
ProcInfo(4).Text = "PowerPC"
Case 5
ProcInfo(4).Text = "ARM"
Case 6
ProcInfo(4).Text = "Itanium"
Case 9
ProcInfo(4).Text = "x64"
End Select

End With
'Instruções do processador
'Salva os mesmo valores na array, que serve para impedir a mudança no estado do botao
If IsProcessorFeaturePresent(3) = 0 Then procInstruct(0) = 0 Else procInstruct(0) = 1
If IsProcessorFeaturePresent(6) = 0 Then procInstruct(1) = 0 Else procInstruct(1) = 1
If IsProcessorFeaturePresent(10) = 0 Then procInstruct(2) = 0 Else procInstruct(2) = 1
If IsProcessorFeaturePresent(13) = 0 Then procInstruct(3) = 0 Else procInstruct(3) = 1
If IsProcessorFeaturePresent(12) = 0 Then procInstruct(4) = 0 Else procInstruct(4) = 1
If IsProcessorFeaturePresent(9) = 0 Then procInstruct(5) = 0 Else procInstruct(5) = 1
If IsProcessorFeaturePresent(8) = 0 Then procInstruct(6) = 0 Else procInstruct(6) = 1
If IsProcessorFeaturePresent(20) = 0 Then procInstruct(7) = 0 Else procInstruct(7) = 1
If IsProcessorFeaturePresent(16) = 0 Then procInstruct(8) = 0 Else procInstruct(8) = 1
If IsProcessorFeaturePresent(7) = 0 Then procInstruct(9) = 0 Else procInstruct(9) = 1
If IsProcessorFeaturePresent(21) = 0 Then procInstruct(10) = 0 Else procInstruct(10) = 1
If IsProcessorFeaturePresent(17) = 0 Then procInstruct(11) = 0 Else procInstruct(11) = 1
If IsProcessorFeaturePresent(2) = 0 Then procInstruct(12) = 0 Else procInstruct(12) = 1
If IsProcessorFeaturePresent(15) = 0 Then procInstruct(13) = 0 Else procInstruct(13) = 1
If IsProcessorFeaturePresent(14) = 0 Then procInstruct(14) = 0 Else procInstruct(14) = 1
If IsProcessorFeaturePresent(2) = 0 Then procInstruct(15) = 0 Else procInstruct(15) = 1
If IsProcessorFeaturePresent(14) = 0 And IsProcessorFeaturePresent(15) = 0 Then procInstruct(16) = 0 Else procInstruct(16) = 1
If IsProcessorFeaturePresent(22) = 0 Then procInstruct(17) = 0 Else procInstruct(17) = 1
'Checkboxes

If IsProcessorFeaturePresent(3) = 0 Then ProcCheck(0).Value = 0 Else ProcCheck(0).Value = 1
If IsProcessorFeaturePresent(6) = 0 Then ProcCheck(1).Value = 0 Else ProcCheck(1).Value = 1
If IsProcessorFeaturePresent(10) = 0 Then ProcCheck(2).Value = 0 Else ProcCheck(2).Value = 1
If IsProcessorFeaturePresent(13) = 0 Then ProcCheck(3).Value = 0 Else ProcCheck(3).Value = 1
If IsProcessorFeaturePresent(12) = 0 Then ProcCheck(4).Value = 0 Else ProcCheck(4).Value = 1
If IsProcessorFeaturePresent(9) = 0 Then ProcCheck(5).Value = 0 Else ProcCheck(5).Value = 1
If IsProcessorFeaturePresent(8) = 0 Then ProcCheck(6).Value = 0 Else ProcCheck(6).Value = 1
If IsProcessorFeaturePresent(20) = 0 Then ProcCheck(7).Value = 0 Else ProcCheck(7).Value = 1
If IsProcessorFeaturePresent(16) = 0 Then ProcCheck(8).Value = 0 Else ProcCheck(8).Value = 1
If IsProcessorFeaturePresent(7) = 0 Then ProcCheck(9).Value = 0 Else ProcCheck(9).Value = 1
If IsProcessorFeaturePresent(21) = 0 Then ProcCheck(10).Value = 0 Else ProcCheck(10).Value = 1
If IsProcessorFeaturePresent(17) = 0 Then ProcCheck(11).Value = 0 Else ProcCheck(11).Value = 1
If IsProcessorFeaturePresent(2) = 0 Then ProcCheck(12).Value = 0 Else ProcCheck(12).Value = 1
If IsProcessorFeaturePresent(15) = 0 Then ProcCheck(13).Value = 0 Else ProcCheck(13).Value = 1
If IsProcessorFeaturePresent(14) = 0 Then ProcCheck(14).Value = 0 Else ProcCheck(14).Value = 1
If IsProcessorFeaturePresent(2) = 0 Then ProcCheck(15).Value = 0 Else ProcCheck(15).Value = 1
If IsProcessorFeaturePresent(14) = 0 And IsProcessorFeaturePresent(15) = 0 Then ProcCheck(16).Value = 0 Else ProcCheck(16).Value = 1
If IsProcessorFeaturePresent(22) = 0 Then ProcCheck(17).Value = 0 Else ProcCheck(17).Value = 1

Next

End Sub


Sub CarregaCache()
On Error Resume Next

Dim WMI
Dim wmiWin32Objects
Dim wmiWin32Object
Dim ComputerName As String

Set WMI = GetObject("WinMgmts://./root/cimv2")

Set wmiWin32Objects = WMI.ExecQuery("SELECT * FROM Win32_CacheMemory", , 48)
For Each wmiWin32Object In wmiWin32Objects
With wmiWin32Object

Dim Associativity As String
Select Case .Associativity
    Case 1
    Associativity = "Outro"
    Case 2
    Associativity = "Desconhecido"
    Case 3
    Associativity = "Mapeamento Direto"
    Case 4
    Associativity = "2-way"
    Case 5
    Associativity = "4-way"
    Case 6
    Associativity = "Total"
    Case 7
    Associativity = "8-way"
    Case 8
    Associativity = "16-way"
End Select

Select Case .Purpose
    Case "L1 Cache"
    If .CacheType = 4 Then
        ProcInfo(8).Text = .MaxCacheSize & " KBytes"
        ProcInfo(12).Text = Associativity
    Else
        ProcInfo(9).Text = .MaxCacheSize & " KBytes"
        ProcInfo(13).Text = Associativity
    End If
    
    Case "L2 Cache"
        ProcInfo(10).Text = .MaxCacheSize & " KBytes"
        ProcInfo(14).Text = Associativity

    Case "L3 Cache"
        ProcInfo(11).Text = .MaxCacheSize & " KBytes"
        ProcInfo(15).Text = Associativity
End Select

End With
Next
End Sub

Private Sub ProcCheck_Click(Index As Integer)
    ProcCheck(Index).Value = procInstruct(Index)
End Sub


Sub CarregaMP()
On Error Resume Next

Dim WMI
Dim wmiWin32Objects
Dim wmiWin32Object
Dim ComputerName As String

'geral
Set WMI = GetObject("WinMgmts://./root/cimv2")
Set wmiWin32Objects = WMI.ExecQuery("SELECT TotalPhysicalMemory FROM Win32_ComputerSystem")

For Each wmiWin32Object In wmiWin32Objects
With wmiWin32Object

memInfoGeral(0).Text = Format(.TotalPhysicalMemory / 1024 / 1024, "######") & " Mbytes"
MemTotal = .TotalPhysicalMemory / 1024 / 1024

End With
Next

'SLOTS
Dim i As Integer
Dim tipo As String
i = 0
Set WMI = GetObject("WinMgmts://./root/cimv2")
Set wmiWin32Objects = WMI.ExecQuery("SELECT * FROM Win32_PhysicalMemory")

For Each wmiWin32Object In wmiWin32Objects
With wmiWin32Object

Select Case .MemoryType
Case 0
tipo = "Desconhecido"
Case 1
tipo = "Outro"
Case 2
tipo = "DRAM"
Case 3
tipo = "DRAM Síncrona"
Case 4
tipo = "Cache DRAM"
Case 5
tipo = "EDO"
Case 6
tipo = "EDRAM"
Case 7
tipo = "VRAM"
Case 8
tipo = "SRAM"
Case 9
tipo = "RAM"
Case 10
tipo = "ROM"
Case 11
tipo = "Flash"
Case 12
tipo = "EEPROM"
Case 13
tipo = "FEPROM"
Case 14
tipo = "EPROM"
Case 15
tipo = "CDRAM"
Case 16
tipo = "3DRAM"
Case 17
tipo = "SDRAM"
Case 18
tipo = "SGRAM"
Case 19
tipo = "RDRAM"
Case 20
tipo = "DDR"
Case 21
tipo = "DDR2"
Case 22
tipo = "DDR2 FB - DIMM"
Case 24
tipo = "DDR3"
Case 25
tipo = "FBD2"
End Select


memInfo(0 + i).Text = .DeviceLocator
memInfo(1 + i).Text = tipo
memInfo(2 + i).Text = .Capacity / 1048576 & " MBytes"
memInfo(3 + i).Text = .speed & " MHz"

i = i + 4
End With
Next

End Sub

Sub CarregaMB()
On Error Resume Next

Dim WMI
Dim wmiWin32Objects
Dim wmiWin32Object
Dim ComputerName As String

Set WMI = GetObject("WinMgmts://./root/cimv2")
Set wmiWin32Objects = WMI.ExecQuery("SELECT * FROM Win32_BaseBoard")

For Each wmiWin32Object In wmiWin32Objects
With wmiWin32Object

MBInfo(0).Text = .Product
MBInfo(1).Text = .Version
MBInfo(2).Text = .manufacturer

End With
Next

End Sub

Sub CarregaBIOS()
On Error Resume Next

Dim WMI
Dim wmiWin32Objects
Dim wmiWin32Object
Dim ComputerName As String

Set WMI = GetObject("WinMgmts://./root/cimv2")
Set wmiWin32Objects = WMI.ExecQuery("SELECT * FROM Win32_BIOS")

For Each wmiWin32Object In wmiWin32Objects
With wmiWin32Object

BIOSInfo(0).Text = .SMBIOSBIOSVersion
BIOSInfo(1).Text = .manufacturer
BIOSInfo(2).Text = Mid(.ReleaseDate, 7, 2) & "/" & Mid(.ReleaseDate, 5, 2) & "/" & Mid(.ReleaseDate, 1, 4)

End With
Next

End Sub

Sub CarregaVid()
On Error Resume Next

Dim WMI
Dim wmiWin32Objects
Dim wmiWin32Object
Dim ComputerName As String
Dim datai As String
Dim ram As Integer

Set WMI = GetObject("WinMgmts://./root/cimv2")
Set wmiWin32Objects = WMI.ExecQuery("SELECT * FROM Win32_VideoController")

For Each wmiWin32Object In wmiWin32Objects
With wmiWin32Object

VidInfo(0).Text = .Name
VidInfo(1).Text = .AdapterCompatibility '(manufacturer)
VidInfo(2).Text = .DriverVersion
datai = Mid(.DriverDate, 7, 2) & "/" & Mid(.DriverDate, 5, 2) & "/" & Mid(.DriverDate, 1, 4)
VidInfo(3).Text = datai
If .AdapterRAM / 1048576 < 0 Then _
ram = .AdapterRAM / 1048576 * -1 Else _
ram = .AdapterRAM / 1048576 * 1

VidInfo(4).Text = ram & " MBytes"

End With

GoTo fim
Next
fim:
End Sub

Sub CarregaSO()
On Error Resume Next

Dim WMI
Dim wmiWin32Objects
Dim wmiWin32Object
Dim ComputerName As String

Set WMI = GetObject("WinMgmts://./root/cimv2")
Set wmiWin32Objects = WMI.ExecQuery("SELECT * FROM Win32_OperatingSystem")

For Each wmiWin32Object In wmiWin32Objects
With wmiWin32Object

SOInfo(0).Text = .Caption
SOInfo(1).Text = .Version
SOInfo(2).Text = .OSArchitecture '32/64bit
SOInfo(3).Text = Mid(.InstallDate, 7, 2) & "/" & Mid(.InstallDate, 5, 2) & "/" & Mid(.InstallDate, 1, 4)

End With
Next

End Sub



Private Sub RefreshTimer_Timer()
On Error Resume Next

Dim WMI
Dim wmiWin32Objects
Dim wmiWin32Object
Dim ComputerName As String

'Processador
Set WMI = GetObject("WinMgmts://./root/cimv2")
Set wmiWin32Objects = WMI.ExecQuery("SELECT CurrentClockSpeed FROM Win32_Processor", , 48)

For Each wmiWin32Object In wmiWin32Objects
With wmiWin32Object

ProcInfo(6).Text = .CurrentClockSpeed & " MHz"

End With
Next

'MP
Set WMI = GetObject("WinMgmts://./root/cimv2")
Set wmiWin32Objects = WMI.ExecQuery("SELECT AvailableMBytes FROM Win32_PerfFormattedData_PerfOS_Memory", , 48)

For Each wmiWin32Object In wmiWin32Objects
With wmiWin32Object

memInfoGeral(2).Text = .AvailableMBytes & " MBytes"
memInfoGeral(1).Text = (MemTotal - .AvailableMBytes) & " MBytes"

End With
Next


End Sub

Private Sub StartupTimer_Timer()
    Dim pStep As Integer
    
    step = step + 1
    pStep = 8
    
    Select Case step
    Case 1
    LoadingInfo.Caption = "Carregando dados do processador"
    Loading.Value = step / pStep * 100
    
    Case 2
    CarregaProcessador
    LoadingInfo.Caption = "Carregando dados da memória cache"
    Loading.Value = step / pStep * 100
    
    Case 3
    CarregaCache
    LoadingInfo.Caption = "Carregando dados da memória principal"
    Loading.Value = step / pStep * 100
    
    Case 4
    CarregaMP
    LoadingInfo.Caption = "Carregando dados da Placa mãe"
    Loading.Value = step / pStep * 100
    
    Case 5
    CarregaMB
    LoadingInfo.Caption = "Carregando dados do BIOS"
    Loading.Value = step / pStep * 100
    
    Case 6
    CarregaBIOS
    LoadingInfo.Caption = "Carregando dados da Placa Gráfica"
    Loading.Value = step / pStep * 100
    
    Case 7
    CarregaVid
    LoadingInfo.Caption = "Carregando dados do Sistema Operacional"
    Loading.Value = step / pStep * 100
    
    Case 8
    CarregaSO
    LoadingInfo.Caption = "Inicializando..."
    Loading.Value = step / pStep * 100
    
    Case Else
    limpa
    RefreshTimer.Enabled = True
    StartupTimer.Enabled = False
    ContainerInit.Visible = False
    
    End Select
End Sub

Sub limpa()

'Aba processador
Dim Campos As Integer
Dim i As Integer

Campos = ProcInfo.Count
For i = 0 To Campos - 1 Step 1
    If (ProcInfo(i).Text = "" And (i <> 6 And i <> 16)) Then ProcInfo(i).BackColor = &H8000000F
Next i


Campos = SOInfo.Count
For i = 0 To Campos - 1 Step 1
    If SOInfo(i).Text = "" Then SOInfo(i).BackColor = &H8000000F
Next i

Campos = memInfo.Count
For i = 0 To Campos - 1 Step 1
    If memInfo(i).Text = "" Then memInfo(i).BackColor = &H8000000F
Next i

Campos = MBInfo.Count
For i = 0 To Campos - 1 Step 1
    If MBInfo(i).Text = "" Then MBInfo(i).BackColor = &H8000000F
Next i

Campos = BIOSInfo.Count
For i = 0 To Campos - 1 Step 1
    If BIOSInfo(i).Text = "" Then BIOSInfo(i).BackColor = &H8000000F
Next i

Campos = VidInfo.Count
For i = 0 To Campos - 1 Step 1
    If VidInfo(i).Text = "" Then VidInfo(i).BackColor = &H8000000F
Next i

Campos = SOInfo.Count
For i = 0 To Campos - 1 Step 1
    If SOInfo(i).Text = "" Then SOInfo(i).BackColor = &H8000000F
Next i
End Sub


Private Sub TabStrip1_Click()
    container(SelectedTab - 1).Visible = False
    SelectedTab = TabStrip1.SelectedItem.Index
    container(SelectedTab - 1).Visible = True
End Sub
