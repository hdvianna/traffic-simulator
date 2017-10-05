VERSION 5.00
Begin VB.Form frmSimA 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Simulação Semáforo de 2 Tempos"
   ClientHeight    =   4140
   ClientLeft      =   1560
   ClientTop       =   2010
   ClientWidth     =   4830
   ControlBox      =   0   'False
   Icon            =   "frmSimA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   276
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   322
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture12 
      Height          =   1140
      Left            =   4320
      ScaleHeight     =   1140
      ScaleWidth      =   30
      TabIndex        =   181
      Top             =   2295
      Width           =   30
   End
   Begin VB.PictureBox Picture11 
      Height          =   1140
      Left            =   4320
      ScaleHeight     =   1140
      ScaleWidth      =   30
      TabIndex        =   180
      Top             =   630
      Width           =   30
   End
   Begin VB.PictureBox Picture10 
      Height          =   1140
      Left            =   2400
      ScaleHeight     =   1140
      ScaleWidth      =   30
      TabIndex        =   179
      Top             =   2310
      Width           =   30
   End
   Begin VB.PictureBox Picture9 
      Height          =   1140
      Left            =   2385
      ScaleHeight     =   1140
      ScaleWidth      =   30
      TabIndex        =   178
      Top             =   630
      Width           =   30
   End
   Begin VB.PictureBox Picture8 
      Height          =   1140
      Left            =   465
      ScaleHeight     =   1140
      ScaleWidth      =   30
      TabIndex        =   177
      Top             =   2340
      Width           =   30
   End
   Begin VB.PictureBox Picture7 
      Height          =   1140
      Left            =   465
      ScaleHeight     =   1140
      ScaleWidth      =   30
      TabIndex        =   176
      Top             =   660
      Width           =   30
   End
   Begin VB.PictureBox Picture6 
      Height          =   15
      Left            =   2700
      ScaleHeight     =   15
      ScaleWidth      =   1440
      TabIndex        =   175
      Top             =   3735
      Width           =   1440
   End
   Begin VB.PictureBox Picture5 
      Height          =   15
      Left            =   735
      ScaleHeight     =   15
      ScaleWidth      =   1440
      TabIndex        =   174
      Top             =   3735
      Width           =   1440
   End
   Begin VB.PictureBox Picture4 
      Height          =   15
      Left            =   2685
      ScaleHeight     =   15
      ScaleWidth      =   1440
      TabIndex        =   173
      Top             =   2040
      Width           =   1440
   End
   Begin VB.PictureBox Picture3 
      Height          =   15
      Left            =   735
      ScaleHeight     =   15
      ScaleWidth      =   1440
      TabIndex        =   172
      Top             =   2040
      Width           =   1440
   End
   Begin VB.PictureBox Picture2 
      Height          =   15
      Left            =   2670
      ScaleHeight     =   15
      ScaleWidth      =   1440
      TabIndex        =   171
      Top             =   345
      Width           =   1440
   End
   Begin VB.PictureBox Picture1 
      Height          =   15
      Left            =   705
      ScaleHeight     =   15
      ScaleWidth      =   1440
      TabIndex        =   170
      Top             =   345
      Width           =   1440
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   225
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   0
      Top             =   105
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   225
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   1
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   225
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   2
      Top             =   600
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   225
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   3
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   225
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   4
      Top             =   1080
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   225
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   5
      Top             =   1320
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   225
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   6
      Top             =   1560
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   225
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   7
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   225
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   8
      Top             =   2055
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   225
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   9
      Top             =   2295
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   225
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   10
      Top             =   2535
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   225
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   11
      Top             =   2775
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   225
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   12
      Top             =   3015
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   225
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   13
      Top             =   3255
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   16
      Left            =   480
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   16
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   17
      Left            =   720
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   17
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   18
      Left            =   960
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   18
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   19
      Left            =   1200
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   19
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   20
      Left            =   1440
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   20
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   21
      Left            =   1680
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   21
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   22
      Left            =   1920
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   22
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   23
      Left            =   2160
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   23
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   24
      Left            =   2415
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   24
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   25
      Left            =   2655
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   25
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   26
      Left            =   2895
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   26
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   27
      Left            =   3135
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   27
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   28
      Left            =   3375
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   28
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   29
      Left            =   3615
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   29
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   30
      Left            =   3855
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   30
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   31
      Left            =   480
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   31
      Top             =   105
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   32
      Left            =   480
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   32
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   33
      Left            =   480
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   33
      Top             =   600
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   34
      Left            =   480
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   34
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   35
      Left            =   480
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   35
      Top             =   1080
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   36
      Left            =   480
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   36
      Top             =   1320
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   37
      Left            =   480
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   37
      Top             =   1560
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   38
      Left            =   480
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   38
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   39
      Left            =   480
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   39
      Top             =   2055
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   40
      Left            =   480
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   40
      Top             =   2295
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   41
      Left            =   480
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   41
      Top             =   2535
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   42
      Left            =   480
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   42
      Top             =   2775
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   43
      Left            =   480
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   43
      Top             =   3015
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   44
      Left            =   480
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   44
      Top             =   3255
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   47
      Left            =   720
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   47
      Top             =   2055
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   48
      Left            =   960
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   48
      Top             =   2055
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   49
      Left            =   1200
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   49
      Top             =   2055
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   50
      Left            =   1440
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   50
      Top             =   2055
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   51
      Left            =   1680
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   51
      Top             =   2055
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   52
      Left            =   1920
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   52
      Top             =   2055
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   53
      Left            =   2160
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   53
      Top             =   2055
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   54
      Left            =   2400
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   54
      Top             =   2055
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   55
      Left            =   2655
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   55
      Top             =   2055
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   56
      Left            =   2895
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   56
      Top             =   2055
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   57
      Left            =   3135
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   57
      Top             =   2055
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   58
      Left            =   3375
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   58
      Top             =   2055
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   59
      Left            =   3615
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   59
      Top             =   2055
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   60
      Left            =   3855
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   60
      Top             =   2055
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   61
      Left            =   720
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   61
      Top             =   105
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   62
      Left            =   960
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   62
      Top             =   105
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   63
      Left            =   1200
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   63
      Top             =   105
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   64
      Left            =   1440
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   64
      Top             =   105
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   65
      Left            =   1680
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   65
      Top             =   105
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   66
      Left            =   1920
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   66
      Top             =   105
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   67
      Left            =   2160
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   67
      Top             =   105
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   68
      Left            =   2415
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   68
      Top             =   105
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   69
      Left            =   2655
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   69
      Top             =   105
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   70
      Left            =   2895
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   70
      Top             =   105
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   71
      Left            =   3135
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   71
      Top             =   105
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   72
      Left            =   3375
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   72
      Top             =   105
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   73
      Left            =   3615
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   73
      Top             =   105
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   74
      Left            =   3855
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   74
      Top             =   105
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   75
      Left            =   720
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   75
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   76
      Left            =   960
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   76
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   77
      Left            =   1200
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   77
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   78
      Left            =   1440
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   78
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   79
      Left            =   1680
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   79
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   80
      Left            =   1920
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   80
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   81
      Left            =   2160
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   81
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   82
      Left            =   2415
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   82
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   83
      Left            =   2670
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   83
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   84
      Left            =   2895
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   84
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   85
      Left            =   3135
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   85
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   86
      Left            =   3375
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   86
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   87
      Left            =   3615
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   87
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   88
      Left            =   3855
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   88
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   117
      Left            =   2160
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   117
      Top             =   600
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   118
      Left            =   2160
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   118
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   119
      Left            =   2160
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   119
      Top             =   1080
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   120
      Left            =   2160
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   120
      Top             =   1320
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   121
      Left            =   2160
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   121
      Top             =   1560
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   122
      Left            =   2415
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   122
      Top             =   600
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   123
      Left            =   2415
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   123
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   124
      Left            =   2415
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   124
      Top             =   1065
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   125
      Left            =   2415
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   125
      Top             =   1320
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   126
      Left            =   2415
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   126
      Top             =   1560
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   127
      Left            =   2160
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   127
      Top             =   2295
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   128
      Left            =   2160
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   128
      Top             =   2535
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   129
      Left            =   2160
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   129
      Top             =   2775
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   130
      Left            =   2160
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   130
      Top             =   3015
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   131
      Left            =   2160
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   131
      Top             =   3255
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   132
      Left            =   2415
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   132
      Top             =   2295
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   133
      Left            =   2415
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   133
      Top             =   2535
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   134
      Left            =   2415
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   134
      Top             =   2775
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   135
      Left            =   2415
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   135
      Top             =   3015
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   136
      Left            =   2415
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   136
      Top             =   3255
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   137
      Left            =   4095
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   137
      Top             =   105
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   138
      Left            =   4095
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   138
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   139
      Left            =   4095
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   139
      Top             =   600
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   140
      Left            =   4095
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   140
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   141
      Left            =   4095
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   141
      Top             =   1080
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   143
      Left            =   4095
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   143
      Top             =   1560
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   144
      Left            =   4095
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   144
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   145
      Left            =   4095
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   145
      Top             =   2055
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   146
      Left            =   4095
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   146
      Top             =   2295
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   147
      Left            =   4095
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   147
      Top             =   2535
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   148
      Left            =   4095
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   148
      Top             =   2775
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   149
      Left            =   4095
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   149
      Top             =   3015
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   150
      Left            =   4095
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   150
      Top             =   3255
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   153
      Left            =   4350
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   153
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   154
      Left            =   4350
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   154
      Top             =   105
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   155
      Left            =   4350
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   155
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   156
      Left            =   4350
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   156
      Top             =   600
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   157
      Left            =   4350
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   157
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   158
      Left            =   4350
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   158
      Top             =   1080
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   159
      Left            =   4350
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   159
      Top             =   1320
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   160
      Left            =   4350
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   160
      Top             =   1560
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   161
      Left            =   4350
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   161
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   162
      Left            =   4350
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   162
      Top             =   2055
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   163
      Left            =   4350
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   163
      Top             =   2295
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   164
      Left            =   4350
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   164
      Top             =   2535
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   165
      Left            =   4350
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   165
      Top             =   2775
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   166
      Left            =   4350
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   166
      Top             =   3015
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   167
      Left            =   4350
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   167
      Top             =   3255
      Width           =   255
   End
   Begin VB.Timer Tick 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   915
      Top             =   690
   End
   Begin VB.Timer geraCarro 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   1485
      Top             =   705
   End
   Begin VB.Timer tsemA 
      Enabled         =   0   'False
      Left            =   930
      Top             =   1140
   End
   Begin VB.Timer tsemB 
      Enabled         =   0   'False
      Left            =   1515
      Top             =   1140
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   225
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   15
      Top             =   3750
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   45
      Left            =   480
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   45
      Top             =   3495
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   46
      Left            =   480
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   46
      Top             =   3750
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   89
      Left            =   720
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   89
      Top             =   3495
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   90
      Left            =   960
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   90
      Top             =   3495
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   92
      Left            =   1440
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   92
      Top             =   3495
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   93
      Left            =   1695
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   93
      Top             =   3495
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   94
      Left            =   1920
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   94
      Top             =   3495
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   95
      Left            =   2160
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   95
      Top             =   3495
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   96
      Left            =   2415
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   96
      Top             =   3495
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   97
      Left            =   2655
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   97
      Top             =   3495
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   98
      Left            =   2895
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   98
      Top             =   3495
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   99
      Left            =   3135
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   99
      Top             =   3495
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   100
      Left            =   3375
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   100
      Top             =   3495
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   101
      Left            =   3615
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   101
      Top             =   3495
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   102
      Left            =   3855
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   102
      Top             =   3495
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   103
      Left            =   720
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   103
      Top             =   3750
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   104
      Left            =   960
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   104
      Top             =   3750
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   105
      Left            =   1200
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   105
      Top             =   3750
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   106
      Left            =   1440
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   106
      Top             =   3750
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   107
      Left            =   1680
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   107
      Top             =   3750
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   108
      Left            =   1920
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   108
      Top             =   3750
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   109
      Left            =   2160
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   109
      Top             =   3750
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   110
      Left            =   2415
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   110
      Top             =   3750
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   111
      Left            =   2655
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   111
      Top             =   3750
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   112
      Left            =   2895
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   112
      Top             =   3750
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   113
      Left            =   3135
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   113
      Top             =   3750
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   114
      Left            =   3375
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   114
      Top             =   3750
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   115
      Left            =   3615
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   115
      Top             =   3750
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   116
      Left            =   3855
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   116
      Top             =   3750
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   151
      Left            =   4095
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   151
      Top             =   3495
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   152
      Left            =   4095
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   152
      Top             =   3750
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   168
      Left            =   4350
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   168
      Top             =   3495
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   169
      Left            =   4350
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   169
      Top             =   3750
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   225
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   14
      Top             =   3495
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   91
      Left            =   1200
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   91
      Top             =   3495
      Width           =   255
   End
   Begin VB.PictureBox pista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   142
      Left            =   4095
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   142
      Top             =   1320
      Width           =   255
   End
   Begin VB.Line Line12 
      BorderWidth     =   2
      X1              =   14
      X2              =   32
      Y1              =   23
      Y2              =   23
   End
   Begin VB.Line Line11 
      BorderWidth     =   2
      X1              =   288
      X2              =   306
      Y1              =   251
      Y2              =   251
   End
   Begin VB.Line Line10 
      BorderWidth     =   2
      X1              =   288
      X2              =   306
      Y1              =   138
      Y2              =   138
   End
   Begin VB.Line Line9 
      BorderWidth     =   2
      X1              =   288
      X2              =   306
      Y1              =   23
      Y2              =   23
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   14
      X2              =   32
      Y1              =   251
      Y2              =   251
   End
   Begin VB.Image imgSemD 
      Height          =   465
      Left            =   2670
      Picture         =   "frmSimA.frx":000C
      Top             =   1335
      Width           =   195
   End
   Begin VB.Image imgSemC 
      Height          =   465
      Left            =   1965
      Picture         =   "frmSimA.frx":0526
      Top             =   2310
      Width           =   195
   End
   Begin VB.Image imgSemB 
      Height          =   195
      Left            =   1695
      Picture         =   "frmSimA.frx":0A40
      Top             =   1605
      Width           =   465
   End
   Begin VB.Image imgSemA 
      Height          =   195
      Left            =   2655
      Picture         =   "frmSimA.frx":0F62
      Top             =   2325
      Width           =   465
   End
End
Attribute VB_Name = "frmSimA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TOTAL_CARROS As Long 'Total de carros no universo

Private semA As Boolean
Private sinA As Long

Private semB As Boolean
Private sinB As Long

Private carid   As Long

Private nodos(0 To 169) As Object

Private carros As Collection

Private fila30 As Collection
Private fila47 As Collection
Private fila117 As Collection
Private fila136 As Collection

'**********************************************
'* INICIALIZAÇÃO DO AMBIENTE
Public Sub Inicializa_Mundo()

    Dim i As Integer
    Dim rota As clsRota
    
    semA = False
    sinA = semaforo.VERMELHO
    semB = True
    sinB = semaforo.VERDE

    tsemA.Enabled = True
    tsemB.Enabled = True
    Tick.Enabled = True
    geraCarro.Enabled = True
    
    Set fila30 = New Collection
    Set fila47 = New Collection
    Set fila117 = New Collection
    Set fila136 = New Collection
    Set carros = New Collection
    
    '**********************************************
    '* INSTANCIA TODOS OS NODOS DA SIMULAÇÃO
    
    For i = 0 To 169
        Set nodos(i) = New clsNodo
    Next i
            
    '**********************************************
    '* CRIA ROTEADORES E TABELAS DE ROTAS
    
    Set nodos(23) = New clsRoteador
    Set nodos(24) = New clsRoteador
    Set nodos(53) = New clsRoteador
    Set nodos(54) = New clsRoteador
    
    'R1
    Set nodos(23).picture = pista(23)
    nodos(23).ocupado = False
    nodos(23).numero = 23
    
    Set rota = New clsRota
    rota.destino = 131
    Set rota.proximo = nodos(53)
    nodos(23).rotas.Add rota
    
    Set rota = New clsRota
    rota.destino = 17
    Set rota.proximo = nodos(22)
    nodos(23).rotas.Add rota

    Set rota = New clsRota
    rota.destino = 60
    Set rota.proximo = nodos(53)
    nodos(23).rotas.Add rota

    Set rota = New clsRota
    rota.destino = 131
    Set rota.proximo = nodos(22)
    nodos(23).rotas.Add rota

    
    'R2
    Set nodos(24).picture = pista(24)
    nodos(24).ocupado = False
    
    Set rota = New clsRota
    rota.destino = 122
    Set rota.proximo = nodos(126)
    nodos(24).rotas.Add rota
    
    Set rota = New clsRota
    rota.destino = 17
    Set rota.proximo = nodos(23)
    nodos(24).rotas.Add rota
    
    Set rota = New clsRota
    rota.destino = 17
    Set rota.proximo = nodos(126)
    nodos(24).rotas.Add rota
    
    Set rota = New clsRota
    rota.destino = 131
    Set rota.proximo = nodos(23)
    nodos(24).rotas.Add rota
     
    'R3
    Set nodos(53).picture = pista(53)
    nodos(53).ocupado = False
    nodos(53).numero = 53
    
    Set rota = New clsRota
    rota.destino = 60
    Set rota.proximo = nodos(54)
    nodos(53).rotas.Add rota
    
    Set rota = New clsRota
    rota.destino = 131
    Set rota.proximo = nodos(127)
    nodos(53).rotas.Add rota
    
    Set rota = New clsRota
    rota.destino = 60
    Set rota.proximo = nodos(127)
    nodos(53).rotas.Add rota
    
    Set rota = New clsRota
    rota.destino = 122
    Set rota.proximo = nodos(54)
    nodos(53).rotas.Add rota
    
    'R4
    Set nodos(54).picture = pista(54)
    nodos(54).ocupado = False
    
    Set rota = New clsRota
    rota.destino = 60
    Set rota.proximo = nodos(55)
    nodos(54).rotas.Add rota
    
    Set rota = New clsRota
    rota.destino = 122
    Set rota.proximo = nodos(24)
    nodos(54).rotas.Add rota
    
    Set rota = New clsRota
    rota.destino = 17
    Set rota.proximo = nodos(24)
    nodos(54).rotas.Add rota
    
    Set rota = New clsRota
    rota.destino = 122
    Set rota.proximo = nodos(60)
    nodos(54).rotas.Add rota
    
    '**********************************************
    
    '* A PARTIR DAQUI, CRIA AS LIGAÇÕES ENTRE NODOS
    Set nodos(17).proximo = nodos(38)
    Set nodos(17).picture = pista(17)
    
    Set nodos(18).proximo = nodos(17)
    Set nodos(18).picture = pista(18)
            
    Set nodos(19).proximo = nodos(18)
    Set nodos(19).picture = pista(19)
        
    Set nodos(20).proximo = nodos(19)
    Set nodos(20).picture = pista(20)
        
    Set nodos(21).proximo = nodos(20)
    Set nodos(21).picture = pista(21)
        
    Set nodos(22).proximo = nodos(21)
    Set nodos(22).picture = pista(22)
        
    Set nodos(25).proximo = nodos(24)
    Set nodos(25).picture = pista(25)
        
    Set nodos(26).proximo = nodos(25)
    Set nodos(26).picture = pista(26)
        
    Set nodos(27).proximo = nodos(26)
    Set nodos(27).picture = pista(27)
        
    Set nodos(28).proximo = nodos(27)
    Set nodos(28).picture = pista(28)
        
    Set nodos(29).proximo = nodos(28)
    Set nodos(29).picture = pista(29)
        
    Set nodos(30).proximo = nodos(29)
    Set nodos(30).picture = pista(30)
        
    Set nodos(32).proximo = nodos(75)
    Set nodos(32).picture = pista(32)
        
    Set nodos(33).proximo = nodos(32)
    Set nodos(33).picture = pista(33)
        
    Set nodos(34).proximo = nodos(33)
    Set nodos(34).picture = pista(34)
        
    Set nodos(35).proximo = nodos(34)
    Set nodos(35).picture = pista(35)
        
    Set nodos(36).proximo = nodos(35)
    Set nodos(36).picture = pista(36)
        
    Set nodos(37).proximo = nodos(36)
    Set nodos(37).picture = pista(37)
    
    Set nodos(38).proximo = nodos(37)
    Set nodos(38).picture = pista(38)
    
    Set nodos(39).proximo = nodos(47)
    Set nodos(39).picture = pista(39)
    
    Set nodos(40).proximo = nodos(39)
    Set nodos(40).picture = pista(40)
    
    Set nodos(41).proximo = nodos(40)
    Set nodos(41).picture = pista(41)
    
    Set nodos(42).proximo = nodos(41)
    Set nodos(42).picture = pista(42)
    
    Set nodos(43).proximo = nodos(42)
    Set nodos(43).picture = pista(43)
    
    Set nodos(44).proximo = nodos(43)
    Set nodos(44).picture = pista(44)
    
    Set nodos(45).proximo = nodos(44)
    Set nodos(45).picture = pista(45)
    
    Set nodos(47).proximo = nodos(48)
    Set nodos(47).picture = pista(47)
    
    Set nodos(48).proximo = nodos(49)
    Set nodos(48).picture = pista(48)
    
    Set nodos(49).proximo = nodos(50)
    Set nodos(49).picture = pista(49)
    
    Set nodos(50).proximo = nodos(51)
    Set nodos(50).picture = pista(50)
    
    Set nodos(51).proximo = nodos(52)
    Set nodos(51).picture = pista(51)
    
    Set nodos(52).proximo = nodos(53)
    Set nodos(52).picture = pista(52)
        
    Set nodos(55).proximo = nodos(56)
    Set nodos(55).picture = pista(55)
        
    Set nodos(56).proximo = nodos(57)
    Set nodos(56).picture = pista(56)
    
    Set nodos(57).proximo = nodos(58)
    Set nodos(57).picture = pista(57)
    
    Set nodos(58).proximo = nodos(59)
    Set nodos(58).picture = pista(58)
    
    Set nodos(59).proximo = nodos(60)
    Set nodos(59).picture = pista(59)
    
    Set nodos(60).proximo = nodos(145)
    Set nodos(60).picture = pista(60)
        
    Set nodos(75).proximo = nodos(76)
    Set nodos(75).picture = pista(75)
    
    Set nodos(76).proximo = nodos(77)
    Set nodos(76).picture = pista(76)
    
    Set nodos(77).proximo = nodos(78)
    Set nodos(77).picture = pista(77)
    
    Set nodos(78).proximo = nodos(79)
    Set nodos(78).picture = pista(78)
    
    Set nodos(79).proximo = nodos(80)
    Set nodos(79).picture = pista(79)
    
    Set nodos(80).proximo = nodos(81)
    Set nodos(80).picture = pista(80)
    
    Set nodos(81).proximo = nodos(117)
    Set nodos(81).picture = pista(81)
    
    Set nodos(82).proximo = nodos(83)
    Set nodos(82).picture = pista(82)
    
    Set nodos(83).proximo = nodos(84)
    Set nodos(83).picture = pista(83)
       
    Set nodos(84).proximo = nodos(85)
    Set nodos(84).picture = pista(84)
    
    Set nodos(85).proximo = nodos(86)
    Set nodos(85).picture = pista(85)
        
    Set nodos(86).proximo = nodos(87)
    Set nodos(86).picture = pista(86)
       
    Set nodos(87).proximo = nodos(88)
    Set nodos(87).picture = pista(87)
    
    Set nodos(88).proximo = nodos(138)
    Set nodos(88).picture = pista(88)
            
    Set nodos(89).proximo = nodos(45)
    Set nodos(89).picture = pista(89)
    
    Set nodos(90).proximo = nodos(89)
    Set nodos(90).picture = pista(90)
    
    Set nodos(91).proximo = nodos(90)
    Set nodos(91).picture = pista(91)
    
    Set nodos(92).proximo = nodos(91)
    Set nodos(92).picture = pista(92)
    
    Set nodos(93).proximo = nodos(92)
    Set nodos(93).picture = pista(93)
    
    Set nodos(94).proximo = nodos(93)
    Set nodos(94).picture = pista(94)
        
    Set nodos(95).proximo = nodos(94)
    Set nodos(95).picture = pista(95)
        
    Set nodos(96).proximo = nodos(136)
    Set nodos(96).picture = pista(96)
        
    Set nodos(97).proximo = nodos(96)
    Set nodos(97).picture = pista(97)
        
    Set nodos(98).proximo = nodos(97)
    Set nodos(98).picture = pista(98)
        
    Set nodos(99).proximo = nodos(98)
    Set nodos(99).picture = pista(99)
        
    Set nodos(100).proximo = nodos(99)
    Set nodos(100).picture = pista(100)
        
    Set nodos(101).proximo = nodos(100)
    Set nodos(101).picture = pista(101)
            
    Set nodos(102).proximo = nodos(101)
    Set nodos(102).picture = pista(102)
    
    Set nodos(117).proximo = nodos(118)
    Set nodos(117).picture = pista(117)
        
    Set nodos(118).proximo = nodos(119)
    Set nodos(118).picture = pista(118)
        
    Set nodos(119).proximo = nodos(120)
    Set nodos(119).picture = pista(119)
        
    Set nodos(120).proximo = nodos(121)
    Set nodos(120).picture = pista(120)
        
    Set nodos(121).proximo = nodos(23)
    Set nodos(121).picture = pista(121)
    
    Set nodos(122).proximo = nodos(82)
    Set nodos(122).picture = pista(122)
        
    Set nodos(123).proximo = nodos(122)
    Set nodos(123).picture = pista(123)
        
    Set nodos(124).proximo = nodos(123)
    Set nodos(124).picture = pista(124)
        
    Set nodos(125).proximo = nodos(124)
    Set nodos(125).picture = pista(125)
        
    Set nodos(126).proximo = nodos(125)
    Set nodos(126).picture = pista(126)
    
    Set nodos(127).proximo = nodos(128)
    Set nodos(127).picture = pista(127)
    
    Set nodos(128).proximo = nodos(129)
    Set nodos(128).picture = pista(128)
    
    Set nodos(129).proximo = nodos(130)
    Set nodos(129).picture = pista(129)
    
    Set nodos(130).proximo = nodos(131)
    Set nodos(130).picture = pista(130)
    
    Set nodos(131).proximo = nodos(95)
    Set nodos(131).picture = pista(131)
    
    Set nodos(132).proximo = nodos(54)
    Set nodos(132).picture = pista(132)
        
    Set nodos(133).proximo = nodos(132)
    Set nodos(133).picture = pista(133)
        
    Set nodos(134).proximo = nodos(133)
    Set nodos(134).picture = pista(134)
        
    Set nodos(135).proximo = nodos(134)
    Set nodos(135).picture = pista(135)
        
    Set nodos(136).proximo = nodos(135)
    Set nodos(136).picture = pista(136)
        
    Set nodos(138).proximo = nodos(139)
    Set nodos(138).picture = pista(138)
    
    Set nodos(139).proximo = nodos(140)
    Set nodos(139).picture = pista(139)
        
    Set nodos(140).proximo = nodos(141)
    Set nodos(140).picture = pista(140)
    
    Set nodos(141).proximo = nodos(142)
    Set nodos(141).picture = pista(141)
        
    Set nodos(142).proximo = nodos(143)
    Set nodos(142).picture = pista(142)
    
    Set nodos(143).proximo = nodos(144)
    Set nodos(143).picture = pista(143)
       
    Set nodos(144).proximo = nodos(30)
    Set nodos(144).picture = pista(144)
    
    Set nodos(145).proximo = nodos(146)
    Set nodos(145).picture = pista(145)
    
    Set nodos(146).proximo = nodos(147)
    Set nodos(146).picture = pista(146)
    
    Set nodos(147).proximo = nodos(148)
    Set nodos(147).picture = pista(147)
    
    Set nodos(148).proximo = nodos(149)
    Set nodos(148).picture = pista(148)
        
    Set nodos(149).proximo = nodos(150)
    Set nodos(149).picture = pista(149)
        
    Set nodos(150).proximo = nodos(151)
    Set nodos(150).picture = pista(150)
    
    Set nodos(150).proximo = nodos(151)
    Set nodos(150).picture = pista(150)
        
    Set nodos(151).proximo = nodos(102)
    Set nodos(151).picture = pista(151)
    
    nodos(25).semaforo = True
    nodos(52).semaforo = True
    nodos(121).semaforo = True
    nodos(132).semaforo = True
    
    nodos(25).aberto = semB
    nodos(52).aberto = semB
    nodos(121).aberto = semA
    nodos(132).aberto = semA

End Sub



Private Sub geraCarro_Timer()
    
    Dim carro As clsCarro
    
    geraCarro.Enabled = False
    
    If TOTAL_CARROS < MAX_CARROS Then
        Set carro = New clsCarro
        carid = carid + 1
        carro.carid = "CARID" & carid
        poeNaFila carro.origem, carro
        TOTAL_CARROS = TOTAL_CARROS + 1
    End If
    
    extraiFila
    
    geraCarro.Enabled = True
    
End Sub

Private Sub Tick_Timer()
    
    Dim tmp As Object
    
    Dim prox As Object
    Dim carro As clsCarro
    Dim i As Integer
    Tick.Enabled = False
    i = 1
    For Each carro In carros
        
        carro.nodo.picture.picture = LoadPicture("")
        carro.nodo.ocupado = False
        
        If Not carro.chegou Then
            ' Debug.Print carro.nodo.semaforo
            If (Not carro.nodo.semaforo) Or (carro.nodo.semaforo And carro.nodo.aberto) Then
                
                Set tmp = carro.nodo
                
                Set prox = carro.nodo.getProximo(carro.destino, nodos)
                If Not prox.ocupado Then
                    Set carro.nodo = prox
                End If
            End If
            
            carro.nodo.picture.picture = LoadPicture(App.Path & "\" & "carro_" & carro.num & ".bmp")
            carro.nodo.ocupado = True
        Else
            'Debug.Print carro.destino
            TOTAL_CARROS = TOTAL_CARROS - 1
            carro.saida = Time
            carros.Remove (carro.carid)
        End If
        i = i + 1
    Next carro
    
    Tick.Enabled = True
End Sub

Private Sub poeNaFila(idxNodo As Long, carro As clsCarro)

    Select Case idxNodo
        Case 30: fila30.Add carro, carro.carid
        Case 47: fila47.Add carro, carro.carid
        Case 117: fila117.Add carro, carro.carid
        Case 136: fila136.Add carro, carro.carid
    End Select

End Sub

Private Sub extraiFila()
    
    Dim nodo As Object
    Dim carro As clsCarro
    Dim i As Long
    
    i = 1
    For Each carro In fila30
        If Not nodos(30).ocupado Then
            Set carro.nodo = nodos(30)
            carro.entrada = Time
            carros.Add carro, carro.carid
            nodos(30).picture.picture = LoadPicture(App.Path & "\" & "carro_" & carro.num & ".bmp")
            fila30.Remove (carro.carid)
        End If
        i = i + 1
    Next carro
    
    i = 1
    For Each carro In fila47
        If Not nodos(47).ocupado Then
            Set carro.nodo = nodos(47)
            carro.entrada = Time
            carros.Add carro, carro.carid
            nodos(47).picture.picture = LoadPicture(App.Path & "\" & "carro_" & carro.num & ".bmp")
            fila47.Remove (carro.carid)
        End If
        i = i + 1
    Next carro
    
    i = 1
    For Each carro In fila117
        If Not nodos(117).ocupado Then
            Set carro.nodo = nodos(117)
            carro.entrada = Time
            carros.Add carro, carro.carid
            nodos(117).picture.picture = LoadPicture(App.Path & "\" & "carro_" & carro.num & ".bmp")
            fila117.Remove (carro.carid)
        End If
        i = i + 1
    Next carro
    
    i = 1
    For Each carro In fila136
        If Not nodos(136).ocupado Then
            Set carro.nodo = nodos(136)
            carro.entrada = Time
            carros.Add carro, carro.carid
            nodos(136).picture.picture = LoadPicture(App.Path & "\" & "carro_" & carro.num & ".bmp")
            fila136.Remove (carro.carid)
        End If
        i = i + 1
    Next carro

End Sub

Private Sub tsemA_Timer()
    
    tsemA.Enabled = False
        
    Select Case sinA
        Case semaforo.VERMELHO:
            If semA Then
                sinA = semaforo.VERDE
                Set imgSemA.picture = LoadPicture(App.Path & "\verde_1.bmp")
                Set imgSemB.picture = LoadPicture(App.Path & "\verde_2.bmp")
            End If
        Case semaforo.AMARELO:
            If Not semA Then
                sinA = semaforo.VERMELHO
                Set imgSemA.picture = LoadPicture(App.Path & "\vermelho_1.bmp")
                Set imgSemB.picture = LoadPicture(App.Path & "\vermelho_2.bmp")
            End If
        Case semaforo.VERDE:
            If semA Then
                sinA = semaforo.AMARELO
                Set imgSemA.picture = LoadPicture(App.Path & "\amarelo_1.bmp")
                Set imgSemB.picture = LoadPicture(App.Path & "\amarelo_2.bmp")
            End If
    End Select
        
    Select Case sinB
        Case semaforo.VERMELHO:
            If semB Then
                sinB = semaforo.VERDE
                Set imgSemC.picture = LoadPicture(App.Path & "\verde_3.bmp")
                Set imgSemD.picture = LoadPicture(App.Path & "\verde_4.bmp")
            End If
        Case semaforo.AMARELO:
            If Not semB Then
                sinB = semaforo.VERMELHO
                Set imgSemC.picture = LoadPicture(App.Path & "\vermelho_3.bmp")
                Set imgSemD.picture = LoadPicture(App.Path & "\vermelho_4.bmp")
            End If
        Case semaforo.VERDE:
            If semB Then
                sinB = semaforo.AMARELO
                Set imgSemC.picture = LoadPicture(App.Path & "\amarelo_3.bmp")
                Set imgSemD.picture = LoadPicture(App.Path & "\amarelo_4.bmp")
            End If
    End Select
   
    tsemA.Enabled = True
    End Sub

Private Sub tsemB_Timer()
    If semA Then
        semA = False
        semB = True
    Else
        semA = True
        semB = False
    End If
    
    tsemA_Timer
    
    nodos(25).aberto = semB
    nodos(52).aberto = semB
    nodos(121).aberto = semA
    nodos(132).aberto = semA
    
End Sub

Public Sub pausa()
    Tick.Enabled = False
    geraCarro = False
    tsemA = False
    tsemB = False
End Sub

Public Sub continua()
    Tick.Enabled = True
    geraCarro = True
    tsemA = True
    tsemB = True
End Sub
