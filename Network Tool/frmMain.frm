VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Network Utility"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7890
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   7890
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock tcpPing 
      Left            =   6810
      Top             =   4005
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrPing 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6405
      Top             =   4050
   End
   Begin VB.Timer tmrCurrent 
      Interval        =   1000
      Left            =   7230
      Top             =   4050
   End
   Begin TabDlg.SSTab MainTab 
      Height          =   4065
      Left            =   45
      TabIndex        =   22
      Top             =   0
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   7170
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   14737632
      TabCaption(0)   =   "IP"
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdStopPing"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdPrint"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdClear"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "TxtRemotePort"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtIp_Scan(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtIp_Scan(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtIp_Scan(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtIp_Scan(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtPacketSize"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtEcho"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtPortAddress"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtLag"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtStatus"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtUserName"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtDataSend"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtIp_IP(3)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtIp_IP(2)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtIp_IP(1)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmdPing"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtIp_IP(0)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "rtxtIP"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label9"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "lblRemotePort"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lblDataSend"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "lblUserName"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "lblPing"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "lblLag"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "lblPort"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "lblSize"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "lblServerStatus"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "lblEcho"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).ControlCount=   31
      TabCaption(1)   =   "Port"
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblTo"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label5"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdClearList"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdScan"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdStop"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lstOpenPorts"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtBegPort"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtEndPort"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtIP_Port(1)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtIP_Port(0)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtIP_Port(2)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtIP_Port(3)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Frame1"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).ControlCount=   15
      TabCaption(2)   =   "Sys Info"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Whois"
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtDomainName"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "CmdWhois"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "TxtWhois"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "txtInfoSource"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label14"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Label13"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Label1"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).ControlCount=   7
      Begin VB.CommandButton cmdStopPing 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Stop"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -69885
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   3450
         Width           =   930
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Port Listing"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3510
         Left            =   4155
         TabIndex        =   65
         Top             =   480
         Width           =   3675
         Begin VB.ListBox lstPorts 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   3000
            Left            =   120
            TabIndex        =   66
            Top             =   360
            Width           =   3525
         End
      End
      Begin VB.TextBox txtIP_Port 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   3
         Left            =   3540
         MaxLength       =   3
         TabIndex        =   13
         Top             =   585
         Width           =   540
      End
      Begin VB.TextBox txtIP_Port 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   2
         Left            =   2985
         MaxLength       =   3
         TabIndex        =   12
         Top             =   585
         Width           =   540
      End
      Begin VB.TextBox txtIP_Port 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   0
         Left            =   1875
         MaxLength       =   3
         TabIndex        =   10
         Top             =   585
         Width           =   540
      End
      Begin VB.TextBox txtIP_Port 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   1
         Left            =   2430
         MaxLength       =   3
         TabIndex        =   11
         Top             =   585
         Width           =   540
      End
      Begin VB.TextBox txtEndPort 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   3180
         TabIndex        =   15
         Text            =   "500"
         Top             =   975
         Width           =   900
      End
      Begin VB.TextBox txtBegPort 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1875
         TabIndex        =   14
         Text            =   "1"
         Top             =   975
         Width           =   900
      End
      Begin VB.ListBox lstOpenPorts 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1950
         Left            =   150
         TabIndex        =   47
         Top             =   1620
         Width           =   3930
      End
      Begin VB.CommandButton cmdStop 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Stop"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2190
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3630
         Width           =   930
      End
      Begin VB.CommandButton cmdScan 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Scan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1230
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3630
         Width           =   930
      End
      Begin VB.CommandButton cmdClearList 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3150
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3630
         Width           =   930
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -68040
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3450
         Width           =   855
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -68925
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3450
         Width           =   855
      End
      Begin VB.TextBox TxtRemotePort 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   -72120
         TabIndex        =   4
         Text            =   "139"
         ToolTipText     =   "Remote Port"
         Top             =   873
         Width           =   1095
      End
      Begin VB.TextBox txtIp_Scan 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   360
         Index           =   3
         Left            =   -68070
         MaxLength       =   3
         TabIndex        =   6
         Top             =   510
         Width           =   540
      End
      Begin VB.TextBox txtIp_Scan 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   360
         Index           =   2
         Left            =   -68625
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   510
         Width           =   540
      End
      Begin VB.TextBox txtIp_Scan 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   360
         Index           =   1
         Left            =   -69180
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   510
         Width           =   540
      End
      Begin VB.TextBox txtIp_Scan 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   360
         Index           =   0
         Left            =   -69735
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   510
         Width           =   540
      End
      Begin VB.TextBox txtPacketSize 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   -73230
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   3420
         Width           =   2205
      End
      Begin VB.TextBox txtEcho 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   -73230
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   3051
         Width           =   2205
      End
      Begin VB.TextBox txtPortAddress 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   -73230
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   2688
         Width           =   2205
      End
      Begin VB.TextBox txtLag 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   -73230
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   2325
         Width           =   2205
      End
      Begin VB.TextBox txtStatus 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   -73230
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1962
         Width           =   2205
      End
      Begin VB.TextBox txtUserName 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   -73230
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1599
         Width           =   2205
      End
      Begin VB.TextBox txtDataSend 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   -73230
         TabIndex        =   5
         Top             =   1236
         Width           =   2205
      End
      Begin VB.TextBox txtIp_IP 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   360
         Index           =   3
         Left            =   -71565
         MaxLength       =   3
         TabIndex        =   3
         Top             =   510
         Width           =   540
      End
      Begin VB.TextBox txtIp_IP 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   360
         Index           =   2
         Left            =   -72120
         MaxLength       =   3
         TabIndex        =   2
         Top             =   510
         Width           =   540
      End
      Begin VB.TextBox txtIp_IP 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   360
         Index           =   1
         Left            =   -72675
         MaxLength       =   3
         TabIndex        =   1
         Top             =   510
         Width           =   540
      End
      Begin VB.TextBox txtDomainName 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   -72315
         TabIndex        =   19
         Top             =   615
         Width           =   3540
      End
      Begin VB.CommandButton CmdWhois 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Whois"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -68295
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3645
         Width           =   930
      End
      Begin VB.TextBox TxtWhois 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1920
         Left            =   -74805
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   38
         Top             =   1710
         Width           =   7410
      End
      Begin VB.ComboBox txtInfoSource 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   -72315
         TabIndex        =   20
         Text            =   "rs.internic.net"
         Top             =   990
         Width           =   3540
      End
      Begin VB.CommandButton cmdPing 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ping"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -70845
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3450
         Width           =   930
      End
      Begin VB.TextBox txtIp_IP 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   360
         Index           =   0
         Left            =   -73215
         MaxLength       =   3
         TabIndex        =   0
         Top             =   510
         Width           =   540
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Height          =   2580
         Left            =   -73560
         TabIndex        =   52
         Top             =   720
         Width           =   4365
         Begin VB.TextBox txtMyOS 
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   2055
            Locked          =   -1  'True
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   1710
            Width           =   2175
         End
         Begin VB.TextBox txtMyCPUName 
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   2055
            Locked          =   -1  'True
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   975
            Width           =   2175
         End
         Begin VB.TextBox txtMyIP 
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   2055
            Locked          =   -1  'True
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   1335
            Width           =   2175
         End
         Begin VB.TextBox txtMyUserName 
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   2055
            Locked          =   -1  'True
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox txtMyEthernet 
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   2055
            Locked          =   -1  'True
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   2085
            Width           =   2175
         End
         Begin VB.TextBox txtMyHostName 
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   2055
            Locked          =   -1  'True
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   225
            Width           =   2175
         End
         Begin VB.Label lblName 
            BackStyle       =   0  'Transparent
            Caption         =   "Computer Name:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   150
            TabIndex        =   64
            Top             =   1020
            Width           =   1665
         End
         Begin VB.Label lblCompIP 
            BackStyle       =   0  'Transparent
            Caption         =   "Computer IP:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   150
            TabIndex        =   63
            Top             =   1395
            Width           =   1395
         End
         Begin VB.Label lblUser 
            BackStyle       =   0  'Transparent
            Caption         =   "Username:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   150
            TabIndex        =   62
            Top             =   645
            Width           =   1095
         End
         Begin VB.Label lblOS 
            BackStyle       =   0  'Transparent
            Caption         =   "Operating System:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   150
            TabIndex        =   61
            Top             =   1770
            Width           =   1815
         End
         Begin VB.Label lblNet 
            BackStyle       =   0  'Transparent
            Caption         =   "Ethernet Address:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   150
            TabIndex        =   60
            Top             =   2145
            Width           =   1875
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Host Name:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   150
            TabIndex        =   59
            Top             =   285
            Width           =   1155
         End
      End
      Begin RichTextLib.RichTextBox rtxtIP 
         Height          =   2490
         Left            =   -70875
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   930
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   4392
         _Version        =   393217
         BackColor       =   14737632
         ScrollBars      =   2
         TextRTF         =   $"frmMain.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Open Ports:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   150
         TabIndex        =   51
         Top             =   1395
         Width           =   1185
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Port Range:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   150
         TabIndex        =   50
         Top             =   1020
         Width           =   1395
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Remote IP:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   150
         TabIndex        =   49
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label lblTo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2790
         TabIndex        =   48
         Top             =   1035
         Width           =   375
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Info Source:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74805
         TabIndex        =   46
         Top             =   1043
         Width           =   1935
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Search Information:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74805
         TabIndex        =   45
         Top             =   1470
         Width           =   2355
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Ending IP:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   -70740
         TabIndex        =   44
         Top             =   555
         Width           =   1095
      End
      Begin VB.Label lblRemotePort 
         BackStyle       =   0  'Transparent
         Caption         =   "Port:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   -74820
         TabIndex        =   43
         Top             =   915
         Width           =   585
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Remote Address:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74805
         TabIndex        =   39
         Top             =   683
         Width           =   1935
      End
      Begin VB.Label lblDataSend 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Data to Send:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74820
         TabIndex        =   36
         Top             =   1275
         Width           =   1365
      End
      Begin VB.Label lblUserName 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74820
         TabIndex        =   35
         Top             =   1650
         Width           =   1455
      End
      Begin VB.Label lblPing 
         BackStyle       =   0  'Transparent
         Caption         =   "Remote IP:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   -74820
         TabIndex        =   34
         Top             =   555
         Width           =   1095
      End
      Begin VB.Label lblLag 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Lag Time:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74820
         TabIndex        =   33
         Top             =   2385
         Width           =   1110
      End
      Begin VB.Label lblPort 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Port Address:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74820
         TabIndex        =   32
         Top             =   2745
         Width           =   1440
      End
      Begin VB.Label lblSize 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Packet Size:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74820
         TabIndex        =   31
         Top             =   3480
         Width           =   1320
      End
      Begin VB.Label lblServerStatus 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74820
         TabIndex        =   30
         Top             =   2010
         Width           =   675
      End
      Begin VB.Label lblEcho 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Echo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74820
         TabIndex        =   29
         Top             =   3105
         Width           =   645
      End
   End
   Begin MSComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   37
      Top             =   4065
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8731
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'port scan variables
Dim NextPort As Double
Dim boolPortScan As Boolean
Dim strPortIP As String

'ip scan port
Dim IPBegScan As Integer
Dim IPEndScan As Integer
Dim strIPScan As String
Dim boolIPScan As Boolean

''''''''Winsock States''''''''''

'0 sckClosed
'1 sckOpen
'2 sckListening
'3 sckConnectionPending
'4 sckResolvingHost
'5 sckHostResolved
'6 sckConnecting
'7 sckConnected
'8 sckClosing
'9 sckError

''''''''''''''''''''''''''''''''''GENERAL FORM'''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()

    Dim strFile As String                                       'used for input file for port listing
    CpuInfo                                                     'call cpu info
    sbMain.Panels(2).Text = Format(Date, "DD MMM YYYY")         'set status bar date
        
    Open App.Path & "\portlist.txt" For Input As #1             ' Open file for input.
    Do While Not EOF(1)                                         '** Loop until end of file.
        Input #1, strFile                                       '** Read data into two variables.
        lstPorts.AddItem (strFile)
    Loop
    Close #1                                                    'close connection

End Sub

Private Sub tmrPing_Timer()
    'timer controls movement between port/ip scans
    'looks to see what you are doing via what tab is visible
    If MainTab.Tab = 0 Then
        sbMain.Panels(1).Text = "IP Index: " & IPBegScan
        ScanIP
    Else
        sbMain.Panels(1).Text = "Current Port: " & NextPort
        PortScan
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    End
End Sub

Private Sub mnuExit_Click()
    Unload Me
    End
End Sub

Private Sub tcpPing_Close()
   tcpPing.Close
End Sub

Private Sub tcpPing_Connect()
        
    On Error GoTo errorTCP
    
                                                                            'based on tab perform certain procedures
    If MainTab.Tab = 0 Then
        rtxtIP.Text = rtxtIP.Text & ("IP OK:  " & tcpPing.RemoteHost _
                        & "   " & IPtoDNS(tcpPing.RemoteHost)) & vbCrLf
        ScanIP                                                              'call scanip sub
    ElseIf MainTab.Tab = 1 Then
        lstOpenPorts.AddItem "Port: " + Str(Me.tcpPing.RemotePort)
        PortScan                                                            'call portscan sub
    ElseIf MainTab.Tab = 3 Then
        TxtWhois = "Connected to " & txtInfoSource.Text & vbCrLf
    End If
    Exit Sub

errorTCP:
    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & vbCrLf & Err.Description
        Resume Next
    End If

End Sub

Private Sub MainTab_Click(PreviousTab As Integer)
    
    If MainTab.Tab = 0 Then
        txtIp_IP(0).SetFocus
    ElseIf MainTab.Tab = 1 Then
        txtIP_Port(0).SetFocus
        cmdScan.Enabled = True
        cmdStop.Enabled = False
    ElseIf MainTab.Tab = 3 Then
        txtDomainName.SetFocus
    End If
    sbMain.Panels(1).Text = ""
            
End Sub

Private Sub tcpPing_DataArrival(ByVal bytesTotal As Long)
    
    On Error GoTo errorTCP
    
    Dim strMessage As String
    
    tcpPing.GetData strMessage
    If MainTab.Tab = 0 Then                                 'receive data then put it into corresponding textboxes
        rtxtIP.Text = rtxtIP.Text & ("Incoming : " & strMessage) & vbCrLf
        tcpPing.Close
    ElseIf MainTab.Tab = 1 Then
        lstOpenPorts.AddItem ("Incoming : " & strMessage)
        tcpPing.Close
    ElseIf MainTab.Tab = 3 Then
        TxtWhois.Text = TxtWhois.Text & strMessage
    End If
    Exit Sub
    
errorTCP:
    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & vbCrLf & Err.Description
        tcpPing.Close
        Resume Next
    End If
    
End Sub

Private Sub tcpPing_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    If MainTab.Tab = 0 Then                                 'on error with scanning ips or ports just move to the next one
        ScanIP
    ElseIf MainTab.Tab = 1 Then
        PortScan
    Else
        tcpPing.Close
    End If
End Sub

Private Sub tmrCurrent_Timer()
    sbMain.Panels(3).Text = Time
End Sub

''''''''''''''''''''''''''''''''''IP TAB'''''''''''''''''''''''''''''''''''''''''
Private Sub cmdClear_Click()
    rtxtIP.Text = ""
    txtDataSend.Text = ""
    txtUserName.Text = ""
    txtStatus.Text = ""
    txtLag.Text = ""
    txtPortAddress.Text = ""
    txtEcho.Text = ""
    txtPacketSize.Text = ""
    txtIp_Scan(3).Text = ""
    sbMain.Panels(1).Text = ""
    txtIp_IP(3).SetFocus
End Sub

Private Sub cmdPrint_Click()
    rtxtIP.SelPrint (Printer.hDC)
End Sub

Private Sub cmdPing_Click()
        
    On Error GoTo errorTCP

    Dim intCount As Integer
    For intCount = 0 To 3
        If txtIp_IP(intCount).Text = "" Then
            MsgBox "You must enter a valid IP Address", vbOKOnly, "Invalid IP"
            Exit Sub
        End If
    Next intCount
    
    If TxtRemotePort.Text = "" Then
        MsgBox "You must enter a valid port address", vbOKOnly, "Invalid Port"
        Exit Sub
    End If
    
    sbMain.Panels(1).Text = ""
    If txtIp_Scan(3).Text = "" Then                         'if there is not second ip then perform a ping else perform range ip scan
        PingSingle
    Else
        cmdPing.Enabled = False                             'put together ips
        cmdStopPing.Enabled = True
        sbMain.Panels(1).Text = "Begin IP Scan"
        rtxtIP.Text = ""
        rtxtIP.Text = rtxtIP.Text & "RemotePort: " & TxtRemotePort.Text & vbCrLf
        strIPScan = txtIp_IP(0).Text & "." & txtIp_IP(1).Text & "." & txtIp_IP(2).Text
        IPBegScan = txtIp_IP(3).Text
        IPEndScan = txtIp_Scan(3).Text
        boolIPScan = True                                   'set a switch for ip scanning - will be used when stopping
        tmrPing.Enabled = True                              'start ip scanning
    End If
    Exit Sub
    
errorTCP:
    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & vbCrLf & Err.Description
        Resume Next
    End If
    
End Sub

Private Sub PingSingle()

    On Error GoTo errorTCP

    Dim ECHO As ICMP_ECHO_REPLY
    Dim pos As Integer
    Dim strIP As String
    Dim intCount As Integer
        
    cmdPing.Enabled = False
    
                                                                'put ip address together
    strIP = txtIp_IP(0).Text & "." & txtIp_IP(1).Text & _
                "." & txtIp_IP(2).Text & "." & txtIp_IP(3).Text
    
    sbMain.Panels(1).Text = "Pinging  " & strIP                 'return data from ping
    Call Ping(strIP, txtDataSend.Text, ECHO)
    txtUserName.Text = IPtoDNS(strIP)
    txtStatus.Text = GetStatusCode(ECHO.status)
    txtLag.Text = ECHO.RoundTripTime & " milliseconds"
    txtPortAddress.Text = ECHO.Address
    txtEcho.Text = ECHO.Data & " Data"
    txtPacketSize.Text = ECHO.DataSize & " bytes"
    If Left$(ECHO.Data, 1) <> Chr$(0) Then
        pos = InStr(ECHO.Data, Chr$(0))
        txtEcho.Text = Left$(ECHO.Data, pos - 1)
    End If
    sbMain.Panels(1).Text = "Pinging  " & strIP & "  complete."
    
    cmdPing.Enabled = True
    Exit Sub

errorTCP:
    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & vbCrLf & Err.Description
        Resume Next
    End If
End Sub

''''''''''''''''''''''''''''''''''PORT SCAN TAB'''''''''''''''''''''''''''''''''''''''''
Private Sub cmdClearList_Click()
   lstOpenPorts.Clear
   sbMain.Panels(1).Text = ""
End Sub

Private Sub cmdScan_Click()

    On Error GoTo errorTCP
    
    Dim intCount As Integer
    
    cmdScan.Enabled = False
    cmdStop.Enabled = True
    sbMain.Panels(1).Text = ""
    For intCount = 0 To 3                                       'verify ip
        If txtIP_Port(intCount).Text = "" Then
            MsgBox "You must enter a valid IP Address", vbOKOnly, "Invalid IP"
            Exit Sub
        End If
    Next intCount
    
                                                                'put together ip
    strPortIP = txtIP_Port(0).Text & "." & txtIP_Port(1).Text & "." & _
                txtIP_Port(2).Text & "." & txtIP_Port(3).Text

    lstOpenPorts.Clear
    NextPort = txtBegPort.Text                                  'set starting port
    lstOpenPorts.AddItem "Initializing Port Scan"
    boolPortScan = True                                         'set switch - used for stopping port scan
    tmrPing.Enabled = True                                      'start scan
    Exit Sub
    
errorTCP:
    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & vbCrLf & Err.Description
        Resume Next
    End If
    
    
End Sub

Private Sub PortScan()

    On Error GoTo errorTCP
    
    If NextPort <= txtEndPort.Text And boolPortScan = True Then 'if we reached end point or stop was pushed
        DoEvents                                                'important to release for arrival procedure
        tcpPing.Close                                           'be sure it is not already open
        NextPort = NextPort + 1                                 'increment ports
        tcpPing.RemoteHost = strPortIP                          'set ip
        tcpPing.RemotePort = NextPort                           'set the port
        tcpPing.Connect                                         'connect
    Else
        tcpPing.Close                                           'if stopped enable buttons
        cmdScan.Enabled = True
        cmdStop.Enabled = False
        tmrPing.Enabled = False
    End If
    Exit Sub
    
errorTCP:
    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & vbCrLf & Err.Description
        Resume Next
    End If
End Sub

Private Sub cmdStop_Click()
    boolPortScan = False                                        'set bool switch to stop port scan
End Sub

''''''''''''''''''''''''''''''''''IP SCAN'''''''''''''''''''''''''''''''''''''''''

Private Sub ScanIP()

    On Error GoTo errorTCP
    
    If IPBegScan <= IPEndScan And boolIPScan = True Then        'if we reached end point or stop was pushed
        DoEvents                                                'important to release for arrival procedure
        If tcpPing.State <> sckClosed Then tcpPing.Close        'if socket isn't closed then close it before setting props
        IPBegScan = IPBegScan + 1                               'increment ip
        tcpPing.RemoteHost = strIPScan & "." & IPBegScan        'set ip
        tcpPing.RemotePort = TxtRemotePort.Text                 'set port
        tcpPing.Connect                                         'connect
    Else
        cmdPing.Enabled = True                                  'if stopped enable buttons
        cmdStopPing.Enabled = False
        tmrPing.Enabled = False
        If IPBegScan <= IPEndScan Then                          'if we stopped before ending scan then set text
            sbMain.Panels(1).Text = "IP Scan Stopped at " & IPBegScan
        Else
            sbMain.Panels(1).Text = "IP Scan Complete"
        End If
    End If
    Exit Sub
    
errorTCP:
    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & vbCrLf & Err.Description
        Resume Next
    End If
End Sub

Private Sub cmdStopPing_Click()
    boolIPScan = False                                      'ip scan bool for testing if we user wants to stop
End Sub

Public Function IPtoDNS(ByVal strAddress As String) As String
    
    On Error GoTo errorTCP
    
    Dim Host As HOSTENT
    Dim lAddress As Long
    Dim lTemp As Long
    Dim strHostName As String
    
    lAddress = inet_addr(strAddress)
    lTemp = gethostbyaddr(lAddress, 4, PF_INET)
    If lTemp <> 0 Then
        CopyMemory Host, ByVal lTemp, Len(Host)
        strHostName = String(256, 0)
        CopyMemory ByVal strHostName, ByVal Host.hName, 256
        If strHostName = "" Then
            IPtoDNS = "DNS error : resolution impossible " & Str$(WSAGetLastError())
        Else
            IPtoDNS = Left(strHostName, InStr(strHostName, Chr(0)) - 1)
        End If
    Else
        IPtoDNS = "Unable to determine name"
    End If
    Exit Function
    
errorTCP:
    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & vbCrLf & Err.Description
        Resume Next
    End If
    
End Function

''''''''''''''''''''''''''''''''''WHOIS TAB'''''''''''''''''''''''''''''''''''''''''
Private Sub CmdWhois_Click()
    
    tcpPing.RemotePort = 43                                     'port 43 is who is
    tcpPing.RemoteHost = txtInfoSource
    tcpPing.Connect
    While tcpPing.State <> sckConnected                         'loop until connected or errored out(can put timeout here)
        If tcpPing.State = sckError Then                        'if errored then exit
            TxtWhois.Text = TxtWhois.Text & " " & txtInfoSource.Text & " not responding."
            Exit Sub
        End If
        DoEvents
    Wend
    tcpPing.SendData txtDomainName.Text & vbCrLf                'if we reached this point then send name of address
    
End Sub


''''''''''''''''''''''''''''''''''SYS INFO TAB'''''''''''''''''''''''''''''''''''''''''
Private Sub CpuInfo()
 
    txtMyOS.Text = GetWindowsVersion()                          'api calls to pull cpu info
    txtMyCPUName.Text = ComputerName()
    txtMyIP.Text = tcpPing.LocalIP
    txtMyHostName.Text = tcpPing.LocalHostName
    txtMyUserName.Text = modUserName.UserName()
    txtMyEthernet.Text = EthernetAddress(0)
    If txtMyEthernet.Text = "000000000000" Then
        txtMyEthernet.Text = "No Ethernet Card Detected"
    End If
    
End Sub



''''''''''''''''''''''''''''''''''Textbox Controls'''''''''''''''''''''''''''''''''''''''''
Private Sub txtIp_IP_LostFocus(Index As Integer)
    If Index = 3 Then
        Dim intCount As Integer
        For intCount = 0 To 2
            txtIp_Scan(intCount).Text = txtIp_IP(intCount).Text
            txtIP_Port(intCount).Text = txtIp_IP(intCount).Text
        Next intCount
    End If
End Sub

Private Sub txtIP_Port_LostFocus(Index As Integer)
    If Index = 3 Then
        Dim intCount As Integer
        For intCount = 0 To 2
           txtIp_IP(intCount).Text = txtIP_Port(intCount).Text
        Next intCount
    End If
End Sub

Private Sub txtIP_Port_GotFocus(Index As Integer)
    Focus txtIP_Port(Index)
End Sub

Private Sub txtIP_Port_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = Asc(".") Then SendKeys "{TAB}"
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    KeyAscii = KeyCheck(KeyAscii, "Num")
End Sub

Private Sub txtBegPort_GotFocus()
    Focus txtBegPort
End Sub

Private Sub txtEndPort_GotFocus()
    Focus txtEndPort
End Sub

Private Sub txtBegPort_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    KeyAscii = KeyCheck(KeyAscii, "Num")
End Sub

Private Sub txtEndPort_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    KeyAscii = KeyCheck(KeyAscii, "Num")
End Sub

Private Sub txtIp_IP_GotFocus(Index As Integer)
    Focus txtIp_IP(Index)
End Sub

Private Sub txtDomainName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    KeyAscii = KeyCheck(KeyAscii, "Alpha")
End Sub

Private Sub TxtRemotePort_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    KeyAscii = KeyCheck(KeyAscii, "Num")
End Sub

Private Sub txtIp_IP_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = Asc(".") Then SendKeys "{TAB}"
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    KeyAscii = KeyCheck(KeyAscii, "Num")
End Sub

Private Sub txtIp_Scan_GotFocus(Index As Integer)
    Focus txtIp_Scan(Index)
End Sub

Private Sub txtIp_Scan_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = Asc(".") Then SendKeys "{TAB}"
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    KeyAscii = KeyCheck(KeyAscii, "Num")
End Sub

Private Sub TxtRemotePort_GotFocus()
    Focus TxtRemotePort
End Sub

Private Sub txtDataSend_GotFocus()
    Focus TxtRemotePort
End Sub

Private Sub txtDataSend_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    KeyAscii = KeyCheck(KeyAscii, "Alpha")
End Sub
