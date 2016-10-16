VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmConsole 
   BorderStyle     =   1  '單線固定
   Caption         =   "Camel's Scratch extension"
   ClientHeight    =   5370
   ClientLeft      =   30
   ClientTop       =   300
   ClientWidth     =   8355
   LinkTopic       =   "frmConsole"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   8355
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdUploadHEX 
      BackColor       =   &H00FFFF80&
      Caption         =   "上傳韌體"
      Height          =   525
      Left            =   6600
      Style           =   1  '圖片外觀
      TabIndex        =   92
      Top             =   1860
      Width           =   1455
   End
   Begin VB.CommandButton cmdSetThingSpeak 
      Caption         =   "設定ThingSpeak欄位值"
      Height          =   705
      Left            =   7080
      TabIndex        =   91
      Top             =   3060
      Width           =   1125
   End
   Begin VB.TextBox txtFV 
      Enabled         =   0   'False
      Height          =   270
      Index           =   8
      Left            =   7110
      TabIndex        =   90
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox txtFV 
      Enabled         =   0   'False
      Height          =   270
      Index           =   7
      Left            =   7110
      TabIndex        =   89
      Top             =   4620
      Width           =   1095
   End
   Begin VB.TextBox txtFV 
      Enabled         =   0   'False
      Height          =   270
      Index           =   6
      Left            =   5040
      TabIndex        =   86
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox txtFV 
      Enabled         =   0   'False
      Height          =   270
      Index           =   5
      Left            =   5040
      TabIndex        =   85
      Top             =   4620
      Width           =   1095
   End
   Begin VB.TextBox txtFV 
      Enabled         =   0   'False
      Height          =   270
      Index           =   4
      Left            =   2940
      TabIndex        =   82
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox txtFV 
      Enabled         =   0   'False
      Height          =   270
      Index           =   3
      Left            =   2940
      TabIndex        =   81
      Top             =   4620
      Width           =   1095
   End
   Begin VB.TextBox txtFV 
      Enabled         =   0   'False
      Height          =   270
      Index           =   2
      Left            =   840
      TabIndex        =   78
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox txtFV 
      Enabled         =   0   'False
      Height          =   270
      Index           =   1
      Left            =   840
      TabIndex        =   77
      Top             =   4620
      Width           =   1095
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "frmConsole.frx":0000
      Left            =   6060
      List            =   "frmConsole.frx":0002
      Style           =   2  '單純下拉式
      TabIndex        =   74
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox txtFieldVal 
      Height          =   270
      Left            =   6060
      TabIndex        =   71
      Top             =   3450
      Width           =   975
   End
   Begin VB.CommandButton cmdWriteThingSpeak 
      Caption         =   "寫入ThingSpeak"
      Height          =   645
      Left            =   4020
      TabIndex        =   70
      Top             =   3090
      Width           =   1065
   End
   Begin VB.TextBox txtAPIKey 
      Enabled         =   0   'False
      Height          =   255
      IMEMode         =   3  '暫止
      Left            =   1230
      PasswordChar    =   "*"
      TabIndex        =   53
      Top             =   3540
      Width           =   1575
   End
   Begin VB.CommandButton cmdReadThingSpeak 
      Caption         =   "讀出ThingSpeak"
      Height          =   645
      Left            =   2880
      TabIndex        =   50
      Top             =   3090
      Width           =   1095
   End
   Begin VB.TextBox txtChannelID 
      Enabled         =   0   'False
      Height          =   255
      Left            =   1230
      TabIndex        =   49
      Top             =   3240
      Width           =   1155
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   8160
      Top             =   810
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   975
      Left            =   3480
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   21
      Top             =   60
      Width           =   4575
   End
   Begin VB.CommandButton cmdArduino 
      BackColor       =   &H0080FF80&
      Caption         =   "開始連線"
      Height          =   315
      Left            =   2340
      Style           =   1  '圖片外觀
      TabIndex        =   19
      Top             =   780
      Width           =   1035
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1080
      Style           =   2  '單純下拉式
      TabIndex        =   6
      Top             =   780
      Width           =   1215
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   8160
      Top             =   150
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      BaudRate        =   57600
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   1080
      Max             =   100
      TabIndex        =   2
      Top             =   420
      Value           =   100
      Width           =   1935
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   1080
      Max             =   10
      Min             =   -10
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdEndPro 
      BackColor       =   &H000000FF&
      Caption         =   "結束離開"
      Height          =   525
      Left            =   6600
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   2400
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   6240
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "127.0.0.1"
      LocalPort       =   50001
   End
   Begin InetCtlsObjects.Inet Inet2 
      Left            =   8160
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label label 
      Alignment       =   2  '置中對齊
      Caption         =   "參數請填於Scratch.ini中"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   43
      Left            =   120
      TabIndex        =   93
      Top             =   3000
      Width           =   2625
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "field7"
      Height          =   195
      Index           =   42
      Left            =   6390
      TabIndex        =   88
      Top             =   4680
      Width           =   600
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "field8"
      Height          =   195
      Index           =   41
      Left            =   6390
      TabIndex        =   87
      Top             =   4980
      Width           =   600
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "field5"
      Height          =   195
      Index           =   40
      Left            =   4320
      TabIndex        =   84
      Top             =   4680
      Width           =   600
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "field6"
      Height          =   195
      Index           =   39
      Left            =   4320
      TabIndex        =   83
      Top             =   4980
      Width           =   600
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "field3"
      Height          =   195
      Index           =   38
      Left            =   2220
      TabIndex        =   80
      Top             =   4680
      Width           =   600
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "field4"
      Height          =   195
      Index           =   37
      Left            =   2220
      TabIndex        =   79
      Top             =   4980
      Width           =   600
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "field1"
      Height          =   195
      Index           =   36
      Left            =   120
      TabIndex        =   76
      Top             =   4680
      Width           =   600
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "field2"
      Height          =   195
      Index           =   35
      Left            =   120
      TabIndex        =   75
      Top             =   4980
      Width           =   600
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "Field No."
      Height          =   195
      Index           =   34
      Left            =   5250
      TabIndex        =   73
      Top             =   3150
      Width           =   795
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "Field Value"
      Height          =   195
      Index           =   33
      Left            =   5250
      TabIndex        =   72
      Top             =   3480
      Width           =   795
   End
   Begin VB.Label http 
      BorderStyle     =   1  '單線固定
      Height          =   255
      Index           =   7
      Left            =   7110
      TabIndex        =   69
      Top             =   3930
      Width           =   1095
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "field7"
      Height          =   195
      Index           =   32
      Left            =   6390
      TabIndex        =   68
      Top             =   3960
      Width           =   600
   End
   Begin VB.Label http 
      BorderStyle     =   1  '單線固定
      Height          =   255
      Index           =   6
      Left            =   5040
      TabIndex        =   67
      Top             =   4230
      Width           =   1095
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "field8"
      Height          =   195
      Index           =   31
      Left            =   6390
      TabIndex        =   66
      Top             =   4260
      Width           =   600
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "field6"
      Height          =   195
      Index           =   30
      Left            =   4320
      TabIndex        =   65
      Top             =   4260
      Width           =   600
   End
   Begin VB.Label http 
      BorderStyle     =   1  '單線固定
      Height          =   255
      Index           =   5
      Left            =   5040
      TabIndex        =   64
      Top             =   3930
      Width           =   1095
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "field5"
      Height          =   195
      Index           =   29
      Left            =   4320
      TabIndex        =   63
      Top             =   3960
      Width           =   600
   End
   Begin VB.Label http 
      BorderStyle     =   1  '單線固定
      Height          =   255
      Index           =   4
      Left            =   2940
      TabIndex        =   62
      Top             =   4230
      Width           =   1095
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "field4"
      Height          =   195
      Index           =   28
      Left            =   2220
      TabIndex        =   61
      Top             =   4260
      Width           =   600
   End
   Begin VB.Label http 
      BorderStyle     =   1  '單線固定
      Height          =   255
      Index           =   3
      Left            =   2940
      TabIndex        =   60
      Top             =   3930
      Width           =   1095
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "field3"
      Height          =   195
      Index           =   27
      Left            =   2220
      TabIndex        =   59
      Top             =   3960
      Width           =   600
   End
   Begin VB.Label http 
      BorderStyle     =   1  '單線固定
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   58
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "field2"
      Height          =   195
      Index           =   25
      Left            =   120
      TabIndex        =   57
      Top             =   4260
      Width           =   600
   End
   Begin VB.Label http 
      BorderStyle     =   1  '單線固定
      Height          =   255
      Index           =   8
      Left            =   7110
      TabIndex        =   56
      Top             =   4230
      Width           =   1095
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "field1"
      Height          =   195
      Index           =   26
      Left            =   120
      TabIndex        =   55
      Top             =   3960
      Width           =   600
   End
   Begin VB.Label http 
      BorderStyle     =   1  '單線固定
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   54
      Top             =   3900
      Width           =   1095
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "API Key"
      Height          =   195
      Index           =   24
      Left            =   180
      TabIndex        =   52
      Top             =   3600
      Width           =   1005
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "Channel ID"
      Height          =   195
      Index           =   23
      Left            =   180
      TabIndex        =   51
      Top             =   3270
      Width           =   1005
   End
   Begin VB.Label digital 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Height          =   255
      Index           =   13
      Left            =   6960
      TabIndex        =   48
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label digital 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Height          =   255
      Index           =   12
      Left            =   6960
      TabIndex        =   47
      Top             =   1140
      Width           =   1095
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "D13"
      Height          =   195
      Index           =   22
      Left            =   6240
      TabIndex        =   46
      Top             =   1500
      Width           =   600
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "D12"
      Height          =   195
      Index           =   21
      Left            =   6240
      TabIndex        =   45
      Top             =   1200
      Width           =   600
   End
   Begin VB.Label digital 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Height          =   255
      Index           =   11
      Left            =   4920
      TabIndex        =   44
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label digital 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Height          =   255
      Index           =   10
      Left            =   4920
      TabIndex        =   43
      Top             =   2340
      Width           =   1095
   End
   Begin VB.Label digital 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Height          =   255
      Index           =   9
      Left            =   4920
      TabIndex        =   42
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label digital 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Height          =   255
      Index           =   8
      Left            =   4920
      TabIndex        =   41
      Top             =   1740
      Width           =   1095
   End
   Begin VB.Label digital 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Height          =   255
      Index           =   7
      Left            =   4920
      TabIndex        =   40
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "D11"
      Height          =   195
      Index           =   20
      Left            =   4200
      TabIndex        =   39
      Top             =   2700
      Width           =   600
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "D10"
      Height          =   195
      Index           =   19
      Left            =   4200
      TabIndex        =   38
      Top             =   2400
      Width           =   600
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "D9"
      Height          =   195
      Index           =   18
      Left            =   4200
      TabIndex        =   37
      Top             =   2100
      Width           =   600
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "D8"
      Height          =   195
      Index           =   17
      Left            =   4200
      TabIndex        =   36
      Top             =   1800
      Width           =   600
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "D7"
      Height          =   195
      Index           =   16
      Left            =   4200
      TabIndex        =   35
      Top             =   1500
      Width           =   600
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "D6"
      Height          =   195
      Index           =   15
      Left            =   4200
      TabIndex        =   34
      Top             =   1200
      Width           =   600
   End
   Begin VB.Label digital 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Height          =   255
      Index           =   6
      Left            =   4920
      TabIndex        =   33
      Top             =   1140
      Width           =   1095
   End
   Begin VB.Label digital 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Height          =   255
      Index           =   5
      Left            =   2880
      TabIndex        =   32
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label digital 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Height          =   255
      Index           =   4
      Left            =   2880
      TabIndex        =   31
      Top             =   2340
      Width           =   1095
   End
   Begin VB.Label digital 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Height          =   255
      Index           =   3
      Left            =   2880
      TabIndex        =   30
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label digital 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Height          =   255
      Index           =   2
      Left            =   2880
      TabIndex        =   29
      Top             =   1740
      Width           =   1095
   End
   Begin VB.Label digital 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   28
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "D5"
      Height          =   195
      Index           =   14
      Left            =   2160
      TabIndex        =   27
      Top             =   2700
      Width           =   600
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "D4"
      Height          =   195
      Index           =   13
      Left            =   2160
      TabIndex        =   26
      Top             =   2400
      Width           =   600
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "D3"
      Height          =   195
      Index           =   12
      Left            =   2160
      TabIndex        =   25
      Top             =   2100
      Width           =   600
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "D2"
      Height          =   195
      Index           =   11
      Left            =   2160
      TabIndex        =   24
      Top             =   1800
      Width           =   600
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "D1"
      Height          =   195
      Index           =   10
      Left            =   2160
      TabIndex        =   23
      Top             =   1500
      Width           =   600
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "D0"
      Height          =   195
      Index           =   9
      Left            =   2160
      TabIndex        =   22
      Top             =   1200
      Width           =   600
   End
   Begin VB.Label digital 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Height          =   255
      Index           =   0
      Left            =   2880
      TabIndex        =   20
      Top             =   1140
      Width           =   1095
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "A5"
      Height          =   195
      Index           =   8
      Left            =   180
      TabIndex        =   18
      Top             =   2700
      Width           =   600
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "A4"
      Height          =   195
      Index           =   7
      Left            =   180
      TabIndex        =   17
      Top             =   2400
      Width           =   600
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "A3"
      Height          =   195
      Index           =   6
      Left            =   180
      TabIndex        =   16
      Top             =   2100
      Width           =   600
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "A2"
      Height          =   195
      Index           =   5
      Left            =   180
      TabIndex        =   15
      Top             =   1800
      Width           =   600
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "A1"
      Height          =   195
      Index           =   4
      Left            =   180
      TabIndex        =   14
      Top             =   1500
      Width           =   600
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "A0"
      Height          =   195
      Index           =   3
      Left            =   180
      TabIndex        =   13
      Top             =   1200
      Width           =   600
   End
   Begin VB.Label analog 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   12
      Top             =   1140
      Width           =   1095
   End
   Begin VB.Label analog 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   11
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label analog 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   10
      Top             =   1740
      Width           =   1095
   End
   Begin VB.Label analog 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   9
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label analog 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Height          =   255
      Index           =   4
      Left            =   840
      TabIndex        =   8
      Top             =   2340
      Width           =   1095
   End
   Begin VB.Label analog 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Height          =   255
      Index           =   5
      Left            =   840
      TabIndex        =   7
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "Com Port"
      Height          =   195
      Index           =   2
      Left            =   180
      TabIndex        =   5
      Top             =   840
      Width           =   800
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "TTS音量"
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   4
      Top             =   480
      Width           =   800
   End
   Begin VB.Label label 
      Alignment       =   1  '靠右對齊
      Caption         =   "TTS速度"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   3
      Top             =   180
      Width           =   800
   End
End
Attribute VB_Name = "frmConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const WorkingPortNo = 50001

Private V As SpeechLib.SpVoice
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal ByteLen As Long)

Dim NumSockets As Long
Dim PLength As Integer

Dim digitalPin(0 To 13) As Boolean
Dim analogPin(0 To 6) As Boolean
Dim thingspeakField(1 To 8) As Boolean
Dim thingspeakRead As Boolean
Dim thingspeakWrite As Boolean
Dim thingspeakWorking As Boolean
Dim thingspeakWorking2 As Boolean
Dim httpReadCounter As Integer

Private Const INFINITE = &HFFFFFFFF
Private Const SYNCHRONIZE = &H100000
Private Const PROCESS_QUERY_INFORMATION = &H400&

Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" ( _
    ByVal hProcess As Long, _
    lpExitCode As Long) As Long

Private Declare Function OpenProcess Lib "kernel32" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" ( _
    ByVal hHandle As Long, _
    ByVal dwMilliseconds As Long) As Long
    
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
(ByVal lpApplicationName As String, _
ByVal lpKeyName As String, _
ByVal lpDefault As String, _
ByVal lpReturnedString As String, _
ByVal nSize As Long, _
ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
(ByVal lpApplicationName As String, _
ByVal lpKeyName As String, _
ByVal lpString As String, _
ByVal lpFileName As String) As Long

Public Function ShellSync( _
    ByVal PathName As String, _
    ByVal WindowStyle As VbAppWinStyle) As Long
    'Shell and wait.  Return exit code result, raise an
    'exception on any error.
    Dim lngPid As Long
    Dim lngHandle As Long
    Dim lngExitCode As Long

    lngPid = Shell(PathName, WindowStyle)
    If lngPid <> 0 Then
        lngHandle = OpenProcess(SYNCHRONIZE _
                             Or PROCESS_QUERY_INFORMATION, 0, lngPid)
        If lngHandle <> 0 Then
            WaitForSingleObject lngHandle, INFINITE
            If GetExitCodeProcess(lngHandle, lngExitCode) <> 0 Then
                ShellSync = lngExitCode
                CloseHandle lngHandle
            Else
                CloseHandle lngHandle
                Err.Raise &H8004AA00, "ShellSync", _
                          "Failed to retrieve exit code, error " _
                        & CStr(Err.LastDllError)
            End If
        Else
            Err.Raise &H8004AA01, "ShellSync", _
                      "Failed to open child process"
        End If
    Else
        Err.Raise &H8004AA02, "ShellSync", _
                  "Failed to Shell child process"
    End If
End Function

Public Function ByteToFloat(ByRef b() As Byte) As Single
   RtlMoveMemory ByteToFloat, b(0), 4
End Function

Private Function iniStr(srcStr$, strLen As Long) As String
    Dim s As String
    Dim pos As Integer
    s = Left$(srcStr$, strLen)
    pos = InStr(s, Chr(0))
    If pos > 1 Then
        iniStr = Left$(s, pos - 1)
    ElseIf pos = 1 Then
        iniStr = ""
    Else
        iniStr = s
    End If
End Function

Private Function iniVal(srcStr$, strLen As Long) As Long
    iniVal = Val(Left$(srcStr$, strLen))
End Function

Private Sub DigitalOutput(ByVal pin As Byte, ByVal value As Byte)
   If MSComm1.PortOpen Then
      MSComm1.Output = Chr(255) + Chr(85) + Chr(5) + Chr(0) + Chr(2) + Chr(30) + Chr(pin) + Chr(value)
   End If
End Sub

Private Sub PWMOutput(ByVal pin As Byte, ByVal value As Byte)
   If MSComm1.PortOpen Then
      MSComm1.Output = Chr(255) + Chr(85) + Chr(5) + Chr(0) + Chr(2) + Chr(32) + Chr(pin) + Chr(value)
   End If
End Sub

Private Sub DigitalInput(ByVal pin As Byte)
   If MSComm1.PortOpen Then
      MSComm1.Output = Chr(255) + Chr(85) + Chr(4) + Chr(pin) + Chr(1) + Chr(30) + Chr(pin)
   End If
End Sub

Private Sub AnalogInput(ByVal pin As Byte)
   If MSComm1.PortOpen Then
      MSComm1.Output = Chr(255) + Chr(85) + Chr(4) + Chr(14 + pin) + Chr(1) + Chr(31) + Chr(pin)
   End If
End Sub

Private Sub PlayNote(ByVal pin As Byte, ByVal freq As Long, ByVal ms As Long)
   If MSComm1.PortOpen Then
      MSComm1.Output = Chr(255) + Chr(85) + Chr(8) + Chr(0) + Chr(2) + Chr(34) + Chr(pin) + Chr(freq Mod 256) + Chr(freq \ 256) + Chr(ms Mod 256) + Chr(ms \ 256)
   End If
End Sub

Private Sub ServoMotor(ByVal pin As Byte, ByVal angle As Byte)
   If MSComm1.PortOpen Then
      MSComm1.Output = Chr(255) + Chr(85) + Chr(5) + Chr(0) + Chr(2) + Chr(33) + Chr(pin) + Chr(angle)
   End If
End Sub

Private Sub cmdEndPro_Click()
   If MSComm1.PortOpen Then
      MSComm1.PortOpen = False
   End If
   sckServer(0).Close
   Unload Me
End Sub

Private Sub cmdArduino_Click()
   If Not MSComm1.PortOpen Then
      With MSComm1
        .CommPort = Val(Combo1.Text)
        .DTREnable = True
        .RTSEnable = True
        .RThreshold = 1
        .SThreshold = 0
        .Settings = "115200,n,8,1"
        .InputMode = comInputModeBinary
        .PortOpen = True
      End With
      'Data = ""
      PLength = 0
   End If
End Sub

Private Sub thingspeakReadData(ByVal ChannelID As String)
   On Error Resume Next
   If thingspeakRead And Not thingspeakWorking Then
      thingspeakWorking = True
      Inet1.URL = "https://api.thingspeak.com/channels/" & ChannelID & "/feeds.json?results=1"
      Inet1.Execute
   End If
End Sub

Private Sub thingspeakWriteData(ByVal APIKey As String)
   On Error Resume Next
   Dim urlext As String
   If thingspeakWrite And Not thingspeakWorking2 Then
      thingspeakWorking2 = True
      urlext = ""
      For i = 1 To 8
          If Len(txtFV(i).Text) > 0 Then
             urlext = urlext & "&field" & Trim(CStr(i)) & "=" & Trim(txtFV(i).Text)
          End If
      Next
      Inet2.URL = "https://api.thingspeak.com/update?api_key=" & Trim(txtAPIKey.Text) & urlext
      Inet2.Execute
   End If
End Sub

Private Sub cmdReadThingSpeak_Click()
   If Len(txtChannelID.Text) > 0 Then
      thingspeakReadData Trim(txtChannelID.Text)
   End If
End Sub

Private Sub cmdWriteThingSpeak_Click()
   If Len(txtAPIKey.Text) > 0 Then
      thingspeakWriteData Trim(txtAPIKey.Text)
   End If
End Sub

Private Sub cmdSetThingSpeak_Click()
   thingspeakField(Val(Combo2.Text)) = True
   thingspeakRead = True
   thingspeakWrite = True
   txtFV(Val(Combo2.Text)).Text = Val(txtFieldVal.Text)
End Sub

Private Sub cmdUploadHEX_Click()
  If MSComm1.PortOpen Then
     MSComm1.PortOpen = False
  End If
  If Val(Combo1.Text) > 0 Then
     DisableText
     ShellSync "arduino\avrdude -Carduino\avrdude.conf -s -patmega328p -carduino -P\\.\COM" & Trim(CStr(Val(Combo1.Text))) & " -b115200 -D -Uflash:w:arduino\uno.hex:i", vbNormalFocus
     EnableText
  End If
End Sub

Private Sub Form_Load()
   Dim ArduinoPortNo As Integer
   
   Combo1.Clear
   For i = 0 To 99
      Combo1.AddItem Format(i, "00")
   Next
   
   Set obj = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
   Set items = obj.ExecQuery("Select * from Win32_PnPEntity WHERE ConfigManagerErrorCode = 0")
  
   ArduinoPortNo = 0
   For Each Item In items
      
      If InStr(Item.Name, "Arduino") > 0 Then
         leftbrace = InStr(Item.Name, "(")
         rightbrace = InStr(Item.Name, ")")
          
         ArduinoPortNo = Val(Mid(Item.Name, leftbrace + 4, rightbrace - leftbrace - 3))
         Exit For
      End If
   Next
   Combo1.ListIndex = ArduinoPortNo
   
   Combo2.Clear
   For i = 1 To 8
      Combo2.AddItem Format(i, "0")
   Next
   Combo2.ListIndex = 0
   
   For i = 0 To 13
       digitalPin(i) = False
   Next
   
   For i = 0 To 6
       analogPin(i) = False
   Next
   
   For i = 1 To 8
       thingspeakField(i) = False
   Next
    
   Set V = New SpeechLib.SpVoice
   V.Rate = HScroll1.value
   V.Volume = HScroll2.value
   
   NumSockets = 0
   sckServer(0).LocalPort = WorkingPortNo
   sckServer(0).Listen
   Text1.Text = ""
   httpReadCounter = 0
   
   thingspeakRead = False
   thingspeakWrite = False
   thingspeakWorking = False
   thingspeakWorking2 = False
   
   Dim rc As Long
   Dim Rtvl As String * 260
   Dim iniFN As String
    
   On Error Resume Next
   iniFN = App.Path & "\Scratch.ini"
   
   If Len(Dir(iniFN)) > 0 Then
      rc = GetPrivateProfileString("Data", "APIKey", "", Rtvl, 256, iniFN)
      txtAPIKey.Text = iniStr(Rtvl, rc)
      rc = GetPrivateProfileString("Data", "ChannelID", "", Rtvl, 256, iniFN)
      txtChannelID.Text = iniStr(Rtvl, rc)
   Else
      WritePrivateProfileString "Data", "APIKey", "", iniFN
      WritePrivateProfileString "Data", "ChannelID", "", iniFN
   End If
End Sub


Private Function ProcessRequestData(pstrRequestData As String) As String
   Dim strRequestedFile As String
   Dim http_response As String
   Dim strArray() As String
   Dim response As String
   Dim extraresponse As String

   strRequestedFile = Mid$(pstrRequestData, InStr(1, pstrRequestData, "GET /", vbTextCompare) + 5, InStr(1, pstrRequestData, " HTTP", vbTextCompare) - 6)
   strArray = Split(strRequestedFile, "/")
    
   'If Not strArray(0) = "poll" Then
      'Debug.Print strRequestedFile
   'End If
    
   response = "NoData"
   Select Case strArray(0)
       Case "ttsNoWait"
         Text1.Text = UTF8_UrlDecode(strArray(1)) & vbCrLf & Text1.Text
         V.Speak UTF8_UrlDecode(strArray(1)), SVSFlagsAsync
       Case "ttsWait"
         Text1.Text = UTF8_UrlDecode(strArray(2)) & vbCrLf & Text1.Text
         V.Speak UTF8_UrlDecode(strArray(2)), 0
       Case "voiceSpeed"
         V.Rate = Int(strArray(1))
         HScroll1.value = Int(strArray(1))
       Case "voiceVolume"
         V.Volume = Int(strArray(1))
         HScroll2.value = Int(strArray(1))
       Case "playNote"
         digital(Val(strArray(1))) = Val(strArray(2)) & "," & Val(strArray(3))
         PlayNote Val(strArray(1)), Val(strArray(2)), Val(strArray(3))
       Case "digitalOutput"
         digital(Val(strArray(1))) = Val(strArray(2))
         DigitalOutput Val(strArray(1)), Val(strArray(2))
       Case "httpOutput"
         txtFV(Val(strArray(1))).Text = Val(strArray(2))
         'thingspeakField(Val(strArray(1))) = True
         'thingspeakRead = True
         thingspeakWrite = True
       Case "httpOutputGo"
         If Len(Trim(txtAPIKey.Text)) > 0 Then
            'Do While thingspeakWorking2
            '   DoEvents
            'Loop
            thingspeakWriteData Trim(txtAPIKey.Text)
         End If
       Case "pwmOutput"
         digital(Val(strArray(1))) = Val(strArray(2))
         PWMOutput Val(strArray(1)), Val(strArray(2))
       Case "servoMotor"
         digital(Val(strArray(1))) = Val(strArray(2)) & "Deg"
         ServoMotor Val(strArray(1)), Val(strArray(2))
       Case "sethttpInput"
         http(Val(strArray(1))).BackColor = vbGreen
         thingspeakField(Val(strArray(1))) = True
         thingspeakRead = True
       Case "setdigitalInput"
         digital(Val(strArray(1))).BackColor = vbGreen
         digitalPin(Val(strArray(1))) = True
       Case "setanalogInput"
         analog(Val(strArray(1))).BackColor = vbGreen
         analogPin(Val(strArray(1))) = True
   End Select
    
   extraresponse = ""
   For i = 2 To 13
       If digitalPin(i) Then
          If Len(digital(i).Caption) > 0 Then
             extraresponse = extraresponse + "digitalInput/" & Trim(CStr(i)) & " " & digital(i).Caption + vbCrLf
          End If
          DigitalInput i
       End If
   Next
   For i = 0 To 5
       If analogPin(i) Then
          If Len(analog(i).Caption) > 0 Then
             extraresponse = extraresponse + "analogInput/" & Trim(CStr(i)) & " " & analog(i).Caption + vbCrLf
          End If
          AnalogInput i
       End If
   Next
    
   httpReadCounter = httpReadCounter + 1
   If Not thingspeakWorking2 Then
      If httpReadCounter >= 4 Then
         For i = 1 To 8
             If thingspeakField(i) Then
                If Len(http(i).Caption) > 0 Then
                   extraresponse = extraresponse + "httpInput/" & Trim(CStr(i)) & " " & http(i).Caption + vbCrLf
                End If
             End If
             If Not thingspeakWorking Then
                If Len(txtChannelID.Text) > 0 Then
                   thingspeakReadData Trim(txtChannelID.Text)
                End If
             End If
         Next
         httpReadCounter = 0
      End If
   End If
   
   If Len(extraresponse) > 0 Then
      response = extraresponse
   End If
    
   If Not response = "NoData" Then
      http_response = "HTTP/1.1 200 OK" + vbCrLf
      http_response = http_response + "Content-Type: text/html; charset=ISO-8859-1" + vbCrLf
      http_response = http_response + "Content-Length" + Str(Len(response)) + vbCrLf
      http_response = http_response + "Access-Control-Allow-Origin: *" + vbCrLf
      http_response = http_response + vbCrLf
      http_response = http_response + response + vbCrLf
   Else
      http_response = "HTTP/1.1 200 OK" + vbCrLf
      http_response = http_response + "Content-Type: text/html; charset=ISO-8859-1" + vbCrLf
      http_response = http_response + "Content-Length" + Str(0) + vbCrLf
      http_response = http_response + "Access-Control-Allow-Origin: *" + vbCrLf
      http_response = http_response + vbCrLf
   End If
    
   ProcessRequestData = http_response
End Function

Public Function UTF8_UrlDecode(ByVal URL As String)
   Dim b, ub
   Dim UtfB
   Dim UtfB1, UtfB2, UtfB3
   Dim i, n, s
   n = 0
   ub = 0
   For i = 1 To Len(URL)
       b = Mid(URL, i, 1)
       Select Case b
           Case "+"
               s = s & " "
           Case "%"
               ub = Mid(URL, i + 1, 2)
               UtfB = CInt("&H" & ub)
               If UtfB < 128 Then
                   i = i + 2
                   s = s & ChrW(UtfB)
               Else
                   UtfB1 = (UtfB And &HF) * &H1000
                   UtfB2 = (CInt("&H" & Mid(URL, i + 4, 2)) And &H3F) * &H40
                   UtfB3 = CInt("&H" & Mid(URL, i + 7, 2)) And &H3F
                   s = s & ChrW(UtfB1 Or UtfB2 Or UtfB3)
                   i = i + 8
               End If
           Case Else
               s = s & b
       End Select
   Next
   UTF8_UrlDecode = s
End Function

Private Sub Inet1_StateChanged(ByVal State As Integer)
   Dim substrTemp As String
   Dim strTemp As String
   Dim pos As Integer
   Dim fieldsData() As String
   Dim fieldData() As String
    
   Select Case State
    Case icConnecting
    Case icConnected
    Case icRequesting
    Case icRequestSent
    Case icReceivingResponse
    Case icResponseReceived
    
    Case icResponseCompleted
        strTemp = Inet1.GetChunk(1024)
        pos = InStr(strTemp, "field1")
        If pos > 0 Then
           Do While True
              pos = InStr(strTemp, "field1")
              If pos > 0 Then
                 strTemp = Mid(strTemp, pos + 6, Len(strTemp) - pos - 5)
              Else
                 Exit Do
              End If
              DoEvents
           Loop
        
           strTemp = """field1" & Left(strTemp, Len(strTemp) - 3)
           strTemp = Replace(strTemp, "field", "")
           strTemp = Replace(strTemp, """", "")
        
           fieldsData = Split(strTemp, ",")
           For i = 0 To UBound(fieldsData)
               fieldData = Split(fieldsData(i), ":")
               If UBound(fieldData) = 1 Then
                  http(Val(fieldData(0))).Caption = fieldData(1)
               End If
           Next
       End If
       thingspeakWorking = False
   End Select
End Sub

Private Sub Inet2_StateChanged(ByVal State As Integer)
   Dim strTemp As String
    
   Select Case State
    Case icConnecting
    Case icConnected
    Case icRequesting
    Case icRequestSent
    Case icReceivingResponse
    Case icResponseReceived
    
    Case icResponseCompleted
        thingspeakWorking2 = False
        strTemp = Inet1.GetChunk(1024)

   End Select
End Sub

Private Sub MSComm1_OnComm()
   Dim sBuffer() As Byte
   Dim databyte(0 To 255) As Byte
   Dim fval(0 To 3) As Byte

   Select Case MSComm1.CommEvent
      ' Handle each event or error by placing
      ' code below each case statement.
      ' This template is found in the Example
      ' section of the OnComm event Help topic
      ' in VB Help.

      ' Errors
       Case comEventBreak   ' A Break was received.
       Case comEventCDTO    ' CD (RLSD) Timeout.
       Case comEventCTSTO   ' CTS Timeout.
       Case comEventDSRTO   ' DSR Timeout.
       Case comEventFrame   ' Framing Error.
       Case comEventOverrun ' Data Lost.
       Case comEventRxOver  ' Receive buffer overflow.
       Case comEventRxParity   ' Parity Error.
       Case comEventTxFull  ' Transmit buffer full.
       Case comEventDCB     ' Unexpected error retrieving DCB]

      ' Events
       Case comEvCD   ' Change in the CD line.
       Case comEvCTS  ' Change in the CTS line.
       Case comEvDSR  ' Change in the DSR line.
       Case comEvRing ' Change in the Ring Indicator.
       Case comEvReceive ' Received RThreshold # of chars.

         sBuffer = MSComm1.Input
         For i = 0 To UBound(sBuffer)
             If sBuffer(i) = 255 Then
                If i + 1 <= UBound(sBuffer) Then
                   If sBuffer(i + 1) = 85 Then
                      PLength = 0
                   End If
                End If
             End If
             databyte(PLength) = sBuffer(i)
             If PLength >= 3 Then
                If databyte(PLength) = 10 And databyte(PLength - 1) = 13 Then
                   If databyte(3) = 2 Then
                      fval(0) = databyte(4)
                      fval(1) = databyte(5)
                      fval(2) = databyte(6)
                      fval(3) = databyte(7)
                      If databyte(2) >= 14 Then
                         analog(databyte(2) - 14).Caption = ByteToFloat(fval)
                      ElseIf databyte(2) >= 2 Then
                         digital(databyte(2)).Caption = ByteToFloat(fval)
                      End If
                   End If
                End If
             End If
             PLength = PLength + 1
             DoEvents
         Next

      Case comEvSend ' There are SThreshold number of
                     ' characters in the transmit buffer.
      Case comEvEOF  ' An EOF character was found in the
                     ' input stream.
   End Select
End Sub

Function Dec2Hex(ByVal dec As Byte)
   If Len(Hex(dec)) < 2 Then
      Dec2Hex = "0" & Hex(dec)
   Else
      Dec2Hex = Hex(dec)
   End If
End Function

Private Sub sckServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
   '// Accept the incomming connection from a browser.
   NumSockets = NumSockets + 1
   If NumSockets > 32767 Then
      NumSockets = 1
   End If
    
   '//Increase Number of Sockets by one.
   Load sckServer(NumSockets)
   With sckServer(NumSockets)
      .Accept (requestID)
   End With
End Sub

Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
   Dim strRequestData As String
    
   '// Retreive the incomming data and return the requested file.
   sckServer(Index).GetData strRequestData
   sckServer(Index).SendData ProcessRequestData(strRequestData)
End Sub

Private Sub sckServer_SendComplete(Index As Integer)
   '// Reset the connection state after it as returned the requested file.
   With sckServer(Index)
       .Close
   End With
   Unload sckServer(Index)
End Sub

Private Sub DisableText()
   cmdArduino.Enabled = False
   cmdUploadHEX.Enabled = False
   cmdEndPro.Enabled = False
   cmdReadThingSpeak.Enabled = False
   cmdWriteThingSpeak.Enabled = False
   cmdSetThingSpeak.Enabled = False
   HScroll1.Enabled = False
   HScroll2.Enabled = False
   Combo1.Enabled = False
   Combo2.Enabled = False
   txtFieldVal.Enabled = False
   'txtChannelID.Enabled = False
   'txtAPIKey.Enabled = False
End Sub

Private Sub EnableText()
   cmdArduino.Enabled = True
   cmdUploadHEX.Enabled = True
   cmdEndPro.Enabled = True
   cmdReadThingSpeak.Enabled = True
   cmdWriteThingSpeak.Enabled = True
   cmdSetThingSpeak.Enabled = True
   HScroll1.Enabled = True
   HScroll2.Enabled = True
   Combo1.Enabled = True
   Combo2.Enabled = True
   txtFieldVal.Enabled = True
   'txtChannelID.Enabled = True
   'txtAPIKey.Enabled = True
End Sub
