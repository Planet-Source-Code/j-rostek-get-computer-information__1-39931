VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Current Windows"
   ClientHeight    =   11925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   ScaleHeight     =   11925
   ScaleWidth      =   10875
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   6720
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   4560
      Width           =   3735
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   6720
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3840
      Width           =   3735
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   6720
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   3120
      Width           =   3735
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   6720
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2400
      Width           =   3735
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   6720
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1680
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6720
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   960
      Width           =   3735
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10980
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6495
   End
   Begin VB.Label Label7 
      Caption         =   "Win\System Dir"
      Height          =   255
      Left            =   6720
      TabIndex        =   13
      Top             =   4320
      Width           =   3615
   End
   Begin VB.Label Label6 
      Caption         =   "Win Dir"
      Height          =   255
      Left            =   6720
      TabIndex        =   11
      Top             =   3600
      Width           =   3615
   End
   Begin VB.Label Label5 
      Caption         =   "System Drive"
      Height          =   255
      Left            =   6720
      TabIndex        =   10
      Top             =   2880
      Width           =   3615
   End
   Begin VB.Label Label4 
      Caption         =   "Username"
      Height          =   255
      Left            =   6720
      TabIndex        =   9
      Top             =   2160
      Width           =   3615
   End
   Begin VB.Label Label3 
      Caption         =   "OS"
      Height          =   255
      Left            =   6720
      TabIndex        =   8
      Top             =   1440
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Windows Temp Dir"
      Height          =   255
      Left            =   6720
      TabIndex        =   2
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Country specifically Computer"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Information As New GetInfomation
Private Sub Form_Load()
Set Information = New GetInfomation
List1.AddItem "Country                   = " & Information.GetCountry
List1.AddItem "Currency                  = " & Information.GetCurrencySymbol
List1.AddItem "Language                  = " & Information.GetLanguage
List1.AddItem "Date Separator            = " & Information.GetDateSeparator
List1.AddItem "Decimal Separator         = " & Information.GetDecimalSeparator
List1.AddItem "Digit Grouping            = " & Information.GetDigitGrouping
List1.AddItem "Leading Zeros For Decimal = " & Information.GetLeadingZerosForDecimal
List1.AddItem "Long Date Format          = " & Information.GetLongDateFormat
List1.AddItem "Long Month 1              = " & Information.GetLongMonthName1
List1.AddItem "Long Month 2              = " & Information.GetLongMonthName2
List1.AddItem "Long Month 3              = " & Information.GetLongMonthName3
List1.AddItem "Long Month 4              = " & Information.GetLongMonthName4
List1.AddItem "Long Month 5              = " & Information.GetLongMonthName5
List1.AddItem "Long Month 6              = " & Information.GetLongMonthName6
List1.AddItem "Long Month 7              = " & Information.GetLongMonthName7
List1.AddItem "Long Month 8              = " & Information.GetLongMonthName8
List1.AddItem "Long Month 9              = " & Information.GetLongMonthName9
List1.AddItem "Long Month 10             = " & Information.GetLongMonthName10
List1.AddItem "Long Month 11             = " & Information.GetLongMonthName11
List1.AddItem "Long Month 12             = " & Information.GetLongMonthName12
List1.AddItem "Long Day 1                = " & Information.GetLongNameDay1
List1.AddItem "Long Day 2                = " & Information.GetLongNameDay2
List1.AddItem "Long Day 3                = " & Information.GetLongNameDay3
List1.AddItem "Long Day 4                = " & Information.GetLongNameDay4
List1.AddItem "Long Day 5                = " & Information.GetLongNameDay5
List1.AddItem "Long Day 6                = " & Information.GetLongNameDay6
List1.AddItem "Long Day 7                = " & Information.GetLongNameDay7
List1.AddItem "Negative Sign             = " & Information.GetNegativeSign
List1.AddItem "Negative Sign Position    = " & Information.GetNegativeSignPosition
List1.AddItem "Number Fractional Digits  = " & Information.GetNumberOfFractionalDigits
List1.AddItem "Positive Sign             = " & Information.GetPositiveSign
List1.AddItem "Positive Sign Position    = " & Information.GetPositiveSignPosition
List1.AddItem "Short Date Format         = " & Information.GetShortDateFormat
List1.AddItem "Short Month 1             = " & Information.GetShortMonthName1
List1.AddItem "Short Month 2             = " & Information.GetShortMonthName2
List1.AddItem "Short Month 3             = " & Information.GetShortMonthName3
List1.AddItem "Short Month 4             = " & Information.GetShortMonthName4
List1.AddItem "Short Month 5             = " & Information.GetShortMonthName5
List1.AddItem "Short Month 6             = " & Information.GetShortMonthName6
List1.AddItem "Short Month 7             = " & Information.GetShortMonthName7
List1.AddItem "Short Month 8             = " & Information.GetShortMonthName8
List1.AddItem "Short Month 9             = " & Information.GetShortMonthName9
List1.AddItem "Short Month 10            = " & Information.GetShortMonthName10
List1.AddItem "Short Month 11            = " & Information.GetShortMonthName11
List1.AddItem "Short Month 12            = " & Information.GetShortMonthName12
List1.AddItem "Short Day 1               = " & Information.GetShortNameDay1
List1.AddItem "Short Day 2               = " & Information.GetShortNameDay2
List1.AddItem "Short Day 3               = " & Information.GetShortNameDay3
List1.AddItem "Short Day 4               = " & Information.GetShortNameDay4
List1.AddItem "Short Day 5               = " & Information.GetShortNameDay5
List1.AddItem "Short Day 6               = " & Information.GetShortNameDay6
List1.AddItem "Short Day 7               = " & Information.GetShortNameDay7
List1.AddItem "Thousand Separator        = " & Information.GetThousandSeparator
List1.AddItem "Time Format               = " & Information.GetTimeFormat
List1.AddItem "Time Separator            = " & Information.GetTimeSeparator
Text1.Text = Information.TempDir

Text2.Text = Information.GetOS
Text3.Text = Information.GetUSERNAME
Text4.Text = Information.GetSystemDrive
Text5.Text = Information.GetWinDir
Text6.Text = Information.SystemDir
End Sub
