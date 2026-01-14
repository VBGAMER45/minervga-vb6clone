VERSION 5.00
Begin VB.Form frmBank
   BackColor       =   &H00404000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "First National Bank & Assay Office"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit
      Caption         =   "&Leave Bank"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   4020
      Width           =   1575
   End
   Begin VB.CommandButton cmdSell
      Caption         =   "&Sell All Minerals"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4020
      Width           =   1575
   End
   Begin VB.Frame fraAccount
      BackColor       =   &H00404000&
      Caption         =   "Account Status"
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.Label lblCash
         BackColor       =   &H00404000&
         Caption         =   "Cash Balance: $0"
         BeginProperty Font
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label lblMinerals
         BackColor       =   &H00404000&
         Caption         =   "Minerals on hand:"
         BeginProperty Font
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   4935
      End
   End
   Begin VB.Frame fraPrices
      BackColor       =   &H00404000&
      Caption         =   "Current Mineral Prices"
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   5175
      Begin VB.Label lblPrices
         BackColor       =   &H00404000&
         Caption         =   "Silver (Ag): $16 per unit"
         BeginProperty Font
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   735
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   4935
      End
   End
   Begin VB.Label lblTotal
      BackColor       =   &H00404000&
      Caption         =   "Total value of minerals: $0"
      BeginProperty Font
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   5175
   End
End
Attribute VB_Name = "frmBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================================
' MinerVGA - Bank Form
' ============================================================================

Private Sub Form_Load()
    Call UpdateDisplay
End Sub

Private Sub UpdateDisplay()
    ' Update cash display
    lblCash.Caption = "Cash Balance: $" & Format(Player.Cash, "#,##0")

    ' Update minerals on hand
    Dim MineralText As String
    MineralText = "Silver (Ag): " & Player.Silver & " units" & vbCrLf
    MineralText = MineralText & "Gold (Au): " & Player.Gold & " units" & vbCrLf
    MineralText = MineralText & "Platinum (Pt): " & Player.Platinum & " units"
    lblMinerals.Caption = MineralText

    ' Update prices
    Dim PriceText As String
    PriceText = "Silver (Ag): $" & SILVER_VALUE & " per unit" & vbCrLf
    PriceText = PriceText & "Gold (Au): $" & GOLD_VALUE & " per unit" & vbCrLf
    PriceText = PriceText & "Platinum (Pt): $" & PLATINUM_VALUE & " per unit"
    lblPrices.Caption = PriceText

    ' Update total value
    Dim TotalValue As Long
    TotalValue = GetMineralValue()
    lblTotal.Caption = "Total value of minerals: $" & Format(TotalValue, "#,##0")

    ' Enable/disable sell button
    cmdSell.Enabled = (TotalValue > 0)
End Sub

Private Sub cmdSell_Click()
    Dim TotalValue As Long
    Dim SilverCount As Integer, GoldCount As Integer, PlatCount As Integer

    ' Store counts for message
    SilverCount = Player.Silver
    GoldCount = Player.Gold
    PlatCount = Player.Platinum

    ' Sell minerals
    TotalValue = SellMinerals()

    If TotalValue > 0 Then
        Dim Msg As String
        Msg = "Sold minerals for $" & Format(TotalValue, "#,##0") & "!" & vbCrLf & vbCrLf

        If SilverCount > 0 Then
            Msg = Msg & "Silver: " & SilverCount & " x $" & SILVER_VALUE & " = $" & (SilverCount * SILVER_VALUE) & vbCrLf
        End If
        If GoldCount > 0 Then
            Msg = Msg & "Gold: " & GoldCount & " x $" & GOLD_VALUE & " = $" & (GoldCount * GOLD_VALUE) & vbCrLf
        End If
        If PlatCount > 0 Then
            Msg = Msg & "Platinum: " & PlatCount & " x $" & PLATINUM_VALUE & " = $" & (PlatCount * PLATINUM_VALUE) & vbCrLf
        End If

        MsgBox Msg, vbInformation, "Assay Office"
    End If

    Call UpdateDisplay
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub
