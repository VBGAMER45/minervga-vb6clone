VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MinerVGA Help"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   5520
      Width           =   1335
   End
   Begin VB.TextBox txtHelp 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   5295
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   6975
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================================
' MinerVGA VB6 Edition by vbgamer45
' https://github.com/VBGAMER45/minervga-vb6clone
' https://www.theprogrammingzone.com/
' ============================================================================
Option Explicit

' ============================================================================
' MinerVGA - Help Form
' ============================================================================

Private Sub Form_Load()
    Dim HelpText As String

    HelpText = "==================== MINER VGA ====================" & vbCrLf
    HelpText = HelpText & vbCrLf
    HelpText = HelpText & "OBJECTIVE:" & vbCrLf
    HelpText = HelpText & "  Collect $20,000 and a Diamond Ring to win" & vbCrLf
    HelpText = HelpText & "  Miss Mimi's hand in marriage!" & vbCrLf
    HelpText = HelpText & vbCrLf
    HelpText = HelpText & "================================================" & vbCrLf
    HelpText = HelpText & "CONTROLS:" & vbCrLf
    HelpText = HelpText & "================================================" & vbCrLf
    HelpText = HelpText & vbCrLf
    HelpText = HelpText & "  Arrow Keys  - Move your miner" & vbCrLf
    HelpText = HelpText & "  H           - Show this help screen" & vbCrLf
    HelpText = HelpText & "  X           - Exit game" & vbCrLf
    HelpText = HelpText & "  S           - Save game" & vbCrLf
    HelpText = HelpText & "  R           - Restore saved game" & vbCrLf
    HelpText = HelpText & "  Q           - Toggle sound on/off" & vbCrLf
    HelpText = HelpText & vbCrLf
    HelpText = HelpText & "  E           - Enter building (at doors)" & vbCrLf
    HelpText = HelpText & "  T           - Elevator to Top" & vbCrLf
    HelpText = HelpText & "  B           - Elevator to Bottom" & vbCrLf
    HelpText = HelpText & vbCrLf
    HelpText = HelpText & "  D           - Drill granite (needs Drill)" & vbCrLf
    HelpText = HelpText & "  P           - Pump water (needs Bucket)" & vbCrLf
    HelpText = HelpText & "  Y           - Use dYnamite (needs Torch)" & vbCrLf
    HelpText = HelpText & vbCrLf
    HelpText = HelpText & "================================================" & vbCrLf
    HelpText = HelpText & "BUILDINGS:" & vbCrLf
    HelpText = HelpText & "================================================" & vbCrLf
    HelpText = HelpText & vbCrLf
    HelpText = HelpText & "  BANK      - Sell minerals for cash" & vbCrLf
    HelpText = HelpText & "  STORE     - Buy equipment & upgrade elevator" & vbCrLf
    HelpText = HelpText & "  HOSPITAL  - Heal injuries (costs money)" & vbCrLf
    HelpText = HelpText & "  SALOON    - Visit Miss Mimi to win!" & vbCrLf
    HelpText = HelpText & vbCrLf
    HelpText = HelpText & "================================================" & vbCrLf
    HelpText = HelpText & "MINERALS:" & vbCrLf
    HelpText = HelpText & "================================================" & vbCrLf
    HelpText = HelpText & vbCrLf
    HelpText = HelpText & "  Silver (Ag)   - $16 per unit" & vbCrLf
    HelpText = HelpText & "  Gold (Au)     - $60 per unit" & vbCrLf
    HelpText = HelpText & "  Platinum (Pt) - $2,000 per unit" & vbCrLf
    HelpText = HelpText & vbCrLf
    HelpText = HelpText & "================================================" & vbCrLf
    HelpText = HelpText & "HAZARDS:" & vbCrLf
    HelpText = HelpText & "================================================" & vbCrLf
    HelpText = HelpText & vbCrLf
    HelpText = HelpText & "  Granite   - Blocks path (use Drill)" & vbCrLf
    HelpText = HelpText & "  Water     - Causes damage (use Bucket)" & vbCrLf
    HelpText = HelpText & "  Whirlpool - Damages and floods area" & vbCrLf
    HelpText = HelpText & "  Cave-in   - Damages and refills tunnels" & vbCrLf
    HelpText = HelpText & vbCrLf
    HelpText = HelpText & "================================================" & vbCrLf
    HelpText = HelpText & "TIPS:" & vbCrLf
    HelpText = HelpText & "================================================" & vbCrLf
    HelpText = HelpText & vbCrLf
    HelpText = HelpText & "  - Deeper = richer minerals (more platinum)" & vbCrLf
    HelpText = HelpText & "  - Buy a Lantern to spot hazards early" & vbCrLf
    HelpText = HelpText & "  - Upgrade elevator to reach deeper mines" & vbCrLf
    HelpText = HelpText & "  - Save often! Mining is dangerous!" & vbCrLf
    HelpText = HelpText & "  - Visit the Saloon for hints" & vbCrLf
    HelpText = HelpText & vbCrLf
    HelpText = HelpText & "================================================" & vbCrLf
    HelpText = HelpText & "Good luck, Miner!" & vbCrLf
    HelpText = HelpText & "================================================" & vbCrLf

    txtHelp.Text = HelpText
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Or KeyCode = vbKeyH Then
        Unload Me
    End If
End Sub
