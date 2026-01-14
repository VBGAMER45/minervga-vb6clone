VERSION 5.00
Begin VB.Form frmStore
   BackColor       =   &H00004080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "General Store"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit
      Caption         =   "&Leave Store"
      Height          =   375
      Left            =   4320
      TabIndex        =   12
      Top             =   5220
      Width           =   1575
   End
   Begin VB.CommandButton cmdElevator
      Caption         =   "&Upgrade Elevator ($500)"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   5220
      Width           =   2295
   End
   Begin VB.ListBox lstItems
      BeginProperty Font
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2340
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   5775
   End
   Begin VB.CommandButton cmdBuy
      Caption         =   "&Buy Selected Item"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Frame fraStatus
      BackColor       =   &H00004080&
      Caption         =   "Your Status"
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin VB.Label lblCash
         BackColor       =   &H00004080&
         Caption         =   "Cash: $0"
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
         Width           =   5535
      End
      Begin VB.Label lblInventory
         BackColor       =   &H00004080&
         Caption         =   "Inventory: None"
         BeginProperty Font
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   5535
      End
   End
   Begin VB.Label lblElevator
      BackColor       =   &H00004080&
      Caption         =   "Elevator Depth: 0 ft (Max: 800 ft)"
      BeginProperty Font
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   5280
      Width           =   3375
   End
   Begin VB.Label lblDescription
      BackColor       =   &H00004080&
      Caption         =   "Select an item to see its description"
      BeginProperty Font
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   4080
      Width           =   5775
   End
   Begin VB.Label lblItems
      BackColor       =   &H00004080&
      Caption         =   "Items for Sale:"
      BeginProperty Font
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   5775
   End
End
Attribute VB_Name = "frmStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================================
' MinerVGA - General Store Form
' ============================================================================

' Item data for listbox
Private Type StoreItem
    ItemID As Integer
    Name As String
    Cost As Long
    Description As String
    Available As Boolean
End Type

Private StoreItems(1 To 9) As StoreItem

Private Sub Form_Load()
    ' Initialize store items
    Call InitStoreItems
    Call UpdateDisplay
End Sub

Private Sub InitStoreItems()
    StoreItems(1).ItemID = ITEM_SHOVEL
    StoreItems(1).Name = "Shovel"
    StoreItems(1).Cost = COST_SHOVEL
    StoreItems(1).Description = "Reduces digging cost by $3 per cell"
    StoreItems(1).Available = True

    StoreItems(2).ItemID = ITEM_PICKAXE
    StoreItems(2).Name = "Pickaxe"
    StoreItems(2).Cost = COST_PICKAXE
    StoreItems(2).Description = "Reduces digging cost by $4 per cell"
    StoreItems(2).Available = True

    StoreItems(3).ItemID = ITEM_DRILL
    StoreItems(3).Name = "Drill"
    StoreItems(3).Cost = COST_DRILL
    StoreItems(3).Description = "Required to drill through granite (press D)"
    StoreItems(3).Available = True

    StoreItems(4).ItemID = ITEM_LANTERN
    StoreItems(4).Name = "Lantern"
    StoreItems(4).Cost = COST_LANTERN
    StoreItems(4).Description = "Light source - helps spot minerals/hazards"
    StoreItems(4).Available = True

    StoreItems(5).ItemID = ITEM_BUCKET
    StoreItems(5).Name = "Bucket"
    StoreItems(5).Cost = COST_BUCKET
    StoreItems(5).Description = "Required to pump out water (press P)"
    StoreItems(5).Available = True

    StoreItems(6).ItemID = ITEM_TORCH
    StoreItems(6).Name = "Torch"
    StoreItems(6).Cost = COST_TORCH
    StoreItems(6).Description = "Light source, also needed to light dynamite"
    StoreItems(6).Available = True

    StoreItems(7).ItemID = ITEM_DYNAMITE
    StoreItems(7).Name = "Dynamite"
    StoreItems(7).Cost = COST_DYNAMITE
    StoreItems(7).Description = "Blasts 3x3 area (needs torch, escape right)"
    StoreItems(7).Available = True

    StoreItems(8).ItemID = ITEM_RING
    StoreItems(8).Name = "Diamond Ring"
    StoreItems(8).Cost = COST_RING
    StoreItems(8).Description = "Required to win Miss Mimi's heart!"
    StoreItems(8).Available = True

    StoreItems(9).ItemID = ITEM_CONDOM
    StoreItems(9).Name = "Condom"
    StoreItems(9).Cost = COST_CONDOM
    StoreItems(9).Description = "For protection when visiting Miss Mimi..."
    StoreItems(9).Available = True
End Sub

Private Sub UpdateDisplay()
    Dim i As Integer
    Dim ItemText As String
    Dim OwnedMark As String

    ' Update cash
    lblCash.Caption = "Cash: $" & Format(Player.Cash, "#,##0")

    ' Update inventory
    lblInventory.Caption = "Inventory: " & GetInventoryString()

    ' Update elevator info
    Dim CurrentDepth As Integer, MaxDepth As Integer
    CurrentDepth = (MaxElevatorDepth - 4) * 20
    MaxDepth = (MAX_MINE_DEPTH - 4) * 20
    lblElevator.Caption = "Elevator: " & CurrentDepth & " ft (Max: " & MaxDepth & " ft)"

    ' Enable elevator upgrade if not at max
    cmdElevator.Enabled = (MaxElevatorDepth < MAX_MINE_DEPTH) And (Player.Cash >= COST_ELEVATOR_UPGRADE)

    ' Populate item list
    lstItems.Clear
    For i = 1 To 9
        If HasItem(StoreItems(i).ItemID) Then
            OwnedMark = " [OWNED]"
        Else
            OwnedMark = ""
        End If

        ItemText = StoreItems(i).Name & " - $" & StoreItems(i).Cost & OwnedMark
        lstItems.AddItem ItemText
    Next i
End Sub

Private Sub lstItems_Click()
    Dim Index As Integer
    Index = lstItems.ListIndex + 1

    If Index >= 1 And Index <= 9 Then
        lblDescription.Caption = StoreItems(Index).Description

        ' Enable/disable buy button
        If HasItem(StoreItems(Index).ItemID) Then
            cmdBuy.Enabled = False
            lblDescription.Caption = lblDescription.Caption & vbCrLf & "(You already own this item)"
        ElseIf Player.Cash < StoreItems(Index).Cost Then
            cmdBuy.Enabled = False
            lblDescription.Caption = lblDescription.Caption & vbCrLf & "(Not enough money!)"
        Else
            cmdBuy.Enabled = True
        End If
    End If
End Sub

Private Sub cmdBuy_Click()
    Dim Index As Integer
    Index = lstItems.ListIndex + 1

    If Index >= 1 And Index <= 9 Then
        If BuyItem(StoreItems(Index).ItemID) Then
            Call UpdateDisplay
            ' Re-select to update description
            lstItems.ListIndex = Index - 1
        End If
    End If
End Sub

Private Sub cmdElevator_Click()
    If Player.Cash >= COST_ELEVATOR_UPGRADE Then
        If MaxElevatorDepth < MAX_MINE_DEPTH Then
            Player.Cash = Player.Cash - COST_ELEVATOR_UPGRADE
            Call UpgradeElevator

            Dim NewDepth As Integer
            NewDepth = (MaxElevatorDepth - 4) * 20
            MsgBox "Elevator upgraded! Now reaches " & NewDepth & " feet.", vbInformation, "Store"

            Call UpdateDisplay
        Else
            MsgBox "Elevator is already at maximum depth!", vbExclamation, "Store"
        End If
    Else
        MsgBox "You need $" & COST_ELEVATOR_UPGRADE & " to upgrade the elevator!", vbExclamation, "Store"
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub
