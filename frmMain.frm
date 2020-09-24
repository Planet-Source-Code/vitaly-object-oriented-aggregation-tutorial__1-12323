VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Tell me 'bout clothe"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmDescription 
      Caption         =   "Description"
      Height          =   1380
      Left            =   315
      TabIndex        =   6
      Top             =   1575
      Width           =   5370
      Begin VB.Label lblDescription 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   210
         TabIndex        =   7
         Top             =   315
         Width           =   5055
      End
   End
   Begin VB.CommandButton cmdClothe 
      Caption         =   "T-Shirt"
      Height          =   435
      Index           =   5
      Left            =   4095
      TabIndex        =   5
      Top             =   840
      Width           =   1590
   End
   Begin VB.CommandButton cmdClothe 
      Caption         =   "Sweater"
      Height          =   435
      Index           =   4
      Left            =   2205
      TabIndex        =   4
      Top             =   840
      Width           =   1590
   End
   Begin VB.CommandButton cmdClothe 
      Caption         =   "Coat"
      Height          =   435
      Index           =   3
      Left            =   315
      TabIndex        =   3
      Top             =   840
      Width           =   1590
   End
   Begin VB.CommandButton cmdClothe 
      Caption         =   "Gloves"
      Height          =   435
      Index           =   2
      Left            =   4095
      TabIndex        =   2
      Top             =   315
      Width           =   1590
   End
   Begin VB.CommandButton cmdClothe 
      Caption         =   "Jeans"
      Height          =   435
      Index           =   1
      Left            =   2205
      TabIndex        =   1
      Top             =   315
      Width           =   1590
   End
   Begin VB.CommandButton cmdClothe 
      Caption         =   "Hat"
      Height          =   435
      Index           =   0
      Left            =   315
      TabIndex        =   0
      Top             =   315
      Width           =   1590
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'\B/-------------------------------General explanation-------------------------------
'This is an example to Object Oriented inheritance in Visual Basic.
'The example uses classes to emulate getting info from the selected item
'and displaying it.
'
'Comments in this example (like this one) are made by the Commentor add-in get it here -
'http://planet-source-code.com/xq/ASP/txtCodeId.11854/lngWId.1/qx/vb/scripts/ShowCode.htm

'Vitaly Belman
'/E\-------------------------------General explanation-------------------------------
Option Explicit
Dim Cloth As New clsCloth

Private Sub cmdClothe_Click(Index As Integer)
Dim ChoosenCloth As Object 'The slection of the user is stored to this object var...
    Select Case cmdClothe(Index).Caption
    Case "Hat"
        Set ChoosenCloth = Cloth.Hat '...like this. (Look on clsCloth class for more info)
    Case "Jeans"
        Set ChoosenCloth = Cloth.Jeans
    Case "Gloves"
        Set ChoosenCloth = Cloth.Gloves
    Case "Coat"
        Set ChoosenCloth = Cloth.Coat
    Case "Sweater"
        Set ChoosenCloth = Cloth.Sweater
    Case "T-Shirt"
        Set ChoosenCloth = Cloth.TShirt
    End Select
    Call GetObjectDescription(ChoosenCloth) 'Now we're using one single Sub call to print the info.
End Sub
Sub GetObjectDescription(ChoosenCloth As Object)
    lblDescription.Caption = _
      "Color: " & ChoosenCloth.Color & vbNewLine & _
      "Manufacturer: " & ChoosenCloth.Manufacturer & vbNewLine & _
      "First time displayed: " & ChoosenCloth.IsFirstTimeViewed 'The description of the item appears!
    
    ChoosenCloth.IsFirstTimeViewed = False 'This changes the var that says if it is the
                                           'first time you see the information about THAT item.
    'Notice that the changed we do in ChoosenCloth object affects the Cloth object as well.
    'Why? Because we Setted it like this: Set ChoosenCloth = Cloth.XXX, the ChoosenCloth
    'is connected byRef to Cloth object.
End Sub

Private Sub Form_Load()
'\B/---------------------This loads the data about the clothes.---------------------
    With Cloth
        .Coat.Color = "Brown"
        .Coat.Manufacturer = "Taiwan"
            
        .Hat.Color = "Yellow"
        .Hat.Manufacturer = "Loser(TM)"
    
        .Sweater.Color = "Red and white"
        .Sweater.Manufacturer = "Whom-do-you-want-to-get-naked-today?"
        
        .Gloves.Color = "Silver"
        .Gloves.Manufacturer = "Knights&&Dragons"
    
        .Jeans.Color = "Blue"
        .Jeans.Manufacturer = "Gnutella"
    
        .TShirt.Color = "Green"
        .TShirt.Manufacturer = "CoolGB"
    End With
'/E\---------------------This loads the data about the clothes.---------------------
End Sub
