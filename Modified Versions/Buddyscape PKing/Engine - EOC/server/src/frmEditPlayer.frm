VERSION 5.00
Begin VB.Form frmPlayerEditor 
   Caption         =   "EO 2.0 Account Editor by Lightning & Updated by Rithy58"
   ClientHeight    =   10125
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   ScaleHeight     =   10125
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstAccountNames 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   1200
      TabIndex        =   75
      Top             =   9240
      Width           =   2055
   End
   Begin VB.Frame fraVitals 
      Caption         =   "Vitals"
      Height          =   1215
      Left            =   360
      TabIndex        =   70
      Top             =   4920
      Width           =   3615
      Begin VB.TextBox txtMP 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   72
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtHP 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   71
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label32 
         Caption         =   "Mana:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label31 
         Caption         =   "Health:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame fraSpells 
      Caption         =   "Spells"
      Height          =   2295
      Left            =   4920
      TabIndex        =   65
      Top             =   6720
      Width           =   4695
      Begin VB.ListBox lstSpellList 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1620
         Left            =   2400
         TabIndex        =   67
         Top             =   480
         Width           =   2055
      End
      Begin VB.ListBox lstPlayerSpells 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1620
         Left            =   240
         TabIndex        =   66
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label28 
         Caption         =   "Local Spell List:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   69
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label27 
         Caption         =   "Player Spells:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   68
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame fraEquipment 
      Caption         =   "Equipment"
      Height          =   1935
      Left            =   5040
      TabIndex        =   50
      Top             =   4680
      Width           =   4095
      Begin VB.CommandButton cmdShieldUnequip 
         Caption         =   "Unequip"
         Height          =   255
         Left            =   3000
         TabIndex        =   62
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton cmdLegsUnequip 
         Caption         =   "Unequip"
         Height          =   255
         Left            =   3000
         TabIndex        =   61
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdArmourUnequip 
         Caption         =   "Unequip"
         Height          =   255
         Left            =   3000
         TabIndex        =   60
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton cmdWeaponUnequip 
         Caption         =   "Unequip"
         Height          =   255
         Left            =   3000
         TabIndex        =   59
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtHelmet 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   56
         Text            =   "--Empty--"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtShield 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   55
         Text            =   "--Empty--"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtWeapon 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   52
         Text            =   "--Empty--"
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtArmour 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   51
         Text            =   "--Empty--"
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label24 
         Caption         =   "Helmet:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label23 
         Caption         =   "Shield:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label22 
         Caption         =   "Weapon:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label21 
         Caption         =   "Armour:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Frame fraStats 
      Caption         =   "Stats"
      Height          =   1575
      Left            =   120
      TabIndex        =   37
      Top             =   7560
      Width           =   4695
      Begin VB.TextBox txtEnd 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   43
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtStr 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   42
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtAgil 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   41
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtInt 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         TabIndex        =   40
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtWill 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   39
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtPoints 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         TabIndex        =   38
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "End:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   49
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "Str:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "Agil:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblInt 
         Caption         =   "Int:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   46
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label19 
         Caption         =   "Will:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   45
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label18 
         Caption         =   "Pts:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   44
         Top             =   720
         Width           =   375
      End
   End
   Begin VB.Frame fraInventory 
      Caption         =   "Inventory"
      Height          =   4455
      Left            =   4440
      TabIndex        =   32
      Top             =   120
      Width           =   5175
      Begin VB.ListBox lstInventory 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3765
         Left            =   240
         TabIndex        =   34
         Top             =   480
         Width           =   2295
      End
      Begin VB.ListBox lstItems 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3765
         Left            =   2640
         TabIndex        =   33
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label13 
         Caption         =   "Inventory:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Local Item List:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   35
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame fraLocation 
      Caption         =   "Location"
      Height          =   1215
      Left            =   240
      TabIndex        =   23
      Top             =   6240
      Width           =   4335
      Begin VB.TextBox txtMap 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   27
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txtX 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   480
         TabIndex        =   26
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtY 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   25
         Top             =   720
         Width           =   975
      End
      Begin VB.ComboBox cboDir 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmEditPlayer.frx":0000
         Left            =   3240
         List            =   "frmEditPlayer.frx":0010
         TabIndex        =   24
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Map:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   31
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "X:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "Y:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   29
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label11 
         Caption         =   "Dir:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   28
         Top             =   720
         Width           =   495
      End
   End
   Begin VB.Frame fraPlayer 
      Caption         =   "Player Data"
      Height          =   2655
      Left            =   360
      TabIndex        =   10
      Top             =   2280
      Width           =   3615
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   16
         Top             =   360
         Width           =   2295
      End
      Begin VB.ComboBox cboGender 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmEditPlayer.frx":002B
         Left            =   1200
         List            =   "frmEditPlayer.frx":0035
         TabIndex        =   15
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtLevel 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   14
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txtExp 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   13
         Top             =   1440
         Width           =   2295
      End
      Begin VB.ComboBox cboAccess 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmEditPlayer.frx":0047
         Left            =   1200
         List            =   "frmEditPlayer.frx":005A
         TabIndex        =   12
         Top             =   1800
         Width           =   2295
      End
      Begin VB.ComboBox cboPK 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmEditPlayer.frx":008B
         Left            =   1200
         List            =   "frmEditPlayer.frx":0095
         TabIndex        =   11
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Gender:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Level 
         Caption         =   "Level:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Experience:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Access:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "PK:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   975
      End
   End
   Begin VB.Frame fraAccount 
      Caption         =   "Account Data"
      Height          =   1215
      Left            =   360
      TabIndex        =   5
      Top             =   1080
      Width           =   3615
      Begin VB.TextBox txtLogin 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Login:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdSavePlayer 
      Caption         =   "Save Player File"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   2
      Top             =   9600
      Width           =   1575
   End
   Begin VB.TextBox txtAccountName 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Text            =   "Account Name Here"
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open Player File"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label26 
      Caption         =   "By Lightning"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   64
      Top             =   9600
      Width           =   1455
   End
   Begin VB.Label Label25 
      Caption         =   "Eclipse Origins Account Editor"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   63
      Top             =   9120
      Width           =   4935
   End
   Begin VB.Label lblNotify 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   4
      Top             =   4200
      Width           =   4695
   End
   Begin VB.Label Label12 
      Caption         =   "Version: 2.0"
      Height          =   255
      Left            =   5280
      TabIndex        =   3
      Top             =   9720
      Width           =   975
   End
End
Attribute VB_Name = "frmPlayerEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const NAME_LENGTH As Byte = 20
Const MAX_INV As Byte = 35
Const MAX_PLAYER_SPELLS As Byte = 35
Const ACCOUNT_LENGTH As Byte = 12
Const MAX_HOTBAR As Long = 12
Dim AccName As String
Dim OnePlayer As PlayerRec
Dim OneItem As ItemRec
Dim OneEq(1 To 4) As ItemRec

Private Type HotbarRec
    Slot As Long
    sType As Byte
End Type


Private Type PlayerInvRec
    Num As Long
    Value As Long
End Type

Private Type PlayerRec
        ' Account
    Login As String * ACCOUNT_LENGTH
    Password As String * NAME_LENGTH
    
    ' General
    Name As String * ACCOUNT_LENGTH
    Sex As Byte
    Class As Long
    Sprite As Long
    Level As Byte
    exp As Long
    Access As Byte
    PK As Byte
    
    ' Vitals
    Vital(1 To 2) As Long
    
    ' Stats
    Stat(1 To 5) As Byte
    POINTS As Long
    
    ' Worn equipment
    Equipment(1 To 4) As Long
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Long
    
    ' Hotbar
    Hotbar(1 To MAX_HOTBAR) As HotbarRec
    
    ' Position
    Map As Long
    x As Byte
    y As Byte
    Dir As Byte
End Type

Private Type ItemRec
    Name As String * NAME_LENGTH
    Desc As String * 255
    Sound As String * NAME_LENGTH
    
    Pic As Long

    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    ClassReq As Long
    AccessReq As Long
    LevelReq As Long
    Mastery As Byte
    price As Long
    Add_Stat(1 To 5) As Byte
    Rarity As Byte
    Speed As Long
    Handed As Long
    BindType As Byte
    Stat_Req(1 To 5) As Byte
    Animation As Long
    Paperdoll As Long
    
    AddHP As Long
    AddMP As Long
    AddEXP As Long
    CastSpell As Long
    instaCast As Byte
End Type

Private Type SpellRec
    Name As String * NAME_LENGTH
    Desc As String * 255
    Sound As String * NAME_LENGTH
    
    Type As Byte
    MPCost As Long
    LevelReq As Long
    AccessReq As Long
    ClassReq As Long
    CastTime As Long
    CDTime As Long
    Icon As Long
    Map As Long
    x As Long
    y As Long
    Dir As Byte
    Vital As Long
    Duration As Long
    Interval As Long
    Range As Byte
    IsAoE As Boolean
    AoE As Long
    CastAnim As Long
    SpellAnim As Long
    StunDuration As Long
End Type

Private Sub cmdArmourUnequip_Click()
If OnePlayer.Equipment(2) <> 0 Then
    For i = 1 To 35
        If OnePlayer.Inv(i).Num = 0 Then
            OnePlayer.Inv(i).Num = OnePlayer.Equipment(2)
            OnePlayer.Inv(i).Value = 1
            OnePlayer.Equipment(2) = 0
            UpdateInventory
            txtArmour.Text = "--Empty--"
            Exit Sub
        End If
    Next
Else
    MsgBox ("There is no item here!")
End If
End Sub

Private Sub cmdHelmetUnequip_Click()
If OnePlayer.Equipment(3) <> 0 Then
    For i = 1 To 35
        If OnePlayer.Inv(i).Num = 0 Then
            OnePlayer.Inv(i).Num = OnePlayer.Equipment(3)
            OnePlayer.Inv(i).Value = 1
            OnePlayer.Equipment(3) = 0
            UpdateInventory
            txtHelmet.Text = "--Empty--"
            Exit Sub
        End If
    Next
ElseIf unequipped = False Then
    MsgBox ("There is no space in your inventory!")
Else
    MsgBox ("There is no item here!")
End If
End Sub

Private Sub cmdOpen_Click()
    LoadPlayer
End Sub

Private Sub cmdSavePlayer_Click()
Dim Filename As String
Dim i As Integer

Filename = App.Path & "\data\accounts\" & AccName & ".bin"

If Not IsNumeric(txtExp.Text) Then
    MsgBox ("Experience must be a number!")
    Exit Sub
End If

If Not IsNumeric(txtLevel.Text) Then
    MsgBox ("Level must be a number!")
    Exit Sub
End If

If Not IsNumeric(txtX.Text) Then
    MsgBox ("X coordinate must be a number!")
    Exit Sub
End If

If Not IsNumeric(txtY.Text) Then
    MsgBox ("Y coordinate must be a number!")
    Exit Sub
End If

If Not IsNumeric(txtStr.Text) Then
    MsgBox ("Str variable must be a number!")
    Exit Sub
End If

If Not IsNumeric(txtEnd.Text) Then
    MsgBox ("End variable must be a number!")
    Exit Sub
End If


If Not IsNumeric(txtAgil.Text) Then
    MsgBox ("Vit variable must be a number!")
    Exit Sub
End If


If Not IsNumeric(txtWill.Text) Then
    MsgBox ("Will variable must be a number!")
    Exit Sub
End If


If Not IsNumeric(txtInt.Text) Then
    MsgBox ("Int variable must be a number!")
    Exit Sub
End If


If Not IsNumeric(txtPoints.Text) Then
    MsgBox ("Points variable must be a number!")
    Exit Sub
End If

If Not IsNumeric(txtHP.Text) Then
    MsgBox ("HP variable must be a number!")
    Exit Sub
End If

If Not IsNumeric(txtMP.Text) Then
    MsgBox ("MP variable must be a number!")
    Exit Sub
End If

With OnePlayer
    .Login = txtLogin.Text
    .Password = txtPassword.Text
    .Name = txtName.Text
    .Sex = cboGender.ListIndex
    .Level = txtLevel.Text
    .exp = txtExp.Text
    .Access = cboAccess.ListIndex
    .PK = cboPK.ListIndex
    
    .Vital(1) = txtHP.Text
    .Vital(2) = txtMP.Text
    
    .Stat(1) = txtStr.Text
    .Stat(2) = txtEnd.Text
    .Stat(3) = txtInt.Text
    .Stat(4) = txtAgil.Text
    .Stat(5) = txtWill.Text
    .POINTS = txtPoints.Text
    
    .Map = txtMap.Text
    .x = txtX.Text
    .y = txtY.Text
    .Dir = cboDir.ListIndex
End With

Open Filename For Binary As #1
    Put #1, , OnePlayer
Close #1

cmdSavePlayer.Enabled = False
MsgBox ("Account " & AccName & " saved.")
cmdOpen.Enabled = True
txtAccountName.Enabled = True
fraAccount.Enabled = False
fraPlayer.Enabled = False
fraLocation.Enabled = False
fraStats.Enabled = False
fraInventory.Enabled = False
fraEquipment.Enabled = False
fraSpells.Enabled = False
fraVitals.Enabled = False
    
End Sub

Sub UpdateInventory()
Dim i As Integer
Dim itemfilename As String, ItemNumber As String

lstInventory.Clear
    For i = 1 To MAX_INV
            ItemNumber = OnePlayer.Inv(i).Num
            itemfilename = App.Path & "\data\items\Item" & ItemNumber & ".dat"
            Open itemfilename For Binary As #1
                Get #1, , OneItem
                lstInventory.AddItem (i & ": " & OneItem.Name)
            Close #1
    Next
    
    'MsgBox OnePlayer.Inv(35).Num & " " & OnePlayer.Inv(35).Value
        
End Sub

Sub LoadItemList()
Dim i As Integer
Dim Filename As String

lstItems.Clear
lstItems.AddItem ("--None--")
    For i = 1 To 255
        Filename = App.Path & "\data\items\Item" & i & ".dat"
        
        Open Filename For Binary As #1
            Get #1, , OneItem
            lstItems.AddItem (i & ": " & OneItem.Name)
        Close #1
    Next
End Sub


Private Sub cmdShieldUnequip_Click()
If OnePlayer.Equipment(4) <> 0 Then
    For i = 1 To 35
        If OnePlayer.Inv(i).Num = 0 Then
            OnePlayer.Inv(i).Num = OnePlayer.Equipment(4)
            OnePlayer.Inv(i).Value = 1
            OnePlayer.Equipment(4) = 0
            UpdateInventory
            txtWeapon.Text = "--Empty--"
            unequipped = True
            Exit Sub
        End If
    Next
Else
    MsgBox ("There is no item here!")
End If
End Sub

Private Sub cmdWeaponUnequip_Click()
If OnePlayer.Equipment(1) <> 0 Then
    For i = 1 To 35
        If OnePlayer.Inv(i).Num = 0 Then
            OnePlayer.Inv(i).Num = OnePlayer.Equipment(1)
            OnePlayer.Inv(i).Value = 1
            OnePlayer.Equipment(1) = 0
            UpdateInventory
            txtWeapon.Text = "--Empty--"
            unequipped = True
        End If
    Next
Else
    MsgBox ("There is no item here!")
End If
End Sub

Private Sub Form_Load()
Dim Filename As String, store As String, iStore As Integer
Dim i As Integer

fraAccount.Enabled = False
fraPlayer.Enabled = False
fraLocation.Enabled = False
fraStats.Enabled = False
fraInventory.Enabled = False
fraEquipment.Enabled = False
fraVitals.Enabled = False
fraSpells.Enabled = False

Filename = App.Path & "\data\accounts\charlist.txt"

Open Filename For Input As #1

Do Until EOF(1)
    Input #1, store
    i = i + 1
Loop
iStore = i
Close #1

Filename = App.Path & "\data\accounts\"
For i = 1 To iStore
    
Next


End Sub

Private Sub lstItems_Click()
Dim ItmIndex As Integer

ItmIndex = lstItems.ListIndex

If lstInventory.ListIndex >= 0 Then
    If ItmIndex = 0 Then
        OnePlayer.Inv(lstInventory.ListIndex + 1).Num = 0
        OnePlayer.Inv(lstInventory.ListIndex + 1).Value = 1
        lblNotify.Caption = "Item removed."
    Else
        OnePlayer.Inv(lstInventory.ListIndex + 1).Num = ItmIndex
        OnePlayer.Inv(lstInventory.ListIndex + 1).Value = 1
        lblNotify.Caption = "Item added."
    End If
    
    UpdateInventory
End If
End Sub

Private Sub lstPlayers_Click()
    LoadPlayer
End Sub

Sub LoadPlayer()
Dim Filename As String

AccName = txtAccountName.Text
Filename = App.Path & "\data\accounts\" & AccName & ".bin"

If LenB(Dir(Filename)) > 0 Then
    Open Filename For Binary As #1
    Get #1, , OnePlayer
    Close #1
    
    With OnePlayer
    
        'account data
        txtLogin.Text = .Login
        txtPassword.Text = .Password
        
        'player data
        txtName.Text = .Name
        
        If .Sex = 0 Then
            cboGender.ListIndex = 0
        Else
            cboGender.ListIndex = 1
        End If
        
        txtLevel.Text = .Level
        txtExp.Text = .exp
        
        'other player data
        cboAccess.ListIndex = .Access
        cboPK.ListIndex = .PK
        
        'coordinates
        txtMap.Text = .Map
        txtX.Text = .x
        txtY.Text = .y
        cboDir.ListIndex = .Dir
        
        'stats
        txtStr.Text = .Stat(1)
        txtEnd.Text = .Stat(2)
        txtInt.Text = .Stat(3)
        txtAgil.Text = .Stat(4)
        txtWill.Text = .Stat(5)
        txtPoints.Text = .POINTS
        
        'vitals
        txtHP.Text = .Vital(1)
        txtMP.Text = .Vital(2)
        
        'equipment
        LoadEquipment
        
        'inventories
        UpdateInventory
        LoadItemList
        
        'spells
        LoadPlayerSpells
        LoadSpellList
        
        fraAccount.Enabled = True
        fraPlayer.Enabled = True
        fraLocation.Enabled = True
        fraStats.Enabled = True
        fraInventory.Enabled = True
        fraEquipment.Enabled = True
        fraSpells.Enabled = True
        fraVitals.Enabled = True
        cmdSavePlayer.Enabled = True
        txtAccountName.Enabled = False
        cmdOpen.Enabled = False
        
    End With
Else
    MsgBox ("Player File does not exist!")
    Exit Sub
End If
End Sub

Sub LoadEquipment()
Dim i As Integer

txtWeapon.Text = "--Empty--"
txtArmour.Text = "--Empty--"
txtHelmet.Text = "--Empty--"
txtShield.Text = "--Empty--"

For i = 1 To 4
    ItemNumber = OnePlayer.Equipment(i)
    
    If ItemNumber = 0 Then
        OneEq(i).Name = "--Empty--"
    Else
        itemfilename = App.Path & "\data\items\Item" & ItemNumber & ".dat"
        Open itemfilename For Binary As #1
            Get #1, , OneEq(i)
        Close #1
    End If
Next

txtWeapon.Text = OneEq(1).Name
txtArmour.Text = OneEq(2).Name
txtHelmet.Text = OneEq(3).Name
txtShield.Text = OneEq(4).Name
        

End Sub

Sub LoadSpellList()
Dim Filename As String
Dim i As Integer
Dim OneSpell As SpellRec

lstSpellList.Clear

lstSpellList.AddItem ("--None--")

For i = 1 To 255
    Filename = App.Path & "\data\spells\spells" & i & ".dat"
    Open Filename For Binary As #1
        Get #1, , OneSpell
    Close #1
    
    lstSpellList.AddItem i & ": " & OneSpell.Name
Next
End Sub

Sub LoadPlayerSpells()
Dim Filename As String
Dim i As Integer
Dim OneSpell As SpellRec

lstPlayerSpells.Clear

For i = 1 To MAX_PLAYER_SPELLS
            SpellNumber = OnePlayer.Spell(i)
            itemfilename = App.Path & "\data\spells\spells" & SpellNumber & ".dat"
            Open itemfilename For Binary As #1
                Get #1, , OneSpell
                lstPlayerSpells.AddItem (i & ": " & OneSpell.Name)
            Close #1
    Next
End Sub

Private Sub lstSpellList_Click()
Dim SpellIndex As Integer

SpellIndex = lstSpellList.ListIndex

If lstPlayerSpells.ListIndex >= 0 Then
    If ItmIndex = 0 Then
        OnePlayer.Spell(lstPlayerSpells.ListIndex + 1) = SpellIndex
    Else
        OnePlayer.Inv(lstPlayerSpells.ListIndex + 1).Num = SpellIndex
    End If
    
    LoadPlayerSpells
End If
End Sub
