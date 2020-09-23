VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   465
      Left            =   45
      TabIndex        =   4
      Top             =   495
      Width           =   4965
      Begin VB.CheckBox ckTooltips 
         Caption         =   "individual line tooltips"
         Height          =   195
         Left            =   1485
         TabIndex        =   6
         Top             =   180
         Width           =   1860
      End
      Begin VB.CheckBox ckHotTracking 
         Caption         =   "hot tracking"
         Height          =   195
         Left            =   135
         TabIndex        =   5
         Top             =   180
         Width           =   1365
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2745
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   45
      Width           =   2445
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   1305
      TabIndex        =   2
      Top             =   45
      Width           =   1365
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search lBox"
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Top             =   45
      Width           =   1185
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   135
      TabIndex        =   0
      Top             =   1305
      Width           =   4740
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


 
 
Private cHeader   As cListHeader

 
 
Private Sub Check1_Click()

End Sub

Private Sub ckHotTracking_Click()

  'hottracking effect
  cHeader.hot_tracking CLng(ckHotTracking.Value)
  
End Sub

Private Sub ckTooltips_Click()
  '
  'tooltips shows item under mouse
  cHeader.individual_item_tooltips CLng(ckTooltips.Value)
  
End Sub

'
'lets search for something in the listbox
Private Sub cmdSearch_Click()
 
 Dim ret_val   As Long
 
 '
 'find something in the listbox and autohilite the found item
 ret_val = cHeader.search_listbox(txtSearch.Text, True)
 
 If ret_val < 0 Then
    MsgBox "not found"
 Else
    MsgBox List1.List(ret_val)
 End If
 
End Sub

Private Sub Combo1_Click()

  cHeader.HOW_TO_HELP Combo1.ListIndex, True

End Sub

Private Sub Form_Load()
  
  With Combo1
       .Text = "help/info"
       .AddItem "how to: function_draw_header"
       .AddItem "how to: sub_additems_with_tabs"
       .AddItem "how to: sub_horiz_scrollbars"
       .AddItem "how to: function_search"
       .AddItem "how to: sub_initialize_listbox"
       .AddItem "how to: sub_hot_tracking"
       .AddItem "how to: sub_individual_item_tooltips"
       .AddItem "how to: sub_tabstops"
       .AddItem "how to: sub_add_many_items"
  End With
   
  Me.AutoRedraw = True
  Set cHeader = New cListHeader
  cHeader.initialize_listbox List1
  '
  'this one line of code sets the listbox header and sets its tabpoints
  cHeader.draw_header _
    "   part #  |     department      |   price   |  availability      " _
   , style_raised, vbBlue, Form1
  
  '
  'create tabpoints
  cHeader.set_tabstops 30, 86, 115
  '
   'add some listitems
  cHeader.additems_with_tabs "3433|parts|5.99|back ordered (available May 2009)"
  cHeader.additems_with_tabs "345|parts|77.89|on display"
  cHeader.additems_with_tabs "66|parts|15.99|back ordered"
  cHeader.additems_with_tabs "6767|parts|11.99|out of stock"
  cHeader.additems_with_tabs "3443|parts|1.99|available"
  cHeader.additems_with_tabs "111|parts|177.89|on display"
  cHeader.additems_with_tabs "1|parts|115.99|available"
  cHeader.additems_with_tabs "122223|parts|111.99|out of stock"
  cHeader.additems_with_tabs "37|parts|2.99|back ordered"
  
  'do we need a horizontal scrollbar
  cHeader.horiz_scrollbars Me
   
End Sub

 

Private Sub Form_Terminate()
   
  Set cHeader = Nothing
  
End Sub
