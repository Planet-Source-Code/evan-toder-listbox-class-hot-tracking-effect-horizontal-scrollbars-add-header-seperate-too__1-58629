VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cListHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As Rect, ByVal wFormat As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As Rect, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function LBItemFromPt Lib "comctl32" (ByVal hLB As Long, ByVal X As Long, ByVal Y As Long, ByVal bAutoScroll As Boolean) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As Rect, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SendMessageArray Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetRect Lib "user32" (lpRect As Rect, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type
 
 
Private Type Rect
   Left As Long
   Top  As Long
   Right As Long
   Bottom As Long
End Type

Enum enHelpItem
    function_draw_header = 0
    sub_additems_with_tabs = 1
    sub_horiz_scrollbars = 2
    function_search = 3
    sub_initialize_listbox = 4
    sub_hot_tracking = 5
    sub_individual_item_tooltips = 6
    sub_tabstops = 7
    sub_add_many_items = 8
End Enum

Enum enStyle
   style_raised = 4
   style_framed = 6
End Enum
 
 
Dim cRect                        As cdraw_rect
Dim m_b_hottracking              As Boolean
Dim m_b_individualtooltips       As Boolean
Private WithEvents yourlistbox   As ListBox
Attribute yourlistbox.VB_VarHelpID = -1


Sub initialize_listbox(your_list_box As Object)
   
   Set yourlistbox = your_list_box
   
End Sub
Sub individual_item_tooltips(benable As Boolean)

  m_b_individualtooltips = benable
  
End Sub

Sub hot_tracking(benable As Boolean)
 
 'see yourListbox_mousemove
  m_b_hottracking = benable

End Sub

'======================================
' RETURN: a long value the will be the width
'         of the usercontrol
'======================================
Function draw_header(str_head$, _
                     header_style As enStyle, _
                     header_text_color As Long, _
                     container As Object)
                      
 '
 'variable declarations
 Dim vparts      As Variant
 Dim upper&, left_point&, left_edge&, top_point&, lcnt&
 Dim r()  As Rect
 '
 'constants for drawing text
 Const DT_CALCRECT As Long = &H400
 Const DT_LEFT As Long = &H0
 Const DT_CENTER As Long = &H1
 Const DT_SINGLELINE As Long = &H20
 Const DT_CALC = (DT_CENTER Or DT_SINGLELINE Or DT_CALCRECT)
 Const DT_DRAW = (DT_CENTER Or DT_SINGLELINE)
 
 '
 ' user supplies the header in the format
 '  "string | string |  string"
 ' we will set the column width based upon
 ' the width of each of the string parts
 
  vparts = Split(str_head, "|")
  '
  'user didnt supply any header text so exit ==>
  If Not (IsArray(vparts)) Then Exit Function

  upper = UBound(vparts)
  ReDim r(upper)
  '
  'clear the old drawing
  container.Cls
  '
  'starting left point for the first rect
  left_point = (yourlistbox.Left / Screen.TwipsPerPixelX)
  
  left_edge = left_point
  
  top_point = (yourlistbox.Top / Screen.TwipsPerPixelY) - _
              (container.TextHeight("HELLO") / Screen.TwipsPerPixelY)
  
  '
  'class object reference
  Set cRect = New cdraw_rect
  Set cRect.your_obj_pallete = container
  '
  For lcnt = 0 To upper
      '
      'set preliminary rect that will be adjusted
      'based upon the text val of this part
      SetRect r(lcnt), 0, 0, 0, 0
      '
      'the rect gets recalculated here
      DrawText container.hdc, vparts(lcnt), Len(vparts(lcnt)), r(lcnt), DT_CALC
      '
      'offset the rect so the rects lefts line up one after the other
      OffsetRect r(lcnt), left_point, top_point
      '
      'add a little padding
      InflateRect r(lcnt), 0, 1
      '
      'draw the border for each of the rects
      cRect.draw_rect r(lcnt).Left, r(lcnt).Top, r(lcnt).Right, r(lcnt).Bottom _
                      , svpixels, , header_style
      '
      'the color of the header text
      SetTextColor container.hdc, header_text_color
      '
      'now draw the caption
      DrawText container.hdc, vparts(lcnt), Len(vparts(lcnt)), r(lcnt), DT_DRAW
      '
      'keep track of where the next rects left should be
      left_point = (left_point + (r(lcnt).Right - r(lcnt).Left))
   Next lcnt
  '
  ' width of the header after all the formatting.
  ' set listbox to the same width
  yourlistbox.Width = (left_point * Screen.TwipsPerPixelX) - _
                      (left_edge * Screen.TwipsPerPixelX)
  '
  Set cRect = Nothing
  Erase r

End Function
'
'create owner defined tabstop points for the listbox
'
Sub set_tabstops(ParamArray TabPoints())

  Dim upper&, lcnt&   'long
  Dim LBTab() As Long
  '
  'const for setting listbox tabstops
  Const LB_SETTABSTOPS = &H192
  '
  'here we will set the tab points for the listbox
  upper = UBound(TabPoints) + 1
  ReDim LBTab(1 To upper)
  '
  'param array returns variant..we must convert to longs
  For lcnt = 1 To upper
      LBTab(lcnt) = CLng(TabPoints(lcnt - 1))
  Next lcnt
  
  SendMessageArray yourlistbox.hwnd, LB_SETTABSTOPS, upper, LBTab(1)
  
End Sub

'
'this sub basically takes a string like so...
' "ityoure-mail@yourdomain.com | item2 | item3 "
'and turns the | into tabs
Sub additems_with_tabs(str_items As String)
 
 Dim new_str As String
 
 new_str = Join(Split(str_items, "|"), vbTab)
 
 yourlistbox.AddItem new_str

End Sub
'
'ADD MULTIPLE ITEMS TO A LISTBOX WITH ONE LINE
'
Sub add_many_items(ParamArray ItemsToAdd())
 
Dim lcnt&, upper& 'long

On Error GoTo local_error:

  upper = UBound(ItemsToAdd)
  
  For lcnt = 0 To upper
     '-- add the item to listbox
     yourlistbox.AddItem CStr(ItemsToAdd(lcnt))
  Next lcnt
 
Exit Sub
local_error:
 
End Sub

Sub horiz_scrollbars(your_form As Object, Optional scroll_width As Long)

  Const LB_SETHORIZONTALEXTENT As Long = &H194
  Dim i%                 'integer
  Dim new_len&, max_len& 'long
  
  On Error GoTo local_error:

  With yourlistbox
    '
    'if user supplies a scroll width then he wants to supply the
    'value and we wont attempt to calculate
    If scroll_width <= 0 Then
       For i = 0 To (.ListCount - 1)
           new_len = (your_form.ScaleX( _
                    your_form.TextWidth(.List(i)), _
                    your_form.ScaleMode, vbPixels) * 1.1)
   
           If scroll_width < new_len Then scroll_width = new_len
       Next i
     End If
     
    SendMessage .hwnd, LB_SETHORIZONTALEXTENT, scroll_width, 0
    
  End With
  
Exit Sub
local_error:
 
End Sub
'
'search listbox for an item that matches
'the return is the index of the item found in the list
'or -1 if its not found
Function search(str_item As String, _
               Optional select_found_item As Boolean) As Long
  
  Const LB_FINDSTRING = &H18F
 
  search_listbox = SendMessage( _
       yourlistbox.hwnd, LB_FINDSTRING, -1, ByVal str_item)
  
  '
  'select the item found ?
  If select_found_item Then
     yourlistbox.ListIndex = search_listbox
  End If
  
End Function

'
'if the user wants the hot tracking effect of for the tooltips
'to update to reflect the listitem under the mouse
'
Private Sub yourlistbox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  If m_b_hottracking Or m_b_individualtooltips Then
  
    With yourlistbox
        Dim ret, pt As POINTAPI
        GetCursorPos pt
        ret = LBItemFromPt(.hwnd, pt.X, pt.Y, True)
        
        If m_b_hottracking Then .ListIndex = ret
      
        If m_b_individualtooltips Then .ToolTipText = .List(ret)
    End With
  End If

End Sub

Function HOW_TO_HELP(help_item As enHelpItem, bshow_in_messagebox As Boolean)
  
  If help_item = function_draw_header Then
     HOW_TO_HELP = "This function draws an attractive header at the top of your listbox so it looks more like a listview control. The string you provide should be the text you want for your header, seperated by pipe ""|"" characters. The [container] parameter is the object the listbox is drawn on..ie the form or frame or picturebox"
     
  ElseIf help_item = sub_additems_with_tabs Then
     HOW_TO_HELP = "This sub allows you to additems to the listbox in such a manner that string  ""string|string2|string3"" takes the pipe character and converts it to tabs"
     
  ElseIf help_item = sub_horiz_scrollbars Then
     HOW_TO_HELP = "Allows you to add horizontal scrollbars to the listbox.  If [scroll_width] is not supplied then the items in the list are scanned and if horizontal scrollbars are required, then it is added, making the listbox wide enouph to view the widest list entry. If  [scroll_width] is supplied then you are manually adding them, the width based upon the value you supply for   [scroll_width]."
     
  ElseIf help_item = function_search Then
     HOW_TO_HELP = "Allows you to search for an item in the listbox.  This acts like the autocomplete feature of browsers in that the search is case insensitive, and a partial match returns the index of the first listitem matching so a search of ""VIS"", for example, will return the index of ""visual basic"" entry in the listbox."
     
  ElseIf help_item = sub_initialize_listbox Then
     HOW_TO_HELP = "This must be the first method called and it sets reference to the listbox you will be referring to in this class."
     
  ElseIf help_item = sub_hot_tracking Then
     HOW_TO_HELP = "Calling this sub and setting [benable] to True makes your listbox have a ""hot-tracking"" effect in that moving your mouse over the listbox causes the item under your mouse to be hilighted AND the tooltip text constantly updates to reflect the list item under the mouse so even if only part of the list entry is visible, the full list entry will show in the tooltip text."
     
  ElseIf help_item = sub_individual_item_tooltips Then
     HOW_TO_HELP = "The tooltip for the listbox displays the contents of the listitem the mouse is over. Useful if the contents of listitems width exceed the width of the listbox itself"
     
  ElseIf help_item = sub_tabstops Then
     HOW_TO_HELP = "Set tabstop points for the listbox."
     
  ElseIf help_item = sub_additems_with_tabs Then
     HOW_TO_HELP = "Allows you to add items to the listbox with tabs inclusive, use a pipe character within the string to specify vbTab..i.e ""string1 | string2 | string3"""
  
  ElseIf help_item = sub_add_many_items Then
     HOW_TO_HELP = "Enables you to add multiple items with one call.  Each item to add is seperated by a comma."
     
  End If
  
  If bshow_in_messagebox Then MsgBox HOW_TO_HELP
  
End Function

 
Private Sub Class_Terminate()

  Set yourlistbox = Nothing

End Sub
