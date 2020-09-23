VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Even and Odd Numbers"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMethod 
      Caption         =   "Method 3"
      Height          =   330
      Index           =   2
      Left            =   105
      TabIndex        =   15
      Top             =   4830
      Width           =   1065
   End
   Begin VB.TextBox txtMax 
      Height          =   285
      Left            =   2730
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "100000"
      Top             =   3255
      Width           =   1005
   End
   Begin VB.TextBox txtMin 
      Height          =   285
      Left            =   105
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "1"
      Top             =   3255
      Width           =   1275
   End
   Begin ComCtl2.UpDown udMin 
      Height          =   285
      Left            =   1365
      TabIndex        =   10
      Top             =   3255
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   327681
      Value           =   1
      BuddyControl    =   "txtMin"
      BuddyDispid     =   196620
      OrigLeft        =   1575
      OrigTop         =   3255
      OrigRight       =   1815
      OrigBottom      =   3480
      Increment       =   1000
      Max             =   10000000
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton cmdMethod 
      Caption         =   "Method 2"
      Height          =   330
      Index           =   1
      Left            =   105
      TabIndex        =   6
      Top             =   4410
      Width           =   1065
   End
   Begin VB.CommandButton cmdMethod 
      Caption         =   "Method 1"
      Height          =   330
      Index           =   0
      Left            =   105
      TabIndex        =   5
      Top             =   3990
      Width           =   1065
   End
   Begin VB.CheckBox chkBenchMark 
      Caption         =   "Benchmark only (time speeds)"
      Height          =   225
      Left            =   105
      TabIndex        =   4
      Top             =   3675
      Width           =   2535
   End
   Begin VB.ListBox lstOdd 
      Height          =   2595
      Left            =   2100
      TabIndex        =   1
      Top             =   420
      Width           =   1905
   End
   Begin VB.ListBox lstEven 
      Height          =   2595
      Left            =   105
      TabIndex        =   0
      Top             =   420
      Width           =   1905
   End
   Begin ComCtl2.UpDown udMax 
      Height          =   285
      Left            =   3780
      TabIndex        =   12
      Top             =   3255
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   327681
      Value           =   100000
      BuddyControl    =   "txtMax"
      BuddyDispid     =   196622
      OrigLeft        =   1575
      OrigTop         =   3255
      OrigRight       =   1815
      OrigBottom      =   3480
      Increment       =   10000
      Max             =   10000000
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.Label Label4 
      Caption         =   "A different way of checking if the number is diviisible by 2."
      Height          =   435
      Left            =   1260
      TabIndex        =   16
      Top             =   4830
      Width           =   2745
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "to"
      Height          =   225
      Left            =   1785
      TabIndex        =   14
      Top             =   3255
      Width           =   855
   End
   Begin VB.Label lblBenchMark 
      Caption         =   "Time:"
      Height          =   225
      Left            =   2730
      TabIndex        =   9
      Top             =   3675
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   "Check if the last bit of the number is set (meaning it's odd)"
      Height          =   435
      Left            =   1260
      TabIndex        =   8
      Top             =   4410
      Width           =   2745
   End
   Begin VB.Label Label1 
      Caption         =   "Check if the number is divisible by 2"
      Height          =   435
      Left            =   1260
      TabIndex        =   7
      Top             =   3990
      Width           =   2850
   End
   Begin VB.Label lblOdd 
      Caption         =   "Odd Numbers"
      Height          =   225
      Left            =   2310
      TabIndex        =   3
      Top             =   105
      Width           =   1275
   End
   Begin VB.Label lblEven 
      Caption         =   "Even Numbers"
      Height          =   225
      Left            =   210
      TabIndex        =   2
      Top             =   105
      Width           =   1275
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' -----------------------------------------------------------------------------
'
' Copyright (c) 2004 Chris Waddell
'
' A program which shows 3 methods for calculating whether or not a number is
' even or odd.
'
' Method 1 uses the well known, and widely used method of checking if the
' number is evenly divisible by 2 by checking if "Number MOD 2 = 0"
'
' Method 2 uses a less widely known, but still often used method of checking
' if the last bit of the number is not set, indicating it is even:
' "Number AND 1 = 0"
'
' Method 3 uses another method of checking if the number is divisible by 2.
' It does this by checking if the number divided by 2 is the same as the integer
' of the number divided by 2 (i.e. if the number divided by 2 is whole)
'
'
' This code was written because I got tired of adding the same paragraphs of
' comments to "Fateha Rahman"s attempts at this code, explaining how to do it
' properly.
'
' So here is a proper implementation.
' It's nothing spectacular but at least it's done right.
'
' -----------------------------------------------------------------------------


Private Sub CheckNumbers(MethodNumber As Integer)

  Dim min       As Long
  Dim max       As Long
  Dim i         As Long
  Dim results() As Boolean
    
    ' No need to perform error checking. We know the contents will always be
    ' numeric since the edit fields are locked, and the only way of changing them
    ' is with the numeric up down control.
    min = CLng(txtMin.Text)
    max = CLng(txtMax.Text)

    ' If bench marking is turned on then start the timer
    If chkBenchMark.Value = vbChecked Then
        StopWatch.StartTimer
    End If
    
    ' Allocate a temporary array of booleans for the results. True indicates an
    ' even number. False indicates an odd number.
    ReDim results(min To max)
    
    ' Now perform the actual method - this should work very fast.
    ' In most cases, less than a second easily, even for the slowest method.
    
    If MethodNumber = 1 Then                        ' Moderate Speed
        For i = min To max
              results(i) = (i Mod 2&) = 0&
        Next i
    ElseIf MethodNumber = 2 Then                    ' Fastest speed
        For i = min To max
              results(i) = (i And 1&) = 0&
        Next i
    ElseIf MethodNumber = 3 Then                    ' Slowest speed
        For i = min To max
              results(i) = (i / 2&) = (i \ 2&)
        Next i
    End If
    
    ' If we were bench marking the results then display the time.
    If chkBenchMark.Value = vbChecked Then
        StopWatch.StopTimer
        lblBenchMark.Caption = "Time: " & StopWatch.GetTime
    End If
        
    ' Ask the user if they wish to add the results to the list
    If MsgBox("Finished calculating the results." & vbCrLf & "Do you wish add the results to the list?" & vbCrLf & "Note that this could take a short while.", vbYesNo, "Add Results?") = vbYes Then
    
        ' Hiding the lists makes the values add faster, since no repainting is required
        lstEven.Visible = False
        lstOdd.Visible = False
    
        ' Go through each number, adding it to the correct list
        For i = min To max
            If results(i) Then
                lstEven.AddItem CStr(i)
            Else
                lstOdd.AddItem CStr(i)
            End If
        Next i
        
        ' Now display the lists again
        lstEven.Visible = True
        lstOdd.Visible = True
        
    End If

    Erase results

End Sub


Private Sub cmdMethod_Click(Index As Integer)
    CheckNumbers Index + 1
End Sub


Private Sub Form_Load()
    If Not StopWatch.CheckSupport Then
        MsgBox "Your hardware does not support a high resolution timer. You will not be able to perform benchmarking (speed tests)"
        chkBenchMark.Enabled = False
    End If
End Sub
