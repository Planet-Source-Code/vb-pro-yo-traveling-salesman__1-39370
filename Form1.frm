VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3450
   ClientLeft      =   5370
   ClientTop       =   4125
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   9240
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   5160
      TabIndex        =   8
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   5160
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   5160
      Width           =   615
   End
   Begin VB.ListBox List3 
      Height          =   1230
      Left            =   4920
      TabIndex        =   3
      Top             =   600
      Width           =   1455
   End
   Begin VB.ListBox List2 
      Height          =   2790
      Left            =   6480
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   7920
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   495
      Left            =   5400
      TabIndex        =   0
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "6"
      Height          =   255
      Left            =   2040
      TabIndex        =   17
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label10 
      Caption         =   "5"
      Height          =   255
      Left            =   1680
      TabIndex        =   16
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label9 
      Caption         =   "4"
      Height          =   255
      Left            =   1320
      TabIndex        =   15
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label8 
      Caption         =   "3"
      Height          =   255
      Left            =   960
      TabIndex        =   14
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label7 
      Caption         =   "2"
      Height          =   255
      Left            =   600
      TabIndex        =   13
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label6 
      Caption         =   "1"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2880
      Width           =   135
   End
   Begin VB.Shape Shape12 
      BorderColor     =   &H000000FF&
      Height          =   135
      Left            =   240
      Shape           =   3  'Circle
      Top             =   3240
      Width           =   135
   End
   Begin VB.Shape Shape11 
      BorderColor     =   &H000080FF&
      Height          =   135
      Left            =   600
      Shape           =   3  'Circle
      Top             =   3240
      Width           =   135
   End
   Begin VB.Shape Shape10 
      Height          =   135
      Left            =   960
      Shape           =   3  'Circle
      Top             =   3240
      Width           =   135
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H00C000C0&
      Height          =   135
      Left            =   1320
      Shape           =   3  'Circle
      Top             =   3240
      Width           =   135
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00FFFFFF&
      Height          =   135
      Left            =   1680
      Shape           =   3  'Circle
      Top             =   3240
      Width           =   135
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H0000FF00&
      Height          =   135
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   3240
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      Height          =   135
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label5 
      Caption         =   "Distance"
      Height          =   255
      Left            =   8160
      TabIndex        =   11
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Order"
      Height          =   255
      Left            =   6960
      TabIndex        =   10
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Cord's"
      Height          =   255
      Left            =   5280
      TabIndex        =   9
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Lowest Dist.:"
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Best Order is:"
      Height          =   255
      Left            =   4080
      TabIndex        =   5
      Top             =   3120
      Width           =   975
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FFFFFF&
      Height          =   135
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C000C0&
      Height          =   135
      Left            =   1680
      Shape           =   3  'Circle
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape Shape3 
      Height          =   135
      Left            =   1560
      Shape           =   3  'Circle
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000080FF&
      Height          =   135
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   135
      Left            =   1320
      Shape           =   3  'Circle
      Top             =   240
      Width           =   135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim corrx(1 To 10)
Dim corry(1 To 10)
Dim a As Long
Private Sub Command1_Click()
For a = 1 To 6
    For b = 1 To 6
        If b = a Then
        GoTo 2
        End If
            For c = 1 To 6
                If c = b Then
                GoTo 3
                End If
                If c = a Then
                GoTo 3
                End If

                    For d = 1 To 6
                    If d = a Then
                    GoTo 4
                    End If
                    If d = b Then
                    GoTo 4
                    End If
                    If d = c Then
                    GoTo 4
                    End If
                                For e = 1 To 6
                                If e = a Then
                                GoTo 5
                                End If
                                If e = b Then
                                GoTo 5
                                End If
                                If e = c Then
                                GoTo 5
                                End If
                                If e = d Then
                                GoTo 5
                                End If
                                        For f = 1 To 6
                                        If f = a Then
                                        GoTo 6
                                        End If
                                        If f = b Then
                                        GoTo 6
                                        End If
                                        If f = c Then
                                        GoTo 6
                                        End If
                                        If f = d Then
                                        GoTo 6
                                        End If
                                        If f = e Then
                                        GoTo 6
                                        End If
                                            Call checkfast(a, b, c, d, e, f)
6
                                        Next f
5
                                Next e
4
                    Next d

       
3
            Next c
2
    Next b
Next a
Call thebest
End Sub

Private Sub Form_Load()
Form1.Show
beginprog
'For i = 1 To 10
'Print corrx(i); corry(i)
'Next i
End Sub

Private Sub List1_Click()
List2.ListIndex = List1.ListIndex
End Sub

Private Sub List2_Click()
List1.ListIndex = List2.ListIndex
End Sub
