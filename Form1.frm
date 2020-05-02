VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4125
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDetails 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   5655
   End
   Begin VB.TextBox txtFooter 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   3480
      Width           =   5655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call renderdata(markuptodct(loadtextfromfile("c:\sandbox\textmark.txt")))
End Sub

Sub displaydata(ByVal datatype As String, ByVal data As Collection)
    Dim idx As Integer
    Dim linedetails As String
    Dim parts() As String
    Dim headers() As String
    Dim col As Integer
    
    Select Case LCase(datatype)
        Case "header"
            Me.Caption = data.Item(1)
            
        Case "details"
            Me.txtDetails.Text = ""
            headers = Split(data.Item(1), Space(1), 3)
            For idx = 2 To data.Count
                parts = Split(data.Item(idx), Space(1), 3)
                linedetails = vbNullString
                For col = LBound(headers) To UBound(headers)
                    linedetails = linedetails _
                        & headers(col) & " = " & parts(col) & vbCrLf
                Next
                Me.txtDetails.SelStart = Len(Me.txtDetails.Text)
                Me.txtDetails.SelLength = Len(linedetails)
                Me.txtDetails.SelText = linedetails
            Next
            
        Case "footer"
            Me.txtFooter.Text = data.Item(1)
            
        Case Else
            ' do whatever, if not expected
            
    End Select
End Sub

Function loadtextfromfile(ByVal fn As String) As String
    Dim content As String
    Open fn For Input As #1
    content = Input(LOF(1), 1)
    Close #1
    loadtextfromfile = content
End Function

Function markuptodct(ByVal markuptext As String) As Scripting.Dictionary
    Dim lines() As String
    Dim line As String
    Dim idx As Integer
    Dim dct As New Scripting.Dictionary
    Dim tag As String
    Dim content As Collection
    
    lines = Split(markuptext, vbCrLf)
    tag = vbNullString
    
    For idx = LBound(lines) To UBound(lines)
        line = Trim(lines(idx))
        If line Like "<*>" Then
            ' it is a tag
            tag = Mid(line, 2, Len(line) - 2)
            Set dct(tag) = New Collection
            Set content = dct(tag)
        Else
            ' it is data
            If tag > "" Then
                ' add to content
                content.Add line
            Else
                ' if no tag, consider garbage; ignore
            End If
        End If
    Next
    
    Set markuptodct = dct
End Function

Sub renderdata(ByVal data As Scripting.Dictionary)
    Dim datakey As Variant
    
    If data Is Nothing Then
        MsgBox "Invalid data", vbExclamation
    Else
        For Each datakey In data.Keys
            displaydata datakey, data(datakey)
        Next
    End If
End Sub

