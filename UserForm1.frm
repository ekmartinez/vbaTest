VERSION 5.00

Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4680
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   6420
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End

Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()

    Dim Ddate As Date
    Dim WB As Workbook
    Dim ws As Worksheet
    
    Label6.Caption = Now

    With ComboBox1
        .AddItem "cayey"
        .AddItem "Jayuya"
    End With

    With ThisWorkbook
        Set ws = ThisWorkbook.Sheets.Add(After:= _
        .Sheets(.Sheets.Count))
        ws.Name = "dBase"
    End With

    On Error Resume Next

    For Each ws In Workbook
        ws.Delete
        Next

End Sub

Private Sub Calculate_Click()
    Label1.Caption = Val(TextBox1) + Val(TextBox2)
    Label2.Caption = Val(TextBox3) + Val(TextBox4)
    Label3.Caption = Val(TextBox1) + Val(TextBox3)
    Label4.Caption = Val(Label1) + Val(Label2)
    Label6.Caption = Now

    If ComboBox1.Value = "cayey" Then
        MsgBox ("cayey")
    Else
        MsgBox ("Jayuya")
    End If

End Sub

Private Sub CommandButton1_Click()
    TextBox1.Value = ""
    TextBox2.Value = ""
    TextBox3.Value = ""
    TextBox4.Value = ""

    Label1.Caption = ""
    Label2.Caption = ""
    Label3.Caption = ""
    Label4.Caption = ""
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub

Private Sub Send_Click()

    With Sheet3.Range("A4")
        .Value = "A"
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
    End With

    With Sheet3.Range("B4")
        .Value = "B"
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
    End With

    With Sheet3.Range("C4")
        .Value = "C"
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
    End With

    With Range("A5:C8")
        .HorizontalAlignment = xlCenter
    End With

    Sheet3.Range("A5") = TextBox1.Value
    Sheet3.Range("B5") = TextBox2.Value
    Sheet3.Range("C5") = Label1.Caption

    Sheet3.Range("A6") = TextBox3.Value
    Sheet3.Range("B6") = TextBox4.Value
    Sheet3.Range("C6") = Label2.Caption
    Sheet3.Range("A7") = Label3.Caption
    Sheet3.Range("C7") = Label4.Caption

    If OptionButton1 = True Then
        MsgBox ("ok")
    End If

End Sub

Private Sub TextBox1_AfterUpdate()

    Label1.Caption = Val(TextBox1) + Val(TextBox2)
    Label2.Caption = Val(TextBox3) + Val(TextBox4)
    Label3.Caption = Val(TextBox1) + Val(TextBox3)
    Label4.Caption = Val(Label1) + Val(Label2)

End Sub

Private Sub TextBox2_AfterUpdate()

    Label1.Caption = Val(TextBox1) + Val(TextBox2)
    Label2.Caption = Val(TextBox3) + Val(TextBox4)
    Label3.Caption = Val(TextBox1) + Val(TextBox3)
    Label4.Caption = Val(Label1) + Val(Label2)

End Sub

Private Sub TextBox3_AfterUpdate()

    Label1.Caption = Val(TextBox1) + Val(TextBox2)
    Label2.Caption = Val(TextBox3) + Val(TextBox4)
    Label3.Caption = Val(TextBox1) + Val(TextBox3)
    Label4.Caption = Val(Label1) + Val(Label2)

End Sub

Private Sub TextBox4_AfterUpdate()

    Label1.Caption = Val(TextBox1) + Val(TextBox2)
    Label2.Caption = Val(TextBox3) + Val(TextBox4)
    Label3.Caption = Val(TextBox1) + Val(TextBox3)
    Label4.Caption = Val(Label1) + Val(Label2)

End Sub
