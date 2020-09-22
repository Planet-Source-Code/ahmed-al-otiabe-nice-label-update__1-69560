VERSION 5.00
Begin VB.Form FrmAbout 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const mStr1 = "NiceLabel"

Private Const mStr2 = "Nice Label Is More Than 7 BackStyle's , Trans + Gradinet _ "

Private Const mStr3 = "Fill Horizontal Or Vertical ... etc"

Private Const mStr4 = "And Moer Than 7 Text Effect's , Raised , Sunken , Outline _ "
Private Const mStr5 = "Shadow , 3DRaisedShadow ...etc"

Private Const mStr6 = "And More Than 7 Text Style's - GradinetFill Horizontal Or _ "
Private Const mStr7 = "Vertical And Horizontal Vertical Center + Out And More ... "

Private Const mStr8 = "CopyRight © 2007-2008  AHMED AL OTAIBE   Version 1.00"
Private Const mStr9 = "Thank You For Using This Label."

Private Const mStr10 = "Email: x19@hotmail.com"
Dim X, Y                 As Integer
Private Sub Form_Load()
    With FrmAbout
        .ScaleMode = 3
        .BackColor = &HE6EAEB
        .Width = 5000
        .Height = 0
        '---------------------------------------------------------------------'
        ' Postion FrmAbout
        '---------------------------------------------------------------------'
        .Show
        For Y = 0 To 4000 Step 24
            .Cls
            Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
            .Height = Y
        Next Y
    End With
    Call Form_Paint
End Sub
Private Sub Form_Paint()
    On Local Error Resume Next
    With FrmAbout
        .Cls
        '---------------------------------------------------------------------'
        'Draw Form BorderStyle
        '---------------------------------------------------------------------'
        Line (0, 0)-(.ScaleWidth - 1, .ScaleHeight - 1), vb3DDKShadow, B
        Line (0, 0)-(.ScaleWidth - 2, .ScaleHeight - 2), vbButtonFace, B
        Line (0, 0)-(.ScaleWidth, .ScaleHeight), vb3DHighlight, B

        '---------------------------------------------------------------------'
        'Draw Frame  BorderStyle
        '---------------------------------------------------------------------'
        Line (33, 75)-(.ScaleWidth - 9, .ScaleHeight - 45), &HF4F7F7, BF
        Line (33, 75)-(.ScaleWidth - 9, .ScaleHeight - 45), vbButtonShadow, B
        Line (.ScaleWidth - 9, 75)-(.ScaleWidth - 9, .ScaleHeight - 45), vb3DHighlight
        Line (33, .ScaleHeight - 45)-(.ScaleWidth - 8, .ScaleHeight - 45), vb3DHighlight

        '---------------------------------------------------------------------'
        'Print Inforamtions
        '---------------------------------------------------------------------'
        .Font = "Tahoma": .FontBold = False: .FontItalic = False: .FontSize = 8
        .ForeColor = vbWhite
        .CurrentX = .ScaleWidth - 295: .CurrentY = .ScaleHeight - 185
        Print mStr2
        .ForeColor = vbBlack
        .CurrentX = .ScaleWidth - 296: .CurrentY = .ScaleHeight - 186
        Print mStr2
        '---------------------------------------------------------------------'
        .ForeColor = vbWhite
        .CurrentX = .ScaleWidth - 295: .CurrentY = .ScaleHeight - 170
        Print mStr3
        .ForeColor = vbBlack
        .CurrentX = .ScaleWidth - 296: .CurrentY = .ScaleHeight - 171
        Print mStr3
        '---------------------------------------------------------------------'
        .ForeColor = vbWhite
        .CurrentX = .ScaleWidth - 295: .CurrentY = .ScaleHeight - 150
        Print mStr4
        .ForeColor = vbBlack
        .CurrentX = .ScaleWidth - 296: .CurrentY = .ScaleHeight - 151
        Print mStr4
        '---------------------------------------------------------------------'
        .ForeColor = vbWhite
        .CurrentX = .ScaleWidth - 295: .CurrentY = .ScaleHeight - 135
        Print mStr5
        .ForeColor = vbBlack
        .CurrentX = .ScaleWidth - 296: .CurrentY = .ScaleHeight - 136
        Print mStr5
        '---------------------------------------------------------------------'
        .ForeColor = vbWhite
        .CurrentX = .ScaleWidth - 295: .CurrentY = .ScaleHeight - 115
        Print mStr6
        .ForeColor = vbBlack
        .CurrentX = .ScaleWidth - 296: .CurrentY = .ScaleHeight - 116
        Print mStr6
        '---------------------------------------------------------------------'
        .ForeColor = vbWhite
        .CurrentX = .ScaleWidth - 295: .CurrentY = .ScaleHeight - 100
        Print mStr7
        .ForeColor = vbBlack
        .CurrentX = .ScaleWidth - 296: .CurrentY = .ScaleHeight - 101
        Print mStr7
        '---------------------------------------------------------------------'
        .ForeColor = vbWhite
        .CurrentX = .ScaleWidth - 295: .CurrentY = .ScaleHeight - 80
        Print mStr8
        .ForeColor = vbBlack
        .CurrentX = .ScaleWidth - 296: .CurrentY = .ScaleHeight - 81
        Print mStr8
        '---------------------------------------------------------------------'
        .ForeColor = vbWhite
        .CurrentX = .ScaleWidth - 240: .CurrentY = .ScaleHeight - 60
        Print mStr9
        .ForeColor = &H3536FF
        .CurrentX = .ScaleWidth - 241: .CurrentY = .ScaleHeight - 61
        Print mStr9

        '---------------------------------------------------------------------'
        'Print Big Gradient Text
        '---------------------------------------------------------------------'
        .Font = "Times New Roman": .FontBold = True: .FontItalic = False: .FontSize = 50
        .ForeColor = vbRed
        .CurrentX = .ScaleWidth - 301: .CurrentY = .ScaleHeight - 261
        Print mStr1
        .ForeColor = vbWhite
        .CurrentX = .ScaleWidth - 300: .CurrentY = .ScaleHeight - 260
        Print mStr1
        For X = 30 To .ScaleWidth - 10
            For Y = 20 To 70
                If Point(X, Y) = .ForeColor Then PSet (X, Y), RGB(25 * Y, 5 * Y, 5 * Y)
            Next: Next

        '---------------------------------------------------------------------'
        'Draw GradientVertical
        '---------------------------------------------------------------------'
        For X = 3 To .ScaleWidth - 305
            Line (X, 3)-(.ScaleWidth - 305, .ScaleHeight - 4), RGB(75 * X, 10 * X, 10 * X), BF
        Next X

        '---------------------------------------------------------------------'
        'Simple Logo
        '---------------------------------------------------------------------'
        X = .ScaleWidth - 314

        '---------------------------------------------------------------------'
        Y = .ScaleHeight - 250
        Line (X - 0.5, Y - 8.5)-(X - 6.5, Y - 2.5), vbWhite
        Line (X - 9.5, Y - 8.5)-(X - 3.5, Y - 2.5), vbWhite
        '---------------------------------------------------------------------'
        Y = .ScaleHeight - 248
        Line (X - 0.5, Y - 8.5)-(X - 6.5, Y - 2.5), vbWhite
        Line (X - 9.5, Y - 8.5)-(X - 3.5, Y - 2.5), vbWhite
        '---------------------------------------------------------------------'
        Y = .ScaleHeight - 246
        Line (X - 0.5, Y - 8.5)-(X - 6.5, Y - 2.5), vbWhite
        Line (X - 9.5, Y - 8.5)-(X - 3.5, Y - 2.5), vbWhite
        '---------------------------------------------------------------------'
        Y = .ScaleHeight - 243
        Line (X - 0.5, Y - 8.5)-(X - 6.5, Y - 2.5), vbWhite
        Line (X - 9.5, Y - 8.5)-(X - 3.5, Y - 2.5), vbWhite
        '---------------------------------------------------------------------'
        Y = .ScaleHeight - 240
        Line (X - 0.5, Y - 8.5)-(X - 6.5, Y - 2.5), vbWhite
        Line (X - 9.5, Y - 8.5)-(X - 3.5, Y - 2.5), vbWhite
        '---------------------------------------------------------------------'
        Y = .ScaleHeight - 237
        Line (X - 0.5, Y - 8.5)-(X - 6.5, Y - 2.5), vbWhite
        Line (X - 9.5, Y - 8.5)-(X - 3.5, Y - 2.5), vbWhite
        '---------------------------------------------------------------------'
        Y = .ScaleHeight - 234
        Line (X - 0.5, Y - 8.5)-(X - 6.5, Y - 2.5), vbWhite
        Line (X - 9.5, Y - 8.5)-(X - 3.5, Y - 2.5), vbWhite
        '---------------------------------------------------------------------'
        Y = .ScaleHeight - 231
        Line (X - 0.5, Y - 8.5)-(X - 6.5, Y - 2.5), vbWhite
        Line (X - 9.5, Y - 8.5)-(X - 3.5, Y - 2.5), vbWhite
        '---------------------------------------------------------------------'
        Y = .ScaleHeight - 228
        Line (X - 0.5, Y - 8.5)-(X - 6.5, Y - 2.5), vbWhite
        Line (X - 9.5, Y - 8.5)-(X - 3.5, Y - 2.5), vbWhite
        '---------------------------------------------------------------------'
        Y = .ScaleHeight - 226
        Line (X - 0.5, Y - 8.5)-(X - 6.5, Y - 2.5), vbWhite
        Line (X - 9.5, Y - 8.5)-(X - 3.5, Y - 2.5), vbWhite

        '---------------------------------------------------------------------'
        'Print VerticalWord's
        '---------------------------------------------------------------------'
        .Font = "Times New Roman": .FontBold = True: .FontItalic = False: .FontSize = 13

        .ForeColor = vbWhite
        .CurrentY = .ScaleHeight - 151: .CurrentX = .ScaleWidth - 323
        Print "A"
        .ForeColor = vbRed
        .CurrentY = .ScaleHeight - 150: .CurrentX = .ScaleWidth - 322
        Print "A"
        '---------------------------------------------------------------------'
        .ForeColor = vbWhite
        .CurrentY = .ScaleHeight - 131: .CurrentX = .ScaleWidth - 324
        Print "C"
        .ForeColor = vbRed
        .CurrentY = .ScaleHeight - 130: .CurrentX = .ScaleWidth - 323
        Print "C"
        '---------------------------------------------------------------------'
        .ForeColor = vbWhite
        .CurrentY = .ScaleHeight - 111: .CurrentX = .ScaleWidth - 323
        Print "T"
        .ForeColor = vbRed
        .CurrentY = .ScaleHeight - 110: .CurrentX = .ScaleWidth - 322
        Print "T"
        '---------------------------------------------------------------------'
        .ForeColor = vbWhite
        .CurrentY = .ScaleHeight - 91: .CurrentX = .ScaleWidth - 321
        Print "I"
        .ForeColor = vbRed
        .CurrentY = .ScaleHeight - 90: .CurrentX = .ScaleWidth - 320
        Print "I"
        '---------------------------------------------------------------------'
        .ForeColor = vbWhite
        .CurrentY = .ScaleHeight - 71: .CurrentX = .ScaleWidth - 324
        Print "V"
        .ForeColor = vbRed
        .CurrentY = .ScaleHeight - 70: .CurrentX = .ScaleWidth - 323
        Print "V"
        '---------------------------------------------------------------------'
        .ForeColor = vbWhite
        .CurrentY = .ScaleHeight - 51: .CurrentX = .ScaleWidth - 323
        Print "E"
        .ForeColor = vbRed
        .CurrentY = .ScaleHeight - 50: .CurrentX = .ScaleWidth - 322
        Print "E"
        '---------------------------------------------------------------------'
        .ForeColor = vbWhite
        .CurrentY = .ScaleHeight - 31: .CurrentX = .ScaleWidth - 323
        Print "X"
        .ForeColor = vbBlack
        .CurrentY = .ScaleHeight - 30: .CurrentX = .ScaleWidth - 322
        Print "X"

        '---------------------------------------------------------------------'
        'Draw Exit Button BorderStyle >> Mouse Outside
        '---------------------------------------------------------------------'
        Line (.ScaleWidth - 10, .ScaleHeight - 10)-(.ScaleWidth - 80, .ScaleHeight - 35), &HF1F1F1, BF
        Line (.ScaleWidth - 10, .ScaleHeight - 10)-(.ScaleWidth - 55, .ScaleHeight - 35), &HE6EAEB, BF
        Line (.ScaleWidth - 10, .ScaleHeight - 10)-(.ScaleWidth - 80, .ScaleHeight - 35), vbWhite, B

        '---------------------------------------------------------------------'
        'Draw X Flag
        '---------------------------------------------------------------------'
        Line (.ScaleWidth - 73, .ScaleHeight - 27)-(.ScaleWidth - 63, .ScaleHeight - 17), vbRed
        Line (.ScaleWidth - 64, .ScaleHeight - 27)-(.ScaleWidth - 74, .ScaleHeight - 17), vbRed
        Line (.ScaleWidth - 72, .ScaleHeight - 27)-(.ScaleWidth - 62, .ScaleHeight - 17), vbRed
        Line (.ScaleWidth - 63, .ScaleHeight - 27)-(.ScaleWidth - 73, .ScaleHeight - 17), vbRed

        '---------------------------------------------------------------------'
        'Print Text
        '---------------------------------------------------------------------'
        .Font = "Tahoma": .FontBold = False: .FontItalic = False: .FontSize = 10
        .CurrentX = .ScaleWidth - 42: .CurrentY = .ScaleHeight - 30
        .ForeColor = vbBlack
        Print "Exit"
        '---------------------------------------------------------------------'
        'Print Email Text >> Mouse Outside
        '---------------------------------------------------------------------'
        .Font = "MS Sans Serif": .FontBold = False: .FontItalic = False: .FontSize = 10
        .CurrentX = .ScaleWidth - 295: .CurrentY = .ScaleHeight - 30
        Print mStr10
        .MousePointer = 0
    End With
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With FrmAbout
        If X < .ScaleWidth - 10 And Y < .ScaleHeight - 10 And X > .ScaleWidth - 80 And Y > .ScaleHeight - 35 Then

            '---------------------------------------------------------------------'
            'Draw Exit Button BorderStyle
            '---------------------------------------------------------------------'
            Line (.ScaleWidth - 10, .ScaleHeight - 10)-(.ScaleWidth - 80, .ScaleHeight - 35), vbWhite, BF
            Line (.ScaleWidth - 10, .ScaleHeight - 10)-(.ScaleWidth - 55, .ScaleHeight - 35), &HB59285, BF
            Line (.ScaleWidth - 10, .ScaleHeight - 10)-(.ScaleWidth - 80, .ScaleHeight - 35), vbBlack, B

            '---------------------------------------------------------------------'
            'Draw X Flag
            '---------------------------------------------------------------------'
            Line (.ScaleWidth - 72, .ScaleHeight - 24)-(.ScaleWidth - 62, .ScaleHeight - 14), vbRed
            Line (.ScaleWidth - 62, .ScaleHeight - 24)-(.ScaleWidth - 72, .ScaleHeight - 14), vbRed
            Line (.ScaleWidth - 71, .ScaleHeight - 24)-(.ScaleWidth - 61, .ScaleHeight - 14), vbRed
            Line (.ScaleWidth - 61, .ScaleHeight - 24)-(.ScaleWidth - 71, .ScaleHeight - 14), vbRed

            '---------------------------------------------------------------------'
            'Print Text
            '---------------------------------------------------------------------'
            .Font = "Tahoma": .FontBold = False: .FontItalic = False: .FontSize = 10
            .CurrentX = .ScaleWidth - 40: .CurrentY = .ScaleHeight - 28
            .ForeColor = vbWhite
            Print "Exit"
        End If
    End With
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With FrmAbout
        '---------------------------------------------------------------------'
        'Draw Exit Button  BorderStyle >> Mouse Inside
        '---------------------------------------------------------------------'
        If X < .ScaleWidth - 10 And Y < .ScaleHeight - 10 And X > .ScaleWidth - 80 And Y > .ScaleHeight - 35 Then
            Line (.ScaleWidth - 10, .ScaleHeight - 10)-(.ScaleWidth - 80, .ScaleHeight - 35), vbWhite, BF
            Line (.ScaleWidth - 10, .ScaleHeight - 10)-(.ScaleWidth - 55, .ScaleHeight - 35), &HD2BDB6, BF
            Line (.ScaleWidth - 10, .ScaleHeight - 10)-(.ScaleWidth - 80, .ScaleHeight - 35), vbBlack, B

            '---------------------------------------------------------------------'
            'Draw Shadow X Flag
            '---------------------------------------------------------------------'
            Line (.ScaleWidth - 72, .ScaleHeight - 24)-(.ScaleWidth - 62, .ScaleHeight - 14), vbBlack
            Line (.ScaleWidth - 62, .ScaleHeight - 24)-(.ScaleWidth - 72, .ScaleHeight - 14), vbBlack

            '---------------------------------------------------------------------'
            'Draw X Flag
            '---------------------------------------------------------------------'
            Line (.ScaleWidth - 73, .ScaleHeight - 27)-(.ScaleWidth - 63, .ScaleHeight - 17), vbRed
            Line (.ScaleWidth - 64, .ScaleHeight - 27)-(.ScaleWidth - 74, .ScaleHeight - 17), vbRed
            Line (.ScaleWidth - 72, .ScaleHeight - 27)-(.ScaleWidth - 62, .ScaleHeight - 17), vbRed
            Line (.ScaleWidth - 63, .ScaleHeight - 27)-(.ScaleWidth - 73, .ScaleHeight - 17), vbRed

            '---------------------------------------------------------------------'
            'Print Text
            '---------------------------------------------------------------------'
            .Font = "Tahoma": .FontBold = False: .FontItalic = False: .FontSize = 10
            .CurrentX = .ScaleWidth - 42: .CurrentY = .ScaleHeight - 30
            .ForeColor = vbBlack
            Print "Exit"
        Else

            '---------------------------------------------------------------------'
            'Draw Exit Button BorderStyle >> Mouse Outside
            '---------------------------------------------------------------------'
            Line (.ScaleWidth - 10, .ScaleHeight - 10)-(.ScaleWidth - 80, .ScaleHeight - 35), &HF1F1F1, BF
            Line (.ScaleWidth - 10, .ScaleHeight - 10)-(.ScaleWidth - 55, .ScaleHeight - 35), &HE6EAEB, BF
            Line (.ScaleWidth - 10, .ScaleHeight - 10)-(.ScaleWidth - 80, .ScaleHeight - 35), vbWhite, B

            '---------------------------------------------------------------------'
            'Draw X Flag
            '---------------------------------------------------------------------'
            Line (.ScaleWidth - 73, .ScaleHeight - 27)-(.ScaleWidth - 63, .ScaleHeight - 17), vbRed
            Line (.ScaleWidth - 64, .ScaleHeight - 27)-(.ScaleWidth - 74, .ScaleHeight - 17), vbRed
            Line (.ScaleWidth - 72, .ScaleHeight - 27)-(.ScaleWidth - 62, .ScaleHeight - 17), vbRed
            Line (.ScaleWidth - 63, .ScaleHeight - 27)-(.ScaleWidth - 73, .ScaleHeight - 17), vbRed

            '---------------------------------------------------------------------'
            'Print Text
            '---------------------------------------------------------------------'
            .Font = "Tahoma": .FontBold = False: .FontItalic = False: .FontSize = 10
            .CurrentX = .ScaleWidth - 42: .CurrentY = .ScaleHeight - 30
            .ForeColor = vbBlack
            Print "Exit"

            '---------------------------------------------------------------------'
            'Print Email Text >> Mouse Outside
            '---------------------------------------------------------------------'
            Line (.ScaleWidth - 150, .ScaleHeight - 15)-(.ScaleWidth - 296, .ScaleHeight - 15), .BackColor
            .Font = "MS Sans Serif": .FontBold = False: .FontItalic = False: .FontSize = 10
            .CurrentX = .ScaleWidth - 295: .CurrentY = .ScaleHeight - 30
            Print mStr10
            .MousePointer = 0
        End If

        '---------------------------------------------------------------------'
        '---------------------------------------------------------------------'
        If X < .ScaleWidth - 150 And Y < .ScaleHeight - 17 And X > .ScaleWidth - 296 And Y > .ScaleHeight - 27 Then
            '---------------------------------------------------------------------'
            'Print Email Text >> Mouse Inside
            '---------------------------------------------------------------------'
            Line (.ScaleWidth - 150, .ScaleHeight - 15)-(.ScaleWidth - 296, .ScaleHeight - 15), vbBlue
            .Font = "MS Sans Serif": .FontBold = False: .FontItalic = False: .FontSize = 10
            .CurrentX = .ScaleWidth - 295: .CurrentY = .ScaleHeight - 30
            .ForeColor = vbBlue
            Print mStr10
            On Error Resume Next
            .MousePointer = 99    'Custom
            Set .MouseIcon = LoadPicture("Hand.CUR")
        End If
    End With
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With FrmAbout
        If X < .ScaleWidth - 150 And Y < .ScaleHeight - 17 And X > .ScaleWidth - 296 And Y > .ScaleHeight - 27 Then Shell "Explorer.Exe MailTo:x19@hotmail.com", vbNormalFocus
        If X < .ScaleWidth - 10 And Y < .ScaleHeight - 10 And X > .ScaleWidth - 80 And Y > .ScaleHeight - 35 Then
            For Y = 0 To 4000
                .Cls
                DoEvents
                Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
                .Height = .Height - .Height / 24
                '.Width = .Width - .Width / 24
            Next Y
            .BackColor = vbBlack
            For X = 0 To 5000
                .Cls
                DoEvents
                Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
                .Width = .Width + .Width / 24
            Next X
            Unload Me
        End If
    End With
End Sub
