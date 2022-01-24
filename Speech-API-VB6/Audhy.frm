VERSION 5.00
Begin VB.Form Audhy 
   BackColor       =   &H000000FF&
   Caption         =   "Audhy"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17460
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   17460
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Audhy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Sub Form_Load()
'===================================================================================================='
'BARIS KHUSUS'
'===================================================================================================='
'################################################################################'
'Parent'
bg = CreateRoundRectRgn(10, 30, 1200, 650, 0, 0) 'Background'
CombineRgn bg, bg, bg, 4

'================================================= A ==========================================
'lv1 = CreateRoundRectRgn(70, 80, 75, 0, 0, 0)
a1 = CreateRoundRectRgn(30, 30, 50, 40, 10, 5)
a2 = CreateRoundRectRgn(29, 35, 51, 45, 10, 5)
a3 = CreateRoundRectRgn(28, 40, 39, 50, 10, 5)
a4 = CreateRoundRectRgn(27, 45, 38, 55, 10, 5)
a5 = CreateRoundRectRgn(26, 50, 37, 60, 10, 5)
a6 = CreateRoundRectRgn(25, 55, 36, 65, 10, 5)
a7 = CreateRoundRectRgn(24, 60, 56, 70, 10, 5)
a8 = CreateRoundRectRgn(42, 40, 52, 50, 10, 5)
a9 = CreateRoundRectRgn(43, 45, 53, 55, 10, 5)
a10 = CreateRoundRectRgn(44, 50, 54, 60, 10, 5)
a11 = CreateRoundRectRgn(45, 55, 55, 75, 10, 5)
a12 = CreateRoundRectRgn(23, 65, 34, 75, 10, 5)
a13 = CreateRoundRectRgn(22, 70, 33, 80, 10, 5)
a14 = CreateRoundRectRgn(19, 75, 37, 82, 5, 5)
a15 = CreateRoundRectRgn(47, 65, 57, 75, 10, 5)
a16 = CreateRoundRectRgn(46, 70, 58, 80, 10, 5)
a17 = CreateRoundRectRgn(43, 75, 61, 82, 5, 5)

CombineRgn bg, bg, lv1, 2
CombineRgn bg, bg, a1, 2
CombineRgn bg, bg, a2, 2
CombineRgn bg, bg, a3, 2
CombineRgn bg, bg, a4, 2
CombineRgn bg, bg, a5, 2
CombineRgn bg, bg, a6, 2
CombineRgn bg, bg, a7, 2
CombineRgn bg, bg, a8, 2
CombineRgn bg, bg, a9, 2
CombineRgn bg, bg, a10, 2
CombineRgn bg, bg, a11, 2
CombineRgn bg, bg, a12, 2
CombineRgn bg, bg, a13, 2
CombineRgn bg, bg, a14, 2
CombineRgn bg, bg, a15, 2
CombineRgn bg, bg, a16, 2
CombineRgn bg, bg, a17, 2

'================================================= U ==========================================
u1 = CreateRoundRectRgn(92, 30, 115, 40, 5, 5)
u2 = CreateRoundRectRgn(97, 30, 110, 73, 10, 5)

u1a = CreateRoundRectRgn(119, 30, 142, 40, 5, 5)
u2a = CreateRoundRectRgn(124, 30, 137, 73, 10, 5)

u3 = CreateRoundRectRgn(97, 65, 111, 73, 10, 5)
u4 = CreateRoundRectRgn(98, 66, 112, 75, 10, 5)
u5 = CreateRoundRectRgn(99, 67, 113, 76, 10, 5)
u6 = CreateRoundRectRgn(100, 68, 114, 78, 10, 5)
u7 = CreateRoundRectRgn(101, 70, 117, 80, 10, 5)

u8 = CreateRoundRectRgn(132, 65, 137, 73, 10, 5)
u9 = CreateRoundRectRgn(131, 66, 136, 75, 10, 5)
u10 = CreateRoundRectRgn(130, 67, 135, 76, 10, 5)
u11 = CreateRoundRectRgn(119, 68, 134, 78, 10, 5)
u12 = CreateRoundRectRgn(108, 70, 133, 80, 10, 5)


CombineRgn bg, bg, u1, 2
CombineRgn bg, bg, u2, 2
CombineRgn bg, bg, u1a, 2
CombineRgn bg, bg, u2a, 2
CombineRgn bg, bg, u3, 2
CombineRgn bg, bg, u4, 2
CombineRgn bg, bg, u5, 2
CombineRgn bg, bg, u6, 2
CombineRgn bg, bg, u7, 2
CombineRgn bg, bg, u8, 2
CombineRgn bg, bg, u9, 2
CombineRgn bg, bg, u10, 2
CombineRgn bg, bg, u11, 2
CombineRgn bg, bg, u12, 2


'================================================= D ==========================================
d1 = CreateRoundRectRgn(175, 30, 208, 40, 5, 5)
d2 = CreateRoundRectRgn(177, 35, 193, 80, 10, 5)
d3 = CreateRoundRectRgn(175, 80, 208, 70, 5, 5)

d4 = CreateRoundRectRgn(190, 32, 202, 42, 10, 5)
d5 = CreateRoundRectRgn(190, 76, 209, 66, 10, 5)

d6 = CreateRoundRectRgn(195, 34, 210, 44, 10, 5)
d7 = CreateRoundRectRgn(196, 36, 211, 46, 10, 5)
d8 = CreateRoundRectRgn(197, 38, 212, 48, 10, 5)
d9 = CreateRoundRectRgn(198, 40, 213, 50, 10, 5)
d10 = CreateRoundRectRgn(199, 42, 214, 52, 10, 5)
d11 = CreateRoundRectRgn(200, 44, 215, 54, 10, 5)
d12 = CreateRoundRectRgn(201, 46, 216, 56, 10, 5)
d13 = CreateRoundRectRgn(202, 48, 216, 58, 10, 5)
d14 = CreateRoundRectRgn(202, 50, 216, 60, 10, 5)
d15 = CreateRoundRectRgn(201, 52, 216, 62, 10, 5)
d16 = CreateRoundRectRgn(200, 54, 216, 64, 10, 5)
d17 = CreateRoundRectRgn(199, 56, 215, 66, 10, 5)
d18 = CreateRoundRectRgn(198, 58, 214, 68, 10, 5)
d19 = CreateRoundRectRgn(197, 60, 213, 70, 10, 5)
d20 = CreateRoundRectRgn(196, 62, 212, 72, 10, 5)
d21 = CreateRoundRectRgn(195, 64, 211, 74, 10, 5)
d22 = CreateRoundRectRgn(194, 66, 210, 76, 10, 5)

CombineRgn bg, bg, d1, 2
CombineRgn bg, bg, d2, 2
CombineRgn bg, bg, d3, 2
CombineRgn bg, bg, d4, 2
CombineRgn bg, bg, d5, 2
CombineRgn bg, bg, d6, 2
CombineRgn bg, bg, d7, 2
CombineRgn bg, bg, d8, 2
CombineRgn bg, bg, d9, 2
CombineRgn bg, bg, d11, 2
CombineRgn bg, bg, d12, 2
CombineRgn bg, bg, d13, 2
CombineRgn bg, bg, d14, 2
CombineRgn bg, bg, d15, 2
CombineRgn bg, bg, d16, 2
CombineRgn bg, bg, d17, 2
CombineRgn bg, bg, d18, 2
CombineRgn bg, bg, d19, 2
CombineRgn bg, bg, d20, 2
CombineRgn bg, bg, d21, 2
CombineRgn bg, bg, d22, 2

'================================================= H ==========================================
h1 = CreateRoundRectRgn(252, 30, 275, 40, 5, 5)
h2 = CreateRoundRectRgn(256, 30, 271, 80, 10, 5)
h3 = CreateRoundRectRgn(252, 80, 275, 70, 5, 5)

h4 = CreateRoundRectRgn(258, 50, 297, 60, 10, 5)

h5 = CreateRoundRectRgn(280, 30, 303, 40, 5, 5)
h6 = CreateRoundRectRgn(284, 30, 299, 80, 10, 5)
h7 = CreateRoundRectRgn(280, 80, 303, 70, 5, 5)

CombineRgn bg, bg, h1, 2
CombineRgn bg, bg, h2, 2
CombineRgn bg, bg, h3, 2
CombineRgn bg, bg, h4, 2
CombineRgn bg, bg, h5, 2
CombineRgn bg, bg, h6, 2
CombineRgn bg, bg, h7, 2

'================================================= Y ==========================================
Y1 = CreateRoundRectRgn(330, 30, 350, 40, 5, 5)
Y2 = CreateRoundRectRgn(359, 30, 379, 40, 5, 5)

Y3 = CreateRoundRectRgn(338, 30, 350, 40, 10, 5)
Y4 = CreateRoundRectRgn(339, 32, 351, 42, 10, 5)
Y5 = CreateRoundRectRgn(340, 34, 352, 44, 10, 5)
Y6 = CreateRoundRectRgn(341, 36, 353, 46, 10, 5)
Y7 = CreateRoundRectRgn(342, 38, 354, 48, 10, 5)
Y8 = CreateRoundRectRgn(343, 40, 355, 50, 10, 5)
Y9 = CreateRoundRectRgn(344, 42, 356, 52, 10, 5)
Y10 = CreateRoundRectRgn(345, 44, 357, 54, 10, 5)
Y11 = CreateRoundRectRgn(346, 46, 358, 56, 10, 5)
Y12 = CreateRoundRectRgn(347, 48, 359, 58, 10, 5)

Y13 = CreateRoundRectRgn(359, 30, 371, 40, 10, 5)
Y14 = CreateRoundRectRgn(358, 32, 370, 42, 10, 5)
Y15 = CreateRoundRectRgn(357, 34, 369, 44, 10, 5)
Y16 = CreateRoundRectRgn(356, 36, 368, 46, 10, 5)
Y17 = CreateRoundRectRgn(355, 38, 367, 48, 10, 5)
Y18 = CreateRoundRectRgn(354, 40, 366, 50, 10, 5)
Y19 = CreateRoundRectRgn(353, 42, 365, 52, 10, 5)
Y20 = CreateRoundRectRgn(348, 44, 364, 54, 10, 5)
Y21 = CreateRoundRectRgn(352, 46, 363, 56, 10, 5)
Y22 = CreateRoundRectRgn(351, 48, 362, 58, 10, 5)

Y23 = CreateRoundRectRgn(349, 50, 361, 80, 5, 5)

CombineRgn bg, bg, Y1, 2
CombineRgn bg, bg, Y2, 2
CombineRgn bg, bg, Y3, 2
CombineRgn bg, bg, Y4, 2
CombineRgn bg, bg, Y5, 2
CombineRgn bg, bg, Y6, 2
CombineRgn bg, bg, Y7, 2
CombineRgn bg, bg, Y8, 2
CombineRgn bg, bg, Y9, 2
CombineRgn bg, bg, Y10, 2
CombineRgn bg, bg, Y11, 2
CombineRgn bg, bg, Y12, 2
CombineRgn bg, bg, Y13, 2
CombineRgn bg, bg, Y14, 2
CombineRgn bg, bg, Y15, 2
CombineRgn bg, bg, Y16, 2
CombineRgn bg, bg, Y17, 2
CombineRgn bg, bg, Y18, 2
CombineRgn bg, bg, Y19, 2
CombineRgn bg, bg, Y20, 2
CombineRgn bg, bg, Y21, 2
CombineRgn bg, bg, Y22, 2
CombineRgn bg, bg, Y23, 2

'View Hasil Combine'
SetWindowRgn Audhy.hwnd, bg, True
End Sub
'Responsive Object -> Mouse Trigger'
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ReleaseCapture
    SendMessage Audhy.hwnd, &HA1, 2, 0&
End Sub







