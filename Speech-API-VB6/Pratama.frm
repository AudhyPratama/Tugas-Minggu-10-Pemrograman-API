VERSION 5.00
Begin VB.Form Pratama 
   BackColor       =   &H000000FF&
   Caption         =   "Pratama"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17505
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   17505
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Pratama"
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

'================================================= P ==========================================
p1 = CreateRoundRectRgn(200, 120, 230, 130, 5, 5)
p2 = CreateRoundRectRgn(200, 170, 220, 160, 5, 5)
p3 = CreateRoundRectRgn(205, 120, 215, 170, 10, 15)

p4 = CreateRoundRectRgn(224, 122, 232, 135, 10, 5)
p5 = CreateRoundRectRgn(225, 125, 233, 137, 10, 5)
p6 = CreateRoundRectRgn(225, 128, 234, 141, 10, 5)
p7 = CreateRoundRectRgn(225, 131, 234, 141, 10, 5)
p7a = CreateRoundRectRgn(225, 134, 234, 143, 10, 5)
p7b = CreateRoundRectRgn(224, 137, 233, 145, 10, 5)
p8 = CreateRoundRectRgn(223, 140, 232, 147, 10, 5)
p9 = CreateRoundRectRgn(222, 143, 231, 149, 10, 5)
p10 = CreateRoundRectRgn(221, 146, 230, 151, 10, 5)
p11 = CreateRoundRectRgn(205, 143, 229, 151, 5, 5)

CombineRgn bg, bg, p1, 2
CombineRgn bg, bg, p2, 2
CombineRgn bg, bg, p3, 2
CombineRgn bg, bg, p4, 2
CombineRgn bg, bg, p5, 2
CombineRgn bg, bg, p6, 2
CombineRgn bg, bg, p7, 2
CombineRgn bg, bg, p7a, 2
CombineRgn bg, bg, p7b, 2
CombineRgn bg, bg, p8, 2
CombineRgn bg, bg, p9, 2
CombineRgn bg, bg, p10, 2
CombineRgn bg, bg, p11, 2

'================================================= R ==========================================
r1 = CreateRoundRectRgn(305, 120, 338, 130, 5, 5)
r2 = CreateRoundRectRgn(305, 170, 325, 160, 5, 5)
r3 = CreateRoundRectRgn(310, 120, 320, 170, 10, 15)

r4 = CreateRoundRectRgn(330, 122, 340, 135, 10, 5)
r5 = CreateRoundRectRgn(331, 125, 341, 137, 10, 5)
r6 = CreateRoundRectRgn(331, 128, 342, 141, 10, 5)
r7 = CreateRoundRectRgn(331, 131, 342, 141, 10, 5)
r7a = CreateRoundRectRgn(331, 134, 342, 143, 10, 5)
r7b = CreateRoundRectRgn(330, 137, 341, 145, 10, 5)
r8 = CreateRoundRectRgn(329, 140, 340, 147, 10, 5)
r9 = CreateRoundRectRgn(328, 143, 339, 149, 10, 5)
r10 = CreateRoundRectRgn(327, 146, 338, 151, 10, 5)
r11 = CreateRoundRectRgn(315, 143, 337, 151, 5, 5)

r12 = CreateRoundRectRgn(328, 145, 338, 155, 5, 5)
r13 = CreateRoundRectRgn(329, 150, 339, 160, 5, 5)
r14 = CreateRoundRectRgn(330, 155, 340, 165, 5, 5)
r15 = CreateRoundRectRgn(331, 160, 343, 170, 5, 5)

CombineRgn bg, bg, r1, 2
CombineRgn bg, bg, r2, 2
CombineRgn bg, bg, r3, 2
CombineRgn bg, bg, r4, 2
CombineRgn bg, bg, r5, 2
CombineRgn bg, bg, r6, 2
CombineRgn bg, bg, r7, 2
CombineRgn bg, bg, r7a, 2
CombineRgn bg, bg, r7b, 2
CombineRgn bg, bg, r8, 2
CombineRgn bg, bg, r9, 2
CombineRgn bg, bg, r10, 2
CombineRgn bg, bg, r11, 2
CombineRgn bg, bg, r12, 2
CombineRgn bg, bg, r13, 2
CombineRgn bg, bg, r14, 2
CombineRgn bg, bg, r15, 2
CombineRgn bg, bg, r16, 2
CombineRgn bg, bg, r17, 2

'================================================= A ==========================================
a1 = CreateRoundRectRgn(420, 120, 440, 130, 10, 5)
a2 = CreateRoundRectRgn(419, 125, 441, 135, 10, 5)
a3 = CreateRoundRectRgn(418, 130, 429, 140, 10, 5)
a4 = CreateRoundRectRgn(417, 135, 428, 145, 10, 5)
a5 = CreateRoundRectRgn(416, 140, 427, 150, 10, 5)
a6 = CreateRoundRectRgn(415, 145, 426, 155, 10, 5)
a7 = CreateRoundRectRgn(414, 150, 446, 160, 10, 5)
a8 = CreateRoundRectRgn(432, 130, 442, 140, 10, 5)
a9 = CreateRoundRectRgn(433, 135, 443, 145, 10, 5)
a10 = CreateRoundRectRgn(434, 140, 444, 150, 10, 5)
a11 = CreateRoundRectRgn(435, 145, 445, 165, 10, 5)
a12 = CreateRoundRectRgn(413, 155, 424, 165, 10, 5)
a13 = CreateRoundRectRgn(412, 160, 423, 170, 10, 5)
a14 = CreateRoundRectRgn(409, 165, 427, 172, 5, 5)
a15 = CreateRoundRectRgn(437, 155, 447, 165, 10, 5)
a16 = CreateRoundRectRgn(436, 160, 448, 170, 10, 5)
a17 = CreateRoundRectRgn(433, 165, 451, 172, 5, 5)

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

'================================================= T ==========================================
t1 = CreateRoundRectRgn(505, 120, 545, 130, 5, 5)
t2 = CreateRoundRectRgn(520, 120, 530, 170, 5, 5)
t3 = CreateRoundRectRgn(515, 162, 535, 170, 5, 5)

t4 = CreateRoundRectRgn(505, 126, 515, 132, 25, 1)
t5 = CreateRoundRectRgn(505, 127, 514, 133, 25, 1)
t6 = CreateRoundRectRgn(505, 128, 513, 134, 25, 1)

t7 = CreateRoundRectRgn(535, 126, 545, 132, 25, 1)
t8 = CreateRoundRectRgn(536, 127, 545, 133, 25, 1)
t9 = CreateRoundRectRgn(537, 128, 545, 134, 25, 1)

CombineRgn bg, bg, t1, 2
CombineRgn bg, bg, t2, 2
CombineRgn bg, bg, t3, 2
CombineRgn bg, bg, t4, 2
CombineRgn bg, bg, t5, 2
CombineRgn bg, bg, t6, 2
CombineRgn bg, bg, t7, 2
CombineRgn bg, bg, t8, 2
CombineRgn bg, bg, t9, 2

'================================================= A ==========================================
a1 = CreateRoundRectRgn(600, 120, 620, 130, 10, 5)
a2 = CreateRoundRectRgn(599, 125, 621, 135, 10, 5)
a3 = CreateRoundRectRgn(598, 130, 609, 140, 10, 5)
a4 = CreateRoundRectRgn(597, 135, 608, 145, 10, 5)
a5 = CreateRoundRectRgn(596, 140, 607, 150, 10, 5)
a6 = CreateRoundRectRgn(595, 145, 606, 155, 10, 5)
a7 = CreateRoundRectRgn(594, 150, 626, 160, 10, 5)
a8 = CreateRoundRectRgn(612, 130, 622, 140, 10, 5)
a9 = CreateRoundRectRgn(613, 135, 623, 145, 10, 5)
a10 = CreateRoundRectRgn(614, 140, 624, 150, 10, 5)
a11 = CreateRoundRectRgn(615, 145, 625, 165, 10, 5)
a12 = CreateRoundRectRgn(593, 155, 604, 165, 10, 5)
a13 = CreateRoundRectRgn(592, 160, 603, 170, 10, 5)
a14 = CreateRoundRectRgn(589, 165, 607, 172, 5, 5)
a15 = CreateRoundRectRgn(617, 155, 627, 165, 10, 5)
a16 = CreateRoundRectRgn(616, 160, 628, 170, 10, 5)
a17 = CreateRoundRectRgn(613, 165, 631, 172, 5, 5)

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

'================================================= M ==========================================
m1 = CreateRoundRectRgn(675, 120, 698, 130, 5, 5)
m2 = CreateRoundRectRgn(681, 120, 692, 170, 10, 5)
m3 = CreateRoundRectRgn(675, 172, 698, 160, 5, 5)

m1a = CreateRoundRectRgn(706, 120, 730, 130, 5, 5)
m2a = CreateRoundRectRgn(713, 120, 724, 170, 10, 5)
m3a = CreateRoundRectRgn(706, 172, 730, 160, 5, 5)

m4 = CreateRoundRectRgn(685, 125, 699, 135, 5, 5)
m5 = CreateRoundRectRgn(693, 130, 700, 140, 5, 5)
m6 = CreateRoundRectRgn(694, 135, 701, 145, 5, 5)
m7 = CreateRoundRectRgn(695, 140, 702, 150, 5, 5)
m8 = CreateRoundRectRgn(696, 145, 703, 155, 5, 5)
m9 = CreateRoundRectRgn(697, 150, 704, 160, 5, 5)

m10 = CreateRoundRectRgn(704, 125, 720, 135, 5, 5)
m11 = CreateRoundRectRgn(703, 130, 712, 140, 5, 5)
m12 = CreateRoundRectRgn(702, 135, 711, 145, 5, 5)
m13 = CreateRoundRectRgn(701, 140, 710, 150, 5, 5)
m14 = CreateRoundRectRgn(700, 145, 709, 155, 5, 5)
m15 = CreateRoundRectRgn(699, 150, 708, 160, 5, 5)

CombineRgn bg, bg, m1, 2
CombineRgn bg, bg, m2, 2
CombineRgn bg, bg, m3, 2
CombineRgn bg, bg, m1a, 2
CombineRgn bg, bg, m2a, 2
CombineRgn bg, bg, m3a, 2
CombineRgn bg, bg, m4, 2
CombineRgn bg, bg, m5, 2
CombineRgn bg, bg, m6, 2
CombineRgn bg, bg, m7, 2
CombineRgn bg, bg, m8, 2
CombineRgn bg, bg, m9, 2
CombineRgn bg, bg, m10, 2
CombineRgn bg, bg, m11, 2
CombineRgn bg, bg, m12, 2
CombineRgn bg, bg, m13, 2
CombineRgn bg, bg, m14, 2
CombineRgn bg, bg, m15, 2
CombineRgn bg, bg, m16, 2

'================================================= A ==========================================
a1 = CreateRoundRectRgn(780, 120, 800, 130, 10, 5)
a2 = CreateRoundRectRgn(779, 125, 801, 135, 10, 5)
a3 = CreateRoundRectRgn(778, 130, 789, 140, 10, 5)
a4 = CreateRoundRectRgn(777, 135, 788, 145, 10, 5)
a5 = CreateRoundRectRgn(776, 140, 787, 150, 10, 5)
a6 = CreateRoundRectRgn(775, 145, 786, 155, 10, 5)
a7 = CreateRoundRectRgn(774, 150, 806, 160, 10, 5)
a8 = CreateRoundRectRgn(792, 130, 802, 140, 10, 5)
a9 = CreateRoundRectRgn(793, 135, 803, 145, 10, 5)
a10 = CreateRoundRectRgn(794, 140, 804, 150, 10, 5)
a11 = CreateRoundRectRgn(795, 145, 805, 165, 10, 5)
a12 = CreateRoundRectRgn(773, 155, 784, 165, 10, 5)
a13 = CreateRoundRectRgn(772, 160, 783, 170, 10, 5)
a14 = CreateRoundRectRgn(769, 165, 787, 172, 5, 5)
a15 = CreateRoundRectRgn(797, 155, 807, 165, 10, 5)
a16 = CreateRoundRectRgn(796, 160, 808, 170, 10, 5)
a17 = CreateRoundRectRgn(793, 165, 811, 172, 5, 5)

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


'View Hasil Combine'
SetWindowRgn Pratama.hwnd, bg, True
End Sub
'Responsive Object -> Mouse Trigger'
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ReleaseCapture
    SendMessage Pratama.hwnd, &HA1, 2, 0&
End Sub








