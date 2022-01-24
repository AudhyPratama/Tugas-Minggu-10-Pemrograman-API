VERSION 5.00
Begin VB.Form Brilliant 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17370
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   17370
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Brilliant"
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

'================================================= B ==========================================
b1 = CreateRoundRectRgn(450, 30, 485, 40, 10, 5)
b1a = CreateRoundRectRgn(450, 82, 485, 72, 5, 5)

b2 = CreateRoundRectRgn(451, 32, 486, 42, 10, 5)
b2a = CreateRoundRectRgn(451, 80, 486, 70, 25, 25)

b3 = CreateRoundRectRgn(455, 33, 468, 82, 10, 15)

b4 = CreateRoundRectRgn(476, 37, 488, 47, 10, 5)
b5 = CreateRoundRectRgn(477, 42, 489, 52, 10, 5)
b6 = CreateRoundRectRgn(476, 47, 488, 57, 10, 5)
b7 = CreateRoundRectRgn(455, 52, 487, 60, 10, 5)
b8 = CreateRoundRectRgn(475, 55, 489, 65, 10, 5)
b9 = CreateRoundRectRgn(476, 58, 490, 70, 10, 5)
b10 = CreateRoundRectRgn(477, 61, 491, 73, 10, 5)
b11 = CreateRoundRectRgn(476, 64, 490, 75, 10, 5)
b12 = CreateRoundRectRgn(475, 67, 489, 78, 10, 5)
b13 = CreateRoundRectRgn(474, 70, 488, 81, 10, 5)

CombineRgn bg, bg, b1, 2
CombineRgn bg, bg, b1a, 2
CombineRgn bg, bg, b2a, 2
CombineRgn bg, bg, b2, 2
CombineRgn bg, bg, b3, 2
CombineRgn bg, bg, b4, 2
CombineRgn bg, bg, b5, 2
CombineRgn bg, bg, b6, 2
CombineRgn bg, bg, b7, 2
CombineRgn bg, bg, b8, 2
CombineRgn bg, bg, b9, 2
CombineRgn bg, bg, b10, 2
CombineRgn bg, bg, b11, 2
CombineRgn bg, bg, b12, 2
CombineRgn bg, bg, b13, 2


'================================================= R ==========================================
r1 = CreateRoundRectRgn(515, 30, 548, 40, 5, 5)
r2 = CreateRoundRectRgn(515, 70, 535, 80, 5, 5)
r3 = CreateRoundRectRgn(520, 30, 530, 80, 10, 15)

r4 = CreateRoundRectRgn(540, 32, 550, 45, 10, 5)
r5 = CreateRoundRectRgn(541, 35, 551, 47, 10, 5)
r6 = CreateRoundRectRgn(541, 38, 552, 51, 10, 5)
r7 = CreateRoundRectRgn(541, 41, 552, 51, 10, 5)
r7a = CreateRoundRectRgn(541, 44, 552, 53, 10, 5)
r7b = CreateRoundRectRgn(540, 47, 551, 55, 10, 5)
r8 = CreateRoundRectRgn(539, 50, 550, 57, 10, 5)
r9 = CreateRoundRectRgn(538, 53, 549, 59, 10, 5)
r10 = CreateRoundRectRgn(537, 56, 548, 61, 10, 5)
r11 = CreateRoundRectRgn(525, 53, 547, 61, 5, 5)

r12 = CreateRoundRectRgn(538, 55, 548, 65, 5, 5)
r13 = CreateRoundRectRgn(539, 60, 549, 70, 5, 5)
r14 = CreateRoundRectRgn(540, 65, 550, 75, 5, 5)
r15 = CreateRoundRectRgn(541, 70, 553, 80, 5, 5)

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


'================================================= I ==========================================
i1 = CreateRoundRectRgn(583, 30, 606, 40, 5, 5)
i2 = CreateRoundRectRgn(589, 30, 600, 80, 10, 5)
i3 = CreateRoundRectRgn(583, 80, 606, 70, 5, 5)

CombineRgn bg, bg, i1, 2
CombineRgn bg, bg, i2, 2
CombineRgn bg, bg, i3, 2

'================================================= L ==========================================
l1 = CreateRoundRectRgn(630, 30, 653, 40, 5, 5)
l2 = CreateRoundRectRgn(636, 30, 647, 80, 10, 5)
l3 = CreateRoundRectRgn(630, 80, 665, 70, 5, 5)

l4 = CreateRoundRectRgn(658, 60, 665, 70, 15, 1)
l5 = CreateRoundRectRgn(657, 61, 665, 71, 15, 1)
l6 = CreateRoundRectRgn(656, 62, 665, 72, 15, 1)
l7 = CreateRoundRectRgn(655, 63, 665, 73, 15, 1)
l8 = CreateRoundRectRgn(654, 64, 665, 74, 15, 1)
l9 = CreateRoundRectRgn(653, 65, 665, 75, 20, 1)
l10 = CreateRoundRectRgn(652, 66, 665, 76, 25, 1)
l11 = CreateRoundRectRgn(651, 67, 665, 77, 25, 1)
l12 = CreateRoundRectRgn(650, 68, 665, 78, 25, 1)
l13 = CreateRoundRectRgn(649, 69, 665, 79, 25, 1)
l14 = CreateRoundRectRgn(648, 70, 665, 80, 25, 1)

CombineRgn bg, bg, l1, 2
CombineRgn bg, bg, l2, 2
CombineRgn bg, bg, l3, 2
CombineRgn bg, bg, l4, 2
CombineRgn bg, bg, l5, 2
CombineRgn bg, bg, l6, 2
CombineRgn bg, bg, l7, 2
CombineRgn bg, bg, l8, 2
CombineRgn bg, bg, l9, 2
CombineRgn bg, bg, l10, 2
CombineRgn bg, bg, l11, 2
CombineRgn bg, bg, l12, 2
CombineRgn bg, bg, l13, 2
CombineRgn bg, bg, l14, 2

'================================================= L ==========================================
l1 = CreateRoundRectRgn(690, 30, 713, 40, 5, 5)
l2 = CreateRoundRectRgn(696, 30, 707, 80, 10, 5)
l3 = CreateRoundRectRgn(690, 80, 725, 70, 5, 5)

l4 = CreateRoundRectRgn(718, 60, 725, 70, 15, 1)
l5 = CreateRoundRectRgn(717, 61, 725, 71, 15, 1)
l6 = CreateRoundRectRgn(716, 62, 725, 72, 15, 1)
l7 = CreateRoundRectRgn(715, 63, 725, 73, 15, 1)
l8 = CreateRoundRectRgn(714, 64, 725, 74, 15, 1)
l9 = CreateRoundRectRgn(713, 65, 725, 75, 20, 1)
l10 = CreateRoundRectRgn(712, 66, 725, 76, 25, 1)
l11 = CreateRoundRectRgn(711, 67, 725, 77, 25, 1)
l12 = CreateRoundRectRgn(710, 68, 725, 78, 25, 1)
l13 = CreateRoundRectRgn(709, 69, 725, 79, 25, 1)
l14 = CreateRoundRectRgn(708, 70, 725, 80, 25, 1)

CombineRgn bg, bg, l1, 2
CombineRgn bg, bg, l2, 2
CombineRgn bg, bg, l3, 2
CombineRgn bg, bg, l4, 2
CombineRgn bg, bg, l5, 2
CombineRgn bg, bg, l6, 2
CombineRgn bg, bg, l7, 2
CombineRgn bg, bg, l8, 2
CombineRgn bg, bg, l9, 2
CombineRgn bg, bg, l10, 2
CombineRgn bg, bg, l11, 2
CombineRgn bg, bg, l12, 2
CombineRgn bg, bg, l13, 2
CombineRgn bg, bg, l14, 2

'================================================= I ==========================================
i1 = CreateRoundRectRgn(753, 30, 776, 40, 5, 5)
i2 = CreateRoundRectRgn(759, 30, 770, 80, 10, 5)
i3 = CreateRoundRectRgn(753, 80, 776, 70, 5, 5)

CombineRgn bg, bg, i1, 2
CombineRgn bg, bg, i2, 2
CombineRgn bg, bg, i3, 2

'================================================= A ==========================================
a1 = CreateRoundRectRgn(800, 30, 820, 40, 10, 5)
a2 = CreateRoundRectRgn(799, 35, 821, 45, 10, 5)
a3 = CreateRoundRectRgn(798, 40, 809, 50, 10, 5)
a4 = CreateRoundRectRgn(797, 45, 808, 55, 10, 5)
a5 = CreateRoundRectRgn(796, 50, 807, 60, 10, 5)
a6 = CreateRoundRectRgn(795, 55, 806, 65, 10, 5)
a7 = CreateRoundRectRgn(794, 60, 826, 70, 10, 5)
a8 = CreateRoundRectRgn(812, 40, 822, 50, 10, 5)
a9 = CreateRoundRectRgn(813, 45, 823, 55, 10, 5)
a10 = CreateRoundRectRgn(814, 50, 824, 60, 10, 5)
a11 = CreateRoundRectRgn(815, 55, 825, 75, 10, 5)
a12 = CreateRoundRectRgn(793, 65, 804, 75, 10, 5)
a13 = CreateRoundRectRgn(792, 70, 803, 80, 10, 5)
a14 = CreateRoundRectRgn(789, 75, 807, 82, 5, 5)
a15 = CreateRoundRectRgn(817, 65, 827, 75, 10, 5)
a16 = CreateRoundRectRgn(816, 70, 828, 80, 10, 5)
a17 = CreateRoundRectRgn(813, 75, 831, 82, 5, 5)

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

'================================================= N ==========================================
n1 = CreateRoundRectRgn(859, 30, 876, 40, 5, 5)
n2 = CreateRoundRectRgn(864, 30, 874, 80, 5, 5)
n3 = CreateRoundRectRgn(859, 80, 879, 70, 5, 5)

n4 = CreateRoundRectRgn(882, 30, 902, 40, 5, 5)
n5 = CreateRoundRectRgn(887, 30, 897, 80, 5, 5)
n6 = CreateRoundRectRgn(887, 80, 896, 70, 5, 5)

n7 = CreateRoundRectRgn(864, 32, 879, 46, 5, 5)
n8 = CreateRoundRectRgn(865, 34, 880, 48, 5, 5)
n9 = CreateRoundRectRgn(866, 36, 881, 50, 5, 5)
n10 = CreateRoundRectRgn(867, 38, 882, 52, 5, 5)
n11 = CreateRoundRectRgn(868, 40, 883, 54, 5, 5)
n12 = CreateRoundRectRgn(869, 42, 884, 56, 5, 5)
n13 = CreateRoundRectRgn(870, 44, 885, 58, 5, 5)
n14 = CreateRoundRectRgn(873, 46, 886, 60, 5, 5)
n15 = CreateRoundRectRgn(874, 48, 887, 62, 5, 5)
n16 = CreateRoundRectRgn(875, 50, 888, 64, 5, 5)
n17 = CreateRoundRectRgn(876, 52, 889, 66, 5, 5)
n18 = CreateRoundRectRgn(877, 54, 890, 68, 5, 5)
n19 = CreateRoundRectRgn(878, 56, 891, 70, 5, 5)
n20 = CreateRoundRectRgn(879, 58, 892, 72, 5, 5)
n21 = CreateRoundRectRgn(880, 60, 893, 74, 5, 5)
n22 = CreateRoundRectRgn(881, 62, 894, 76, 5, 5)
n23 = CreateRoundRectRgn(882, 64, 895, 78, 5, 5)
n24 = CreateRoundRectRgn(883, 66, 896, 80, 5, 5)


CombineRgn bg, bg, n1, 2
CombineRgn bg, bg, n2, 2
CombineRgn bg, bg, n3, 2
CombineRgn bg, bg, n4, 2
CombineRgn bg, bg, n5, 2
CombineRgn bg, bg, n6, 2
CombineRgn bg, bg, n7, 2
CombineRgn bg, bg, n8, 2
CombineRgn bg, bg, n9, 2
CombineRgn bg, bg, n10, 2
CombineRgn bg, bg, n11, 2
CombineRgn bg, bg, n12, 2
CombineRgn bg, bg, n13, 2
CombineRgn bg, bg, n14, 2
CombineRgn bg, bg, n15, 2
CombineRgn bg, bg, n16, 2
CombineRgn bg, bg, n17, 2
CombineRgn bg, bg, n18, 2
CombineRgn bg, bg, n19, 2
CombineRgn bg, bg, n20, 2
CombineRgn bg, bg, n21, 2
CombineRgn bg, bg, n22, 2
CombineRgn bg, bg, n23, 2
CombineRgn bg, bg, n24, 2

'================================================= T ==========================================
t1 = CreateRoundRectRgn(925, 30, 965, 40, 5, 5)
t2 = CreateRoundRectRgn(940, 30, 950, 80, 5, 5)
t3 = CreateRoundRectRgn(935, 72, 955, 80, 5, 5)

t4 = CreateRoundRectRgn(925, 36, 935, 42, 25, 1)
t5 = CreateRoundRectRgn(925, 37, 934, 43, 25, 1)
t6 = CreateRoundRectRgn(925, 38, 933, 44, 25, 1)

t7 = CreateRoundRectRgn(955, 36, 965, 42, 25, 1)
t8 = CreateRoundRectRgn(956, 37, 965, 43, 25, 1)
t9 = CreateRoundRectRgn(957, 38, 965, 44, 25, 1)

CombineRgn bg, bg, t1, 2
CombineRgn bg, bg, t2, 2
CombineRgn bg, bg, t3, 2
CombineRgn bg, bg, t4, 2
CombineRgn bg, bg, t5, 2
CombineRgn bg, bg, t6, 2
CombineRgn bg, bg, t7, 2
CombineRgn bg, bg, t8, 2
CombineRgn bg, bg, t9, 2

'View Hasil Combine'
SetWindowRgn Brilliant.hwnd, bg, True
End Sub
'Responsive Object -> Mouse Trigger'
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ReleaseCapture
    SendMessage Brilliant.hwnd, &HA1, 2, 0&
End Sub







