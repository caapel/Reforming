VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Components"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public e As Double    '�������� ����������� (���/������)
Public Cp As Double   '�������� ����������� (���/��*�)
Public Cp0 As Double  '�������� ����������� (���/��*�) ��� 0 ��� � 298 �
Public dH As Double   '��������� ��������������(���/����)
Public dH0 As Double  '��������� ��������������(���/����) ��� 0 ��� � 298 �
Public S As Double    '��������(���/������)
Public S0 As Double   '��������(���/������) ��� 0 ��� � 298 �
Public N As Double    '���������� �������� (����)
Public X As Double    '�������� ���� (����/����)
Public M As Double    '����� �������� (�����)
Public Ms As Double   '�������� ���� (��/��)
Public Ma As Double   '�������� ����� ������� �.�.�.
Public P As Double    '����������� �������� (���)
Public Name As String '�������� ����������

Public Function Get_Cp()
    Cp = e / Ma
End Function

Public Function Get_dH(T)
    dH = dH0 + (T - 298) * (Cp0 * Ma + e) / 2000
End Function

Public Function Get_S(T, P)
    S = S0 + (Cp0 * Ma / 2 + e / 2) * Log(T / 298) - 8.31 * Log(P)
End Function

Public Function Get_M()
    M = N * Ma
End Function

Public Function Get_Ms(M_Sum)
    Ms = M / M_Sum
End Function

Public Function Get_X(N_Sum)
    X = N / N_Sum
End Function

Public Function Get_P(p0)
    P = X * p0
End Function
