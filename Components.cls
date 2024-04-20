VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Components"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public e As Double    'удельная теплоёмкость (кДж/моль·К)
Public Cp As Double   'удельная теплоёмкость (кДж/кг*К)
Public Cp0 As Double  'удельная теплоёмкость (кДж/кг*К) при 0 ати и 298 К
Public dH As Double   'энтальпия преобразования(кДж/моль)
Public dH0 As Double  'энтальпия преобразования(кДж/моль) при 0 ати и 298 К
Public S As Double    'энтропия(кДж/моль·К)
Public S0 As Double   'энтропия(кДж/моль·К) при 0 ати и 298 К
Public N As Double    'количество вещества (моль)
Public X As Double    'объемная доля (моль/моль)
Public M As Double    'масса вещества (грамм)
Public Ms As Double   'массовая доля (кг/кг)
Public Ma As Double   'молярная масса веществ а.е.м.
Public P As Double    'парциальное давление (бар)
Public Name As String 'название компонента

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
