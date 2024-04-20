VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Reactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public dHr As Double    'dHr(1, 298)
Public dS As Double     'dS(p, T) расч
Public dGr As Double    'dGr(T)
Public K As Double      'K(p, T)

Public Function Get_dGr(T)
    dGr = dHr - T * dS / 1000
End Function

Public Function Get_K(T)
    K = Exp(-dGr * 1000 / (8.314 * T))
End Function
