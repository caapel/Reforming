Attribute VB_Name = "Get_Lists"
'названи€ компонентов
Function Get_Name_list() As Collection

    Dim Name_list As New Collection
    
        Name_list.Add "H2"
        Name_list.Add "H2O"
        Name_list.Add "N2"
        Name_list.Add "O2"
        Name_list.Add "C"
        Name_list.Add "CO"
        Name_list.Add "CO2"
        Name_list.Add "CH4"
        Name_list.Add "C2H6"
        Name_list.Add "C3H8"
        Name_list.Add "C4H10"
    
    Set Get_Name_list = Name_list
    
End Function

'удельные теплоЄмкости веществ при 25∞— (кƒж/кгЈк)
Function Get_Cp0_list() As Collection

    Dim Cp0_list As New Collection
    
        Cp0_list.Add 14.31                          'H2
        Cp0_list.Add (32.8367 + 0.01166 * 25) / 18  'H2O
        Cp0_list.Add 1.04                           'N2
        Cp0_list.Add 0.92                           'O2
        Cp0_list.Add 0.7101                         'C
        Cp0_list.Add 0.93                           'CO
        Cp0_list.Add 0.85                           'CO2
        Cp0_list.Add 2.2                            'CH4 (метан)
        Cp0_list.Add 52.64 / 30                     'C2H6 (этан)
        Cp0_list.Add 1.67                           'C3H8( пропан)
        Cp0_list.Add 97.78 / 58.12                  'C4H10 (бутан)
    
    Set Get_Cp0_list = Cp0_list
    
End Function

'мол€рна€ масса веществ а.е.м.
Function Get_Ma_list() As Collection

    Dim Ma_list As New Collection
    
        Ma_list.Add 2       'H2
        Ma_list.Add 18      'H2O
        Ma_list.Add 28      'N2
        Ma_list.Add 32      'O2
        Ma_list.Add 12      'C
        Ma_list.Add 28      'CO
        Ma_list.Add 44      'CO2
        Ma_list.Add 16      'CH4 (метан)
        Ma_list.Add 30      'C2H6 (этан)
        Ma_list.Add 44      'C3H8 (пропан)
        Ma_list.Add 58      'C4H10 (бутан)
    
    Set Get_Ma_list = Ma_list
    
End Function

'термодинамические константы (0 ати, 298  )
Function Get_dH0_list() As Collection

    Dim dH0_list As New Collection
    
        dH0_list.Add 0          'H2
        dH0_list.Add -241.84    'H2O
        dH0_list.Add 0          'N2
        dH0_list.Add 0          'O2
        dH0_list.Add 0          'C
        dH0_list.Add -110.5     'CO
        dH0_list.Add -393.51    'CO2
        dH0_list.Add -74.85     'CH4 (метан)
        dH0_list.Add -84.67     'C2H6 (этан)
        dH0_list.Add -103.9     'C3H8 (пропан)
        dH0_list.Add -124.7     'C4H10 (бутан)
    
    Set Get_dH0_list = dH0_list
    
End Function

Function Get_S0_list() As Collection

    Dim S0_list As New Collection
    
        S0_list.Add 130.6     'H2
        S0_list.Add 188.74    'H2O
        S0_list.Add 191.5     'N2
        S0_list.Add 205.03    'O2
        S0_list.Add 5.74      'C
        S0_list.Add 197.4     'CO
        S0_list.Add 213.6     'CO2
        S0_list.Add 186.19    'CH4 (метан)
        S0_list.Add 229.5     'C2H6 (этан)
        S0_list.Add 269.9     'C3H8 (пропан)
        S0_list.Add 310       'C4H10 (бутан)
    
    Set Get_S0_list = S0_list
    
End Function

Function Get_N_list(ByRef Name_list As Collection, ByRef textBoxsIM As Collection) As Collection

    Dim N_list As New Collection
    
    
    For i = 1 To 11                         '5 Ц твердый углерод (исключен из исходной газовой смеси)
        If i = 5 Or i = 9 Or i = 11 Then    '9, 11 Ц C2H6 и C4H10 (временно не реализованы в логике программы)
            N_list.Add 0, Name_list(i)
        ElseIf textBoxsIM(i).value <> "" And i < 5 Then
            N_list.Add CDbl(textBoxsIM(i).value), Name_list(i)
        ElseIf textBoxsIM(i).value <> "" And i > 5 Then
            N_list.Add CDbl(textBoxsIM(i - 1).value), Name_list(i)
        Else
            N_list.Add 0, Name_list(i)
        End If
    Next
        
    Set Get_N_list = N_list
        
End Function

Function Get_Components_list(H2, H2O, N2, O2, C, CO, CO2, CH4, C2H6, C3H8, C4H10) As Collection
    
    Dim Components_ As New Collection
        Components_.Add H2
        Components_.Add H2O
        Components_.Add N2
        Components_.Add O2
        Components_.Add C
        Components_.Add CO
        Components_.Add CO2
        Components_.Add CH4
        Components_.Add C2H6
        Components_.Add C3H8
        Components_.Add C4H10
        
    Set Get_Components_list = Components_
        
End Function

Function Get_Reactions_list(reaction1, reaction2, reaction3, reaction4, reaction5, reaction6, _
                            reaction7, reaction8, reaction9, reaction10, reaction11, reaction12) As Collection
    
    Dim Reactions_ As New Collection
        Reactions_.Add reaction1
        Reactions_.Add reaction2
        Reactions_.Add reaction3
        Reactions_.Add reaction4
        Reactions_.Add reaction5
        Reactions_.Add reaction6
        Reactions_.Add reaction7
        Reactions_.Add reaction8
        Reactions_.Add reaction9
        Reactions_.Add reaction10
        Reactions_.Add reaction11
        Reactions_.Add reaction12
        
    Set Get_Reactions_list = Reactions_
        
End Function
