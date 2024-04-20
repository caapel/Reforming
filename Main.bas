Attribute VB_Name = "Main"
Option Explicit

Public Sub reforming(ByRef textBoxsIM As Collection)
    On Error Resume Next
    
    'Инициализация исходных компонентов
    Dim H2 As New Components
    Dim H2O As New Components
    Dim N2 As New Components
    Dim O2 As New Components
    Dim C As New Components
    Dim CO As New Components
    Dim CO2 As New Components
    Dim CH4 As New Components
    Dim C2H6 As New Components  'задел на дальнейшее расширение компонентной базы
    Dim C3H8 As New Components
    Dim C4H10 As New Components 'задел на дальнейшее расширение компонентной базы

    'Инициализация коллекции компонентов
    Dim Components_ As Collection
    Set Components_ = Get_Components_list(H2, H2O, N2, O2, C, CO, CO2, CH4, C2H6, C3H8, C4H10)
    
    'получение коллекций с термодинамическими константами компонентов
    Dim Ma_list, S0_list, Cp0_list, dH0_list, Name_list As Collection
        Set Ma_list = Get_Ma_list()
        Set S0_list = Get_S0_list()
        Set Cp0_list = Get_Cp0_list()
        Set dH0_list = Get_dH0_list()
        Set Name_list = Get_Name_list()
    
    'объявление режимных параметров (°С, бар)
    Dim T, P As Double
        T = CDbl(Main_interface.TB_temperature_K.value)
        P = CDbl(Main_interface.TB_pressure_bar.value)
    
    'объявление опциональных параметров (временно не реализованы в логике программы)
    Dim air_type, catalyst_type As String
        air_type = Main_interface.CB_air_type.value
        catalyst_type = Main_interface.CB_catalyst_type.value
    
    'инициализация состава исходного газа (моль, данные от пользователя)
    Dim N_list As Collection
        Set N_list = Get_N_list(Name_list, textBoxsIM)
    
    'Заполнение полей (характеристик) компонентов исходной газовой смеси
    H2.e = 4.17 * (6.52 + 0.00078 * T + 1200 / T ^ 2)
    H2O.e = 32.8367 + 0.01166 * (T - 273)
    N2.e = 4.17 * (6.66 + 0.00102 * T)
    O2.e = 4.17 * (7.16 + 0.001 * T - 40000 / T ^ 2)
    C.e = 4.17 * (4.01 + 0.00102 * T - 200000 / T ^ 2)
    CO.e = 4.17 * (6.79 + 0.00098 * T - 1100 / T ^ 2)
    CO2.e = 4.17 * (10.55 + 0.00216 * T - 204000 / T ^ 2)
    CH4.e = 4.17 * (5.34 + 0.0115 * T)
    C2H6.e = 5.75 + 175.11 * T / 1000 - 57.85 * T ^ 2 / 1000000
    C3H8.e = 1.72 + 270.75 * T / 1000 - 4.48 * T ^ 2 / 1000000
    C4H10.e = 4.17 * (0.112 + 92.11 * T / 1000 - 47.53 * T ^ 2 / 10000 + 9.55 * T ^ 3 / 1000) 'из справочника К.П. Мищенко и А.А. Равделя
    
    Dim i As Integer
    For i = 1 To Components_.Count
        Components_(i).Name = Name_list(i)
        Components_(i).Ma = Ma_list(i)
        Components_(i).S0 = S0_list(i)
        Components_(i).Cp0 = Cp0_list(i)
        Components_(i).dH0 = dH0_list(i)
        Components_(i).Get_Cp
        Components_(i).Get_dH (T)
        Components_(i).Get_S T, P
        Components_(i).N = N_list(i)
        Components_(i).Get_M
    Next i

    'определение относительных характеристик (Ms, X, P) компонентов
    Get_Specific_Parameters Components_, P
    
    'пропорции реагентов
    Dim a_H2O_CH, a_O2_CH, a_CO2_CH, a_O_C, a_H2O_C As Double
    a_H2O_CH = H2O.N / (CH4.N + 3 * C3H8.N)
    a_O2_CH = O2.N / (CH4.N + 3 * C3H8.N)
    a_CO2_CH = CO2.N / (CH4.N + 3 * C3H8.N)
    a_O_C = (H2O.N + 2 * CO2.N + CO.N) / (CH4.N + 3 * C3H8.N + CO2.N + CO.N)
    a_H2O_C = H2O.N / (CH4.N + 3 * C3H8.N + CO2.N + CO.N)

    'инициализация значимых химических реакций
    Dim reaction1 As New Reactions 'CH4 + H2O = CO + 3H2
        reaction1.dHr = CO.dH + 3 * H2.dH - CH4.dH - H2O.dH
        reaction1.dS = CO.S + 3 * H2.S - CH4.S - H2O.S  'K(I,CH4)
        
    Dim reaction2 As New Reactions 'CO + H2O = CO2 + H2
        reaction2.dHr = CO2.dH + H2.dH - CO.dH - H2O.dH
        reaction2.dS = CO2.S + H2.S - CO.S - H2O.S  'K(CO)
        
    Dim reaction3 As New Reactions 'CH4 + 2H2O = CO2 + 4H2
        reaction3.dHr = CO2.dH + 4 * H2.dH - CH4.dH - 2 * H2O.dH
        reaction3.dS = CO2.S + 4 * H2.S - CH4.S - 2 * H2O.S  'K(II,CH4)
    
    Dim reaction4 As New Reactions 'CH4 = C + 2H2
        reaction4.dHr = C.dH + 2 * H2.dH - CH4.dH
        reaction4.dS = C.S + 2 * H2.S - CH4.S   'K(d, CH4)
    
    Dim reaction5 As New Reactions '2CO = C + CO2
        reaction5.dHr = CO2.dH + C.dH - 2 * CO.dH
        reaction5.dS = CO2.S + C.S - 2 * CO.S   'K(RB)
    
    Dim reaction6 As New Reactions 'CO + H2 = C + H2O
        reaction6.dHr = C.dH + H2O.dH - CO.dH - H2.dH
        reaction6.dS = C.S + H2O.S - CO.S - H2.S  'K(d, CO)
    
    Dim reaction7 As New Reactions 'CH4 + 0.5 O2 = CO + 2H2
        reaction7.dHr = CO.dH + 2 * H2.dH - CH4.dH - 0.5 * O2.dH
        reaction7.dS = CO.S + 2 * H2.S - CH4.S - 0.5 * O2.S  'K(Ox, CH4)
    
    Dim reaction8 As New Reactions 'CO + 0.5 O2 = CO2
        reaction8.dHr = CO2.dH - CO.dH - 0.5 * O2.dH
        reaction8.dS = CO2.S - CO.S - 0.5 * O2.S  'K(Ox, CO)
    
    Dim reaction9 As New Reactions 'H2 + 0.5 O2 = H2O
        reaction9.dHr = H2O.dH - H2.dH - 0.5 * O2.dH
        reaction9.dS = H2O.S - H2.S - 0.5 * O2.S  'K(Ox, H2)
    
    Dim reaction10 As New Reactions 'CH4 + CO2 = 2CO + 2H2
        reaction10.dHr = 2 * CO.dH + 2 * H2.dH - CH4.dH - CO2.dH
        reaction10.dS = 2 * CO.S + 2 * H2.S - CH4.S - CO2.S  'K(CO2)
    
    Dim reaction11 As New Reactions 'C3H8 + 3H2O(г) = 3CO + 7H2
        reaction11.dHr = 3 * CO.dH + 7 * H2.dH - C3H8.dH - 3 * H2O.dH
        reaction11.dS = 3 * CO.S + 7 * H2.S - C3H8.S - 3 * H2O.S  'K(I, C3)
    
    Dim reaction12 As New Reactions 'C3H8 + 6H2O = 3CO2 + 10H2
        reaction12.dHr = 3 * CO2.dH + 10 * H2.dH - C3H8.dH - 6 * H2O.dH
        reaction12.dS = 3 * CO2.S + 10 * H2.S - C3H8.S - 6 * H2O.S  'K(II, C3)
    
    'инициализация коллекции реакций
    Dim Reactions_ As Collection
    Set Reactions_ = Get_Reactions_list(reaction1, reaction2, reaction3, reaction4, reaction5, reaction6, _
                                        reaction7, reaction8, reaction9, reaction10, reaction11, reaction12)
    Dim reaction As New Reactions
    For Each reaction In Reactions_
        reaction.Get_dGr (T)
        reaction.Get_K (T)  'заполнение полей констант равновесий
    Next
    
    Dim N_Sum_before, N_Sum_after As Double, speed As Integer
    i = 0       'счётчик итераций
    speed = CInt(Main_interface.TB_speed)   'коэффициент скорости схождения (5 – 20)
    
    Do
        N_Sum_before = N_Sum(Components_) - C.N 'сумма молей компонентов до итерации
        
        i = i + 1
        If i = 5000 Then 'критическое число итераций
            MsgBox "Расходящийся цикл: измените начальные условия или понизьте точность расчёта", vbExclamation, "Превышено время ожидания"
            Exit Do
        End If

        'H2.N
        If i = 1 Then
            If a_O2_CH > 1 Then
                If a_O2_CH > 2 Then
                    H2.N = 0
                Else
                    H2.N = CH4.N + 3 * C3H8.N - O2.N
                End If
            Else
                H2.N = 4 * C3H8.N + 2 * CH4.N * (T - 273) / 800 + H2O.N / 2
            End If
        Else
            If a_O2_CH > 1 Then
                If a_O2_CH > 2 Then
                    H2.N = 0
                Else
                    H2.N = 2 - 2 * a_O2_CH
                End If
            Else
                H2.N = H2.N - O2.N / speed
            End If
        End If

        'CH4.N
        If a_O2_CH > 1 Then
            CH4.N = 0
        Else
            CH4.N = H2.N ^ 2 / (reaction4.K * N_Sum_before)
        End If

        'H2O.N
        If i = 1 Then
            If 4 * N_list("C3H8") + 2 * N_list("CH4") + N_list("H2O") + N_list("H2") - 2 * CH4.N - H2.N < 0 Then
                H2O.N = 0.2
            Else
                H2O.N = 4 * N_list("C3H8") + 2 * N_list("CH4") + N_list("H2O") + N_list("H2") - 2 * CH4.N - H2.N
            End If
        Else
            H2O.N = 4 * N_list("C3H8") + 2 * N_list("CH4") + N_list("H2O") + N_list("H2") - 2 * CH4.N - H2.N - 4 * C3H8.N
        End If
        
        'CO.N (проверить общий случай)
        If i > 7 And C.N = 0 Then
            CO.N = (N_list("CH4") + 3 * N_list("C3H8") + N_list("CO") + N_list("CO2") - CH4.N - 3 * C3H8.N) / (1 + (reaction2.K * H2O.N / H2.N))
        Else
            CO.N = H2O.N / (reaction6.K * H2.N)
        End If
        
        'CO2.N
        If i = 1 Then
            CO2.N = reaction5.K * CO.N ^ 2 / N_Sum_before
        Else
            CO2.N = reaction2.K * CO.N * H2O.N / H2.N
        End If

        'C3H8.N
        If C3H8.N = 0 Then
            C3H8.N = 0
        Else
            C3H8.N = (CO.N ^ 3 * H2.N ^ 7) / (reaction11.K * H2O.N ^ 3 * N_Sum_before ^ 6)
        End If
        
        'O2.N
        O2.N = (N_list("H2O") + 2 * N_list("O2") + 2 * N_list("CO2") + 2 * N_list("CO") - H2O.N - 2 * CO2.N - CO.N) / 2
        
        'C.N (проверить общий случай)
        If i > 6 And C.N <= 0 Then
            C.N = 0
        Else
            C.N = N_list("CH4") + 3 * N_list("C3H8") + N_list("CO2") + N_list("CO") - CH4.N - 3 * C3H8.N - CO2.N - CO.N
        End If
        
        'N2.N (не меняется при итерациях)
        N_Sum_after = N_Sum(Components_) - C.N 'сумма молей компонентов после итерации

    Loop While Abs(N_Sum_before - N_Sum_after) > CDbl(Main_interface.TB_E.value) 'оценка точности расчёта по сумме мольных долей компонентов

    'определение массы (M) и относительных характеристик (Ms, X, P) компонентов новой системы
    Dim component As Components
    For Each component In Components_
        component.Get_M
    Next

    Get_Specific_Parameters Components_, P

    Dim Q As Double 'Тепловой эффект процесса, кДж/моль
    Q = -2 * N_list("CO") * reaction7.dHr - (CO.N - 2 * N_list("CO")) * reaction1.dHr - CO2.N * reaction3.dHr

    Call Main_interface.Print_Result(Components_, a_H2O_CH, a_O2_CH, a_CO2_CH, a_O_C, a_H2O_C, i, Q, speed)
    
End Sub

Function Get_Specific_Parameters(ByRef Components_ As Collection, ByVal P As Single)

    Dim component As Components
    For Each component In Components_
        component.Get_Ms (M_Sum(Components_))  'заполнение поля относительной массы компонента
        component.Get_X (N_Sum(Components_))   'заполнение поля доли компонента
        component.Get_P (P)                    'заполнение поля парциального давления компонента
    Next
    
End Function

Function N_Sum(ByVal Components_ As Collection) As Double 'суммарное количество вещества
    
    Dim component As Components
    For Each component In Components_
        N_Sum = N_Sum + component.N
    Next
    
End Function

Function M_Sum(ByVal Components_ As Collection) As Double 'суммарная масса всех компонентов
    
    Dim component As Components
    For Each component In Components_
        M_Sum = M_Sum + component.M
    Next
    
End Function

Function X_Sum(ByVal Components_ As Collection) As Double 'сумма мольных долей всех компонентов
    
    Dim component As Components
    For Each component In Components_
        X_Sum = X_Sum + component.X
    Next
    
End Function

Function Ms_Sum(ByVal Components_ As Collection) As Double 'сумма массовых долей всех компонентов
    
    Dim component As Components
    For Each component In Components_
        Ms_Sum = Ms_Sum + component.Ms
    Next
    
End Function

Function P_Sum(ByVal Components_ As Collection) As Double 'сумма парциальных давлений всех компонентов
    
    Dim component As Components
    For Each component In Components_
        P_Sum = P_Sum + component.P
    Next
    
End Function
