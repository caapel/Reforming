VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Main_interface 
   Caption         =   "Риформинг углеводородной газовой смеси"
   ClientHeight    =   6156
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   14328
   OleObjectBlob   =   "Main_interface.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Main_interface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public textBoxsIM As New Collection  'коллекция всех активных TextBox с компонентами из фрейма "Условия процесса"
Public textBoxsRM As New Collection  'коллекция всех TextBox из фрейма "Результат риформинга"
Public textBoxsA As New Collection   'коллекция всех расчетных TextBox со стехиометрическими коэффициентами

'---------заполнение формы дефолтными значениями при инициализации---------

Private Sub UserForm_Initialize()
    CB_air_type.AddItem "воздух"
    CB_air_type.AddItem "кислород"

    CB_catalyst_type.AddItem "NIAP-05-01(STK)"
    CB_catalyst_type.AddItem "NIAP-03-1"
    CB_catalyst_type.AddItem "K-905D1"
    CB_catalyst_type.AddItem "NTK-4"
    CB_catalyst_type.AddItem "Z206"

    TB_E.value = Format(10 ^ -16, "Scientific")
    TB_speed.value = 8
    
    Dim ctrl As MsForms.control
    For Each ctrl In Controls
        If TypeOf ctrl Is MsForms.textBox Then
            If InStr(ctrl.Name, "TB_IM") Then
                textBoxsIM.Add ctrl  'заполняем коллекцию textBoxsIM
            ElseIf InStr(ctrl.Name, "TB_RM") Then
                textBoxsRM.Add ctrl  'заполняем коллекцию textBoxsRM
            ElseIf InStr(ctrl.Name, "TB_a_") Then
                textBoxsA.Add ctrl  'заполняем коллекцию textBoxsA
            End If
        End If
    Next

End Sub

'---------события при раскрытии ComboBox---------

Private Sub CB_air_type_DropButtonClick()
    
    Dim textBox As Variant
    For Each textBox In textBoxsIM
        TB_IM_N_check_empty textBox
    Next
    
    TB_pressure_check_empty
    TB_temperature_check_empty
    Access_check
End Sub

Private Sub CB_catalyst_type_DropButtonClick()
    
    Dim textBox As Variant
    For Each textBox In textBoxsIM
        TB_IM_N_check_empty textBox
    Next
    
    TB_pressure_check_empty
    TB_temperature_check_empty
    Access_check
End Sub

'---------События при переключении в TextBox---------

Private Sub TB_temperature_C_Enter()
    If TB_temperature_C.Text = "750,0" Then TB_temperature_C.Text = ""
    
    Dim textBox As Variant
    For Each textBox In textBoxsIM
        TB_IM_N_check_empty textBox
    Next
    
    TB_pressure_check_empty
    Access_check
End Sub

Private Sub TB_pressure_bar_Enter()
    If TB_pressure_bar.Text = "1,0" Then TB_pressure_bar.Text = ""
    
    Dim textBox As Variant
    For Each textBox In textBoxsIM
        TB_IM_N_check_empty textBox
    Next
    
    TB_temperature_check_empty
    Access_check
End Sub

Private Sub TB_IM_H2_N_Enter()
    Check_textBox_empty TB_IM_H2_N
End Sub

Private Sub TB_IM_H2O_N_Enter()
    Check_textBox_empty TB_IM_H2O_N
End Sub

Private Sub TB_IM_N2_N_Enter()
    Check_textBox_empty TB_IM_N2_N
End Sub

Private Sub TB_IM_O2_N_Enter()
    Check_textBox_empty TB_IM_O2_N
End Sub

Private Sub TB_IM_CO_N_Enter()
    Check_textBox_empty TB_IM_CO_N
End Sub

Private Sub TB_IM_CO2_N_Enter()
    Check_textBox_empty TB_IM_CO2_N
End Sub

Private Sub TB_IM_CH4_N_Enter()
    Check_textBox_empty TB_IM_CH4_N
End Sub

Private Sub TB_IM_C2H6_N_Enter()
    Check_textBox_empty TB_IM_C2H6_N
End Sub

Private Sub TB_IM_C3H8_N_Enter()
    Check_textBox_empty TB_IM_C3H8_N
End Sub

Private Sub TB_IM_C4H10_N_Enter()
    Check_textBox_empty TB_IM_C4H10_N
End Sub

Private Sub Check_textBox_empty(ByRef object)
    If object.Text = "0" Then object.Text = ""
    
    Dim textBox As Variant
    For Each textBox In textBoxsIM
        If Not textBox Is object Then
            TB_IM_N_check_empty textBox
        End If
    Next
    
    TB_temperature_check_empty
    TB_pressure_check_empty
    Access_check
End Sub

'---------Восстановление дефолтных значений TextBox при переключении из пустого поля---------

Private Sub TB_pressure_check_empty()
    If TB_pressure_bar.Text = "" Then
        TB_pressure_bar.Text = "1,0"
        TB_pressure_bar.ForeColor = &H80000000
        TB_pressure_Pa.Text = TB_pressure_bar.Text * 10 ^ 5
        TB_pressure_Pa.ForeColor = &H80000000
    End If
End Sub
    
Private Sub TB_temperature_check_empty()
    If TB_temperature_C.Text = "" Then
        TB_temperature_C.Text = "750,0"
        TB_temperature_C.ForeColor = &H80000000
        TB_temperature_K.Text = TB_temperature_C.Text + 273
        TB_temperature_K.ForeColor = &H80000000
    End If
End Sub

Private Sub TB_IM_N_check_empty(ByRef object)
    If object.Text = "" Then
        object.Text = "0"
        object.ForeColor = vbBlack
    End If
    Check_total_N
End Sub

'---------Валидация переданных пользователем данных---------

'граничные условия для температуры процесса: [350; 1000]
Private Sub TB_temperature_C_KeyUp(ByVal KeyCode As MsForms.ReturnInteger, ByVal Shift As Integer)
    If IsNumeric(TB_temperature_C.Text) = False Then
        TB_temperature_C.ForeColor = vbRed
        TB_temperature_K.Text = "N/A"
        TB_temperature_K.ForeColor = &H80000000
    ElseIf CDbl(TB_temperature_C.Text) < 350 Or CDbl(TB_temperature_C.Text) > 1000 Then
        TB_temperature_C.ForeColor = vbRed
        TB_temperature_K.Text = TB_temperature_C.Text + 273
        TB_temperature_K.ForeColor = &H80000000
    Else
        TB_temperature_C.ForeColor = vbBlack
        TB_temperature_K.Text = TB_temperature_C.Text + 273
        TB_temperature_K.ForeColor = vbBlack
    End If
    Access_check
End Sub

'граничные условия для давления процесса: [0.5; 3.5]
Private Sub TB_pressure_bar_KeyUp(ByVal KeyCode As MsForms.ReturnInteger, ByVal Shift As Integer)
    If IsNumeric(TB_pressure_bar.Text) = False Then
        TB_pressure_bar.ForeColor = vbRed
        TB_pressure_Pa.Text = "N/A"
        TB_pressure_Pa.ForeColor = &H80000000
    ElseIf CDbl(TB_pressure_bar.Text) < 0.5 Or CDbl(TB_pressure_bar.Text) > 3.5 Then
        TB_pressure_bar.ForeColor = vbRed
        TB_pressure_Pa.Text = TB_pressure_bar.Text * 10 ^ 5
        TB_pressure_Pa.ForeColor = &H80000000
    Else
        TB_pressure_bar.ForeColor = vbBlack
        TB_pressure_Pa.Text = TB_pressure_bar.Text * 10 ^ 5
        TB_pressure_Pa.ForeColor = vbBlack
    End If
    Access_check
End Sub

Private Sub TB_IM_H2_N_KeyUp(ByVal KeyCode As MsForms.ReturnInteger, ByVal Shift As Integer)
    Check_Valid_Value TB_IM_H2_N
End Sub

Private Sub TB_IM_H2O_N_KeyUp(ByVal KeyCode As MsForms.ReturnInteger, ByVal Shift As Integer)
    Check_Valid_Value TB_IM_H2O_N
End Sub

Private Sub TB_IM_N2_N_KeyUp(ByVal KeyCode As MsForms.ReturnInteger, ByVal Shift As Integer)
    Check_Valid_Value TB_IM_N2_N
End Sub

Private Sub TB_IM_O2_N_KeyUp(ByVal KeyCode As MsForms.ReturnInteger, ByVal Shift As Integer)
    Check_Valid_Value TB_IM_O2_N
End Sub

Private Sub TB_IM_CO_N_KeyUp(ByVal KeyCode As MsForms.ReturnInteger, ByVal Shift As Integer)
    Check_Valid_Value TB_IM_CO_N
End Sub

Private Sub TB_IM_CO2_N_KeyUp(ByVal KeyCode As MsForms.ReturnInteger, ByVal Shift As Integer)
    Check_Valid_Value TB_IM_CO2_N
End Sub

Private Sub TB_IM_CH4_N_KeyUp(ByVal KeyCode As MsForms.ReturnInteger, ByVal Shift As Integer)
    Check_Valid_Value TB_IM_CH4_N
End Sub

Private Sub TB_IM_C2H6_N_KeyUp(ByVal KeyCode As MsForms.ReturnInteger, ByVal Shift As Integer)
    Check_Valid_Value TB_IM_C2H6_N
End Sub

Private Sub TB_IM_C3H8_N_KeyUp(ByVal KeyCode As MsForms.ReturnInteger, ByVal Shift As Integer)
    Check_Valid_Value TB_IM_C3H8_N
End Sub

Private Sub TB_IM_C4H10_N_KeyUp(ByVal KeyCode As MsForms.ReturnInteger, ByVal Shift As Integer)
    Check_Valid_Value TB_IM_C4H10_N
End Sub

'граничные условия для компонента: [0; 10]
Private Sub Check_Valid_Value(ByRef object)
    If IsNumeric(object.Text) = False Then
        object.ForeColor = vbRed
    ElseIf CDbl(object.Text) < 0 Or CDbl(object.Text) > 10 Then
        object.ForeColor = vbRed
    Else
        object.ForeColor = vbBlack
    End If
    
    Check_total_N
    Access_check
End Sub

'просчёт суммы молей для TB_Sum_N
Private Sub Check_total_N()
    Dim textBox As Variant
    TB_Sum_N.value = 0
    For Each textBox In textBoxsIM
        If IsNumeric(textBox.Text) Then
            TB_Sum_N.value = TB_Sum_N.value + CDbl(textBox.value)
        End If
    Next
    
    If TB_Sum_N.value = 0 Then
        TB_Sum_N.ForeColor = &H80000000
    Else
        TB_Sum_N.ForeColor = &H80000001
    End If
End Sub

'---------Проверка допуска к расчёту---------

Private Sub Access_check()

    If TB_temperature_C.ForeColor = vbBlack _
        And TB_pressure_bar.ForeColor = vbBlack _
        And TB_IM_H2_N.ForeColor <> vbRed _
        And TB_IM_H2O_N.ForeColor <> vbRed _
        And TB_IM_N2_N.ForeColor <> vbRed _
        And TB_IM_O2_N.ForeColor <> vbRed _
        And TB_IM_CO_N.ForeColor <> vbRed _
        And TB_IM_CO2_N.ForeColor <> vbRed _
        And TB_IM_CH4_N.ForeColor <> vbRed _
        And TB_IM_C2H6_N.ForeColor <> vbRed _
        And TB_IM_C3H8_N.ForeColor <> vbRed _
        And TB_IM_C4H10_N.ForeColor <> vbRed _
        And CB_catalyst_type.value <> "катализатор" _
        And (CB_air_type.value = "воздух" Or CB_air_type.value = "кислород") _
        And (TB_IM_CH4_N.value + TB_IM_C2H6_N.value + TB_IM_C3H8_N.value + TB_IM_C4H10_N.value > 0) Then
        
        reforming textBoxsIM 'допуск к расчёту
    Else
        Forbidden_calc  'блокировка расчёта
    End If
End Sub

Private Sub Forbidden_calc()
    
    Dim textBox As Variant
    
    For Each textBox In textBoxsA   'сбрасываем пропорции компонентов
        textBox.ForeColor = &H80000000
        textBox.value = "N/A"
    Next
    
    For Each textBox In textBoxsRM  'сбрасываем TextBox из фрейма "Результат риформинга"
        textBox.ForeColor = &H80000000
        textBox.value = "N/A"
    Next
    
End Sub

Public Sub Print_Result(Components_, a_H2O_CH, a_O2_CH, a_CO2_CH, a_O_C, a_H2O_C, i, Q, speed)
    
    Dim textBox As Variant, component As Components
    
    For Each textBox In textBoxsA   'заполняем TextBox с пропорциями компонентов
        textBox.ForeColor = &H80000001
    Next
    
    Main_interface.TB_a_H2O_CH.value = a_H2O_CH
    Main_interface.TB_a_O2_CH.value = a_O2_CH
    Main_interface.TB_a_CO2_CH.value = a_CO2_CH
    Main_interface.TB_a_O_C.value = a_O_C
    Main_interface.TB_a_H2O_C.value = a_H2O_C
    
    For Each textBox In textBoxsRM  'заполняем TextBox из фрейма "Результат риформинга"
        textBox.ForeColor = &H80000001
        
        For Each component In Components_
            If InStr(textBox.Name, "_" & component.Name & "_") Then
                If InStr(textBox.Name, "_N") Then
                    textBox.value = user_Format(component.N)
                ElseIf InStr(textBox.Name, "_X") Then
                    textBox.value = user_Format(component.X * 100)
                ElseIf InStr(textBox.Name, "_Ms") Then
                    textBox.value = user_Format(component.Ms * 100)
                ElseIf InStr(textBox.Name, "_P") Then
                    textBox.value = user_Format(component.P)
                Else
                    textBox.value = user_Format(component.M)
                End If
            End If
        Next
    Next
    
    'заполнение итоговых сумм компонентов
    TB_RM_Sum_N.value = user_Format(N_Sum(Components_))
    TB_RM_Sum_X.value = user_Format(X_Sum(Components_) * 100)
    TB_RM_Sum_P.value = user_Format(P_Sum(Components_))
    TB_RM_Sum_M.value = user_Format(M_Sum(Components_))
    TB_RM_Sum_Ms.value = user_Format(Ms_Sum(Components_) * 100)

    TB_RM_i.value = i
    TB_RM_Q.value = Round(Q, 3)
    
End Sub

Function user_Format(value)
    If value >= 0.1 And value <= 100 Then
        user_Format = Round(value, 3)
    ElseIf value = 0 Then
        user_Format = value
    Else
        user_Format = Format(value, "0.###e+")
    End If
End Function
