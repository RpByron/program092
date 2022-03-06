Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        'progrmaanod boton calcular 
        If vector <= 5 Then
            If TextBox1.Text <> "" Then
                Nom(vector) = TextBox1.Text
            Else
                MsgBox("Error, Porfavor ingrese un Nombre")
                TextBox1.Focus() : Exit Sub
            End If
            If (IsNumeric(TextBox2.Text) And Val(TextBox2.Text <> "")) Then
                nit(vector) = Val(TextBox2.Text)
            Else
                MsgBox("Error, Porfavor ingrese un numero de nit.")
                TextBox1.Focus() : Exit Sub
            End If
            If (IsNumeric(TextBox3.Text) And Val(TextBox3.Text) > 0) Then
                honorarios(vector) = Val(TextBox3.Text)
            Else
                MsgBox("Error, porfavor ingrese un honorario")
                TextBox1.Focus() : Exit Sub
            End If
            If (IsNumeric(TextBox4.Text) And Val(TextBox4.Text) > 0) Then
                numdias(vector) = Val(TextBox4.Text)
            Else
                MsgBox("Error, Porfavor ingrese el numero de dias ")
                TextBox1.Focus() : Exit Sub
            End If


            'Programacion de los combobox

            Pparcial(vector) = honorarios(vector) + costohabitación
            Select Case ComboBox1.SelectedIndex
                Case 0 Or 1
                    descuento(vector) = Pparcial(vector) * 0.15

                Case Not descuento(vector) = 0
                Case 2
                    cargo(vector) = Pparcial(vector) * 0.03
                Case Not cargo(vector) = 0
                Case 3
                    cargo(vector) = Pparcial(vector) * 0.01
                Case Not cargo(vector) = 0
                Case Else
                    MsgBox("Error,Por favor seleccionar un tipo de pago ")
                    TextBox1.Focus() : Exit Sub
            End Select

            Select Case ComboBox2.SelectedIndex
                Case 0
                    costohabitación = 350 * numdias(vector)
                Case 1
                    costohabitación = 250 * numdias(vector)
                Case 2
                    costohabitación = 150 * numdias(vector)
                Case Else
                    MsgBox("Error, Porfavor seleccionar una habitacion")
                    TextBox1.Focus() : Exit Sub
            End Select



            Thabitación(vector) = ComboBox2.Text
            Tpago(vector) = ComboBox1.Text
            Ptotal(vector) = Pparcial(vector) - descuento(vector) + cargo(vector)
            mostrarvectores()
            vector = vector + 1
        End If


        If vector = 6 Then
            MsgBox("Sobre paso el limite de vectores")
        End If
        lIMPIARDATOS()
    End Sub


    ' Llmando programa para limpiar grind 
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        lIMPIARGRIND()
    End Sub

    ' Llmando programa para limpiar datos
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        lIMPIARDATOS()
    End Sub

    ' Llmando programa para salir
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        SALIR()
    End Sub
End Class




