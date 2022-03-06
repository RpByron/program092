Imports System.Math
Module Module1
    ' Determianado variables 
    Public Nom(6) As String
    Public nit(6) As Integer
    Public honorarios(6) As Double
    Public numdias(6) As Integer
    Public Tpago(6) As String
    Public Thabitación(6) As String
    Public costohabitación As Double
    Public Pparcial(6) As Double
    Public Ptotal(6) As Double
    Public descuento(6) As Double
    Public cargo(6) As Double
    Public vector As Byte = 0

    Sub mostrarvectores()
        Form1.DataGridView1.Rows.Add(Nom(vector), Str(nit(vector)), Str(honorarios(vector)), Str(numdias(vector)), Thabitación(vector), Tpago(vector), Str(Round(Pparcial(vector), 2)), Str(Round(descuento(vector), 2)), Str(Round(cargo(vector), 2)), Str(Round(Ptotal(vector), 2)))
    End Sub


    ' Programacion de  botones 
    Sub lIMPIARGRIND()
        Form1.DataGridView1.Rows.Clear()
    End Sub
    Sub lIMPIARDATOS()
        Form1.TextBox1.Clear()
        Form1.TextBox2.Clear()
        Form1.TextBox3.Clear()
        Form1.TextBox4.Clear()
        Form1.ComboBox1.SelectedIndex = -1
        Form1.ComboBox2.SelectedIndex = -1
        Form1.TextBox1.Focus()

    End Sub

    Sub SALIR()
        If (MsgBox("Desea salir del programa ", vbQuestion + vbYesNo, "Mensaje") = vbYes) Then
            End
        Else
            lIMPIARDATOS()
            lIMPIARGRIND()


        End If
    End Sub
End Module
