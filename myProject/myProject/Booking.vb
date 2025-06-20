Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class Booking
    Dim myCon As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\USER\Documents\KCAU\2.2 sss 3.1 gss\project\Coding\FitFam\Backend.accdb")
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        UserDashboard.Show()
        Me.Close()

    End Sub

    Private Sub Booking_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ListBox1.Items.Clear()
        myCon.Open()
        Dim myselect As New OleDbCommand("SELECT EquipmentName FROM EquipmentTable Where EquipmentType= 'Chest' ", myCon)
        Using reader As OleDbDataReader = myselect.ExecuteReader()
            Try
                While reader.Read()
                    ListBox1.Items.Add(reader("EquipmentName").ToString())
                End While
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Using





        myCon.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        myCon.Open()
        ListBox1.Items.Clear()

        Dim myselect As New OleDbCommand("SELECT EquipmentName FROM EquipmentTable Where EquipmentType= 'Back' ", myCon)
        Using reader As OleDbDataReader = myselect.ExecuteReader()
            Try
                While reader.Read()
                    ListBox1.Items.Add(reader("EquipmentName").ToString())

                End While
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Using
        myCon.Close()
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        ' When an item is clicked, move it to ListBox1 on the UserDashboard form
        Dim selectedEquipment As String = ListBox1.SelectedItem.ToString()
        ' Create an instance of TimeSlotForm using the New keyword
        'Dim timeSlotForm As New Timeslot(selectedEquipment)
        ' timeSlotForm.ShowDialog()
        Timeslot.ShowDialog()



    End Sub
End Class