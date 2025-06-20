Imports System.Data.OleDb
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel
Imports Microsoft.VisualBasic.ApplicationServices

Public Class Timeslot

    Dim myCon As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\USER\Documents\KCAU\2.2 sss 3.1 gss\project\Coding\FitFam\Backend.accdb")


    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        If ListBox1.SelectedItem IsNot Nothing Then
            'get the userid from username

            Using myCon As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\USER\Documents\KCAU\2.2 sss 3.1 gss\project\Coding\FitFam\Backend.accdb")
                myCon.Open()

                Dim username As String = Login.TextBox3.Text
                Dim userID As Integer

                Dim querya As String = "SELECT UserID FROM UserTable WHERE Username = @Username"
                Using command1 As New OleDbCommand(querya, myCon)
                    command1.Parameters.AddWithValue("@Username", username)
                    Dim result As Object = command1.ExecuteScalar()
                    If result IsNot Nothing AndAlso Not DBNull.Value.Equals(result) Then
                        userID = Convert.ToInt32(result)
                        ' Display the UserID in a MessageBox
                        MessageBox.Show("Successfully registered for User ID: " & userID.ToString())
                    End If

                End Using


                '... (previous code)

                ' Get the equipment id from the selected equipment
                Dim eqname As String = Booking.ListBox1.SelectedItem.ToString() ' Assuming a single item is selected
                Dim eqID As Integer

                Dim query2 As String = "SELECT EquipId FROM EquipmentTable WHERE EquipmentName = @eqname"

                Using command3 As New OleDbCommand(query2, myCon)
                    command3.Parameters.AddWithValue("@eqname", eqname)
                    Dim resultEq As Object = command3.ExecuteScalar()

                    If resultEq IsNot Nothing AndAlso Not DBNull.Value.Equals(resultEq) Then
                        eqID = Convert.ToInt32(resultEq)

                        ' Continue with the rest of the code
                        ' ...
                    Else
                        ' Display an error or handle the case where the EquipmentID does not exist
                        MessageBox.Show("Selected equipment does not exist.")
                    End If
                End Using

                '... (rest of the code)



                ' Get time slot id from selected time slot
                Dim timen As String = Me.ListBox1.SelectedItem.ToString() ' Assuming a single item is selected
                Dim timeid As Integer

                Dim query3 As String = "SELECT TimeID FROM TimeSlotTable WHERE Time = @timen"

                Using command5 As New OleDbCommand(query3, myCon)
                    command5.Parameters.AddWithValue("@timen", timen)
                    Dim resultTime As Object = command5.ExecuteScalar()

                    If resultTime IsNot Nothing AndAlso Not DBNull.Value.Equals(resultTime) Then
                        timeid = Convert.ToInt32(resultTime)
                    End If
                End Using

                Dim query As String = "INSERT INTO BookingTable (UserID, EquipmentID, TimeID) VALUES (@userID, @eqID, @timeid)"

                Using command As New OleDbCommand(query, myCon)
                    command.Parameters.AddWithValue("@userID", userID)
                    command.Parameters.AddWithValue("@eqID", eqID)
                    command.Parameters.AddWithValue("@timeid", timeid)
                    command.ExecuteNonQuery()
                End Using

                ' Update time slot status to 'Booked'
                Dim updateQuery As String = "UPDATE TimeSlotTable SET Status = 'Booked' WHERE TimeID = @timeid"

                Using updateCommand As New OleDbCommand(updateQuery, myCon)
                    updateCommand.Parameters.AddWithValue("@timeid", timeid)
                    updateCommand.ExecuteNonQuery()
                End Using
                myCon.Close()
            End Using
        End If
    End Sub





    Private Sub Timeslot_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Using myCon As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\USER\Documents\KCAU\2.2 sss 3.1 gss\project\Coding\FitFam\Backend.accdb")

            myCon.Open()
            Dim mycmd As New OleDbCommand("SELECT Time FROM TimeSlotTable where status= 'unbooked' ", myCon)
            Dim reader As OleDbDataReader = mycmd.ExecuteReader()
            Try
                While reader.Read()
                    ListBox1.Items.Add(reader("Time").ToString())
                End While
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            myCon.Close()
        End Using
    End Sub
End Class