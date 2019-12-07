Imports System.Data.Odbc
Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TampilGrid()
        MunculCombo()

    End Sub
    Sub TampilGrid()
        bukakoneksi()

        DA = New OdbcDataAdapter("select * From data_dompetku", CONN)
        DS = New DataSet
        DA.Fill(DS, "data_dompetku")
        DataGridView1.DataSource = DS.Tables("data_dompetku")

        TutupKoneksi()
    End Sub
    Sub MunculCombo()
        ComboBox1.Items.Add("Masuk")
        ComboBox1.Items.Add("Keluar")
    End Sub
    Sub KosongkanData()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
    End Sub

    Public SaldoSekarang As String
    Sub getSaldoUtama()
        bukakoneksi()

        DA = New OdbcDataAdapter("SELECT * from data_dompetku order by Jumlah desc limit 1", CONN)
        DS = New DataSet
        DA.Fill(DS, "data_dompetku")
        Label6.Text = DS.Tables(0).Rows(0).Item(4)
        SaldoSekarang = DS.Tables(0).Rows(0).Item(4)

        TutupKoneksi()
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox3.Text = "" Or DateTimePicker1.Text = "" Or ComboBox1.Text = "" Or TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Silahkan Isi Semua Form")

        Else
            If ComboBox1.Text = "Masuk" Then
                Dim saldobaru As Integer
                Dim saldoMasuk As Integer = CInt(TextBox1.Text)
                Dim saldoTerakhir As Integer = CInt(SaldoSekarang)
                saldobaru = saldoMasuk + saldoTerakhir
                bukakoneksi()
                Dim simpan As String = "insert into data_dompetku (Id,Tanggal,Jenis,Jumlah,saldo_sekarang,Keterangan) value('" & TextBox3.Text & "','" & DateTimePicker1.Text & "','" & ComboBox1.Text & "','" & TextBox1.Text & "','" & saldobaru & "','" & TextBox2.Text & "')"

                CMD = New OdbcCommand(simpan, CONN)
                CMD.ExecuteNonQuery()
                MsgBox("Input data berhasil")
                TampilGrid()
                getSaldoUtama()
                KosongkanData()
                TutupKoneksi()

            ElseIf ComboBox1.Text = "Keluar" Then
                Dim saldobaru As Integer
                Dim saldoKeluar As Integer = CInt(TextBox1.Text)
                Dim saldoTerakhir As Integer = CInt(SaldoSekarang)
                saldobaru = saldoTerakhir - saldoKeluar
                bukakoneksi()
                Dim simpan As String = "insert into data_dompetku (Id,Tanggal,Jenis,Jumlah,saldo_sekarang,Keterangan) value('" & TextBox3.Text & "','" & DateTimePicker1.Text & "','" & ComboBox1.Text & "','" & TextBox1.Text & "','" & saldobaru & "','" & TextBox2.Text & "')"

                CMD = New OdbcCommand(simpan, CONN)
                CMD.ExecuteNonQuery()
                MsgBox("Input data berhasil")
                TampilGrid()
                getSaldoUtama()
                KosongkanData()
                TutupKoneksi()
            End If
        End If
    End Sub

    Private Sub TextBox1_KeyPress1(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        TextBox1.MaxLength = 6
        If e.KeyChar = Chr(13) Then
            bukakoneksi()
            CMD = New OdbcCommand("Select * from data_dompetku where Jumlah ='" & TextBox1.Text & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            If Not RD.HasRows Then
                MsgBox("No Tidak Ada, Silahkan coba lagi!")
                TextBox1.Focus()
            Else
                TextBox3.Text = RD.Item("Id")
                DateTimePicker1.Text = RD.Item("Tanggal")
                ComboBox1.Text = RD.Item("Jenis")
                TextBox1.Text = RD.Item("Jumlah")

                TextBox2.Focus()
            End If
        End If
    End Sub

    Private Sub TextBox3_KeyPress1(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress
        TextBox3.MaxLength = 6
        If e.KeyChar = Chr(13) Then
            bukakoneksi()
            CMD = New OdbcCommand("Select * from data_dompetku where Id ='" & TextBox3.Text & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            If Not RD.HasRows Then
                MsgBox("Id Tidak ada, Silahkan coba lagi!")
                TextBox3.Focus()
            Else
                DateTimePicker1.Text = RD.Item("Tanggal")
                ComboBox1.Text = RD.Item("Jenis")
                TextBox1.Text = RD.Item("Jumlah")
                TextBox2.Text = RD.Item("Keterangan")
                TextBox1.Focus()
            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        bukakoneksi()
        Dim edit As String = "Update data_dompetku set
        Tanggal='" & DateTimePicker1.Text & "',
        Jenis='" & ComboBox1.Text & "',
        Jumlah='" & TextBox1.Text & "',
        Keterangan='" & TextBox2.Text & "'
        where Id='" & TextBox3.Text & "'"

        CMD = New OdbcCommand(edit, CONN)
        CMD.ExecuteNonQuery()
        MsgBox("Data Berhasil diUpdate")
        TampilGrid()
        KosongkanData()
        TutupKoneksi()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If TextBox3.Text = "" Then
        Else
            If MessageBox.Show("Yakin ingin dihapus?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                bukakoneksi()
                Dim hapus As String = "delete From data_dompetku Where Id='" & TextBox3.Text & "'"
                CMD = New OdbcCommand(hapus, CONN)
                CMD.ExecuteNonQuery()
                TampilGrid()
                KosongkanData()
                TutupKoneksi()
            End If
        End If
    End Sub

End Class
