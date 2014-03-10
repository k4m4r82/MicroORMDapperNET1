Imports System.Data.OleDb

Module Module1

    Private Function GetOpenConnection() As OleDbConnection
        Dim conn As OleDbConnection = Nothing

        Dim appDir As String = System.IO.Directory.GetCurrentDirectory()
        Dim strConn As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + appDir + "\SISWA.MDB;User Id=admin;Password=;"

        Try

            conn = New OleDbConnection(strConn)
            conn.Open()

        Catch ex As Exception

        End Try

        Return conn
    End Function

    Private Function GetDataSiswa() As List(Of Siswa)
        Dim daftarSiswa As New List(Of Siswa)

        Using conn As OleDbConnection = GetOpenConnection()

            Dim strSql = "SELECT nis, nama FROM siswa"
            Using cmd = New OleDbCommand(strSql, conn)
                Using dtr As OleDbDataReader = cmd.ExecuteReader()

                    While dtr.Read()
                        Dim siswa As New Siswa()
                        With siswa
                            .Nis = IIf(IsDBNull(dtr("nis")), String.Empty, dtr("nis").ToString())
                            .Nama = IIf(IsDBNull(dtr("nama")), String.Empty, dtr("nama").ToString())
                        End With

                        daftarSiswa.Add(siswa)
                    End While

                End Using
            End Using
        End Using

        Return daftarSiswa
    End Function

    Sub Main()

        'Using conn As OleDbConnection = GetOpenConnection()

        '    Dim strSql = "SELECT nis, nama FROM siswa"
        '    Using cmd = New OleDbCommand(strSql, conn)
        '        Using dtr As OleDbDataReader = cmd.ExecuteReader()

        '            Console.WriteLine("NIS" & vbTab & "NAMA")
        '            Console.WriteLine("===================================")
        '            While dtr.Read()
        '                Console.WriteLine(dtr("nis") & vbTab & dtr("nama"))
        '            End While

        '        End Using
        '    End Using
        'End Using

        Console.WriteLine("NIS" & vbTab & "NAMA")
        Console.WriteLine("===================================")

        Dim daftarSiswa = GetDataSiswa()

        For Each siswa In daftarSiswa
            Console.WriteLine(siswa.Nis & vbTab & siswa.Nama)
        Next

        Console.ReadKey()

    End Sub

End Module
