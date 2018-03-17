Imports iText.Kernel.Pdf
Imports System.IO

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    'Button 1押下
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Label1.Text = getFolderPath()
    End Sub

    'Button 2：押下
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Label2.Text = getFolderPath()
    End Sub

    'フォルダの選択
    Private Function GetFolderPath() As String
        If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            GetFolderPath = FolderBrowserDialog1.SelectedPath
        Else
            GetFolderPath = GetFolderPath()
        End If
    End Function

    'Button 3押下
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim files As IEnumerable(Of String)
        Dim result As String

        If IsSameFolder() = False Then
            files = GetPDFarray(Label1.Text)

            For Each f As String In files
                ListBox1.Items.Add(f)
                result = UpdatePDF(f)
                ListBox2.Items.Add(result)
            Next

            MessageBox.Show("complete")
        End If
    End Sub

    Private Function IsSameFolder() As Boolean
        If Label1.Text = Label2.Text Then
            MessageBox.Show("source = destination")
            IsSameFolder = True
        Else
            IsSameFolder = False
        End If
    End Function

    'フォルダ内のPDFファイルを抽出する。
    Private Function GetPDFarray(folder_path As String) As IEnumerable(Of String)
        Dim search_pattern = "*.pdf"
        Dim files As IEnumerable(Of String)

        files = System.IO.Directory.EnumerateFiles(folder_path, search_pattern)
        GetPDFarray = files
    End Function

    'PDFファイルの設定を変更し、目的のフォルダへコピーする。
    Private Function UpdatePDF(source_file As String) As String
        Dim myReader As New PdfReader(source_file)
        Dim destination_file As String = Label2.Text + "\" + Path.GetFileName(source_file)
        Dim myWriter As New PdfWriter(destination_file)
        Dim myPDF As New PdfDocument(myReader, myWriter)
        Dim myPreference As New PdfViewerPreferences

        myPreference.SetDirection(PdfViewerPreferences.PdfViewerPreferencesConstants.RIGHT_TO_LEFT)
        myPDF.GetCatalog().SetViewerPreferences(myPreference)

        myPDF.Close()
        myWriter.Close()
        myReader.Close()
        UpdatePDF = destination_file
    End Function

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Close()
    End Sub
End Class
