Imports System.IO
Imports System.IO.Packaging
Imports System.Xml
Imports System.Windows.Forms
Imports System.Drawing

Public Class MainForm
    Inherits Form


    Private Sub btnSelectFile_Click(sender As Object, e As EventArgs) Handles btnSelectFile.Click
        Using openFileDialog As New OpenFileDialog()
            openFileDialog.Filter = "Excel Dosyaları (*.xlsx)|*.xlsx|Tüm Dosyalar (*.*)|*.*"
            openFileDialog.Title = "Temizlenecek Excel Dosyasını Seçiniz..."

            If openFileDialog.ShowDialog() = DialogResult.OK Then
                ProcessXLSXFile(openFileDialog.FileName)
            End If
        End Using
    End Sub

    Private Sub AciklamaYaz(metin As String)

        ProgressBar1.Value = Math.Min(ProgressBar1.Value + 20, ProgressBar1.Maximum)
        Aciklama.Text = metin
        Application.DoEvents()
        System.Threading.Thread.Sleep(1000)
    End Sub

    Private Sub ProcessXLSXFile(filePath As String)
        Dim backupPath As String = filePath & ".backup"
        Dim tempDir As String = Path.Combine(Path.GetTempPath(), "xlsx_temp_" & DateTime.Now.ToString("yyyyMMddHHmmss"))

        Try
            ' 1. Önce yedek oluştur
            File.Copy(filePath, backupPath, True)
            AciklamaYaz("Yedek oluşturuldu: " & backupPath)
            ' 2. Geçici dizin oluştur
            Directory.CreateDirectory(tempDir)
            AciklamaYaz("Geçici dizin oluşturuldu: " & tempDir)
            ' 3. XLSX dosyasını aç ve workbook.xml'i çıkar
            Dim workbookXmlPath As String = ExtractWorkbookXML(filePath, tempDir)
            AciklamaYaz("Temizlenecek veri bloğu ayıklandı")
            If File.Exists(workbookXmlPath) Then
                ' 4. XML'i işle
                ProcessWorkbookXML(workbookXmlPath)
                AciklamaYaz("Temizlenecek veri bloğu siliniyor")
                ' 5. Güncellenmiş workbook.xml'i XLSX'e geri yükle
                UpdateWorkbookXML(filePath, workbookXmlPath)
                AciklamaYaz("Temizlenen veri bloğu XLSX dosyasına yazıldı")
                MessageBox.Show("İşlem başarıyla tamamlandı!" & vbCrLf &
                                "Orijinal dosya: " & backupPath & vbCrLf &
                                "İşlenmiş dosya: " & filePath,
                                "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information)
                AciklamaYaz("İşlem Tamam")
            Else
                MessageBox.Show("workbook.xml bulunamadı!", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            MessageBox.Show("Hata: " & ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' 6. Geçici dizini temizle
            If Directory.Exists(tempDir) Then
                Try
                    Directory.Delete(tempDir, True)
                Catch
                    ' Geçici dizin silinemediyse ignore et
                End Try
            End If
        End Try
    End Sub

    Private Function ExtractWorkbookXML(xlsxPath As String, tempDir As String) As String
        Using package As Package = Package.Open(xlsxPath, FileMode.Open, FileAccess.ReadWrite)
            ' Workbook part'ını bul
            Dim workbookUri As New Uri("/xl/workbook.xml", UriKind.Relative)
            Dim workbookPart As PackagePart = package.GetPart(workbookUri)

            If workbookPart IsNot Nothing Then
                Dim workbookXmlPath As String = Path.Combine(tempDir, "workbook.xml")

                ' Workbook.xml'i geçici dizine kopyala
                Using stream As Stream = workbookPart.GetStream()
                    Using fileStream As FileStream = File.Create(workbookXmlPath)
                        stream.CopyTo(fileStream)
                    End Using
                End Using

                Return workbookXmlPath
            End If
        End Using

        Return Nothing
    End Function

    Private Sub ProcessWorkbookXML(workbookXmlPath As String)
        Dim xmlDoc As New XmlDocument()
        xmlDoc.PreserveWhitespace = True ' Boşlukları koru
        xmlDoc.Load(workbookXmlPath)

        ' Namespace manager oluştur
        Dim nsManager As New XmlNamespaceManager(xmlDoc.NameTable)
        nsManager.AddNamespace("ns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")

        ' definedNames node'unu bul
        Dim definedNamesNode As XmlNode = xmlDoc.SelectSingleNode("//ns:definedNames", nsManager)

        If definedNamesNode IsNot Nothing Then
            ' Tüm definedName elementlerini al
            Dim definedNameNodes As XmlNodeList = definedNamesNode.SelectNodes("ns:definedName", nsManager)

            If definedNameNodes IsNot Nothing Then
                ' _xlnm.Print_Area içermeyen definedName'leri sil
                For i As Integer = definedNameNodes.Count - 1 To 0 Step -1
                    Dim definedNameNode As XmlNode = definedNameNodes(i)
                    Dim nameAttribute As XmlAttribute = definedNameNode.Attributes("name")

                    If nameAttribute IsNot Nothing AndAlso Not nameAttribute.Value.Contains("_xlnm.Print_Area") Then
                        definedNamesNode.RemoveChild(definedNameNode)
                    End If
                Next

                ' Eğer definedNames boşsa, tüm definedNames node'unu sil
                If definedNamesNode.ChildNodes.Count = 0 Then
                    Dim parentNode As XmlNode = definedNamesNode.ParentNode
                    If parentNode IsNot Nothing Then
                        parentNode.RemoveChild(definedNamesNode)
                    End If
                End If
            End If
        End If

        ' Değişiklikleri kaydet
        xmlDoc.Save(workbookXmlPath)
    End Sub

    Private Sub UpdateWorkbookXML(xlsxPath As String, updatedWorkbookXmlPath As String)
        Using package As Package = Package.Open(xlsxPath, FileMode.Open, FileAccess.ReadWrite)
            ' Workbook part'ını bul
            Dim workbookUri As New Uri("/xl/workbook.xml", UriKind.Relative)
            Dim workbookPart As PackagePart = package.GetPart(workbookUri)

            If workbookPart IsNot Nothing Then
                ' Güncellenmiş workbook.xml'i yükle
                Using fileStream As FileStream = File.OpenRead(updatedWorkbookXmlPath)
                    Using stream As Stream = workbookPart.GetStream(FileMode.Create, FileAccess.Write)
                        fileStream.CopyTo(stream)
                    End Using
                End Using
            End If
        End Using
    End Sub

    <STAThread>
    Public Shared Sub Main()
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        Application.Run(New MainForm())
    End Sub
End Class