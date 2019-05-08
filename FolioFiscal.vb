Imports System.Net.Mail
Module FolioFiscal
    'Dim ErrorControl As New EventLog
    Sub Main()
        LlenaPolizas() ' con archivos de polizas
        FoliosFactor100()
        EnviaFolioFiscal() 'en finagil
        FoliosAvios() ' en detalle finagil
    End Sub

    Private Sub EnviaError(ByVal Para As String, ByVal Mensaje As String, ByVal Asunto As String)
        If InStr(Mensaje, Asunto) = 0 Then
            Dim Mensage As New MailMessage("InternoBI2008@cmoderna.com", Trim(Para), Trim(Asunto), Mensaje)
            Dim Cliente As New SmtpClient("smtp01.cmoderna.com", 26)
            Try
                Cliente.Send(Mensage)
            Catch ex As Exception
                'ReportError(ex)
            End Try
        Else
            Console.WriteLine("No se ha encontrado la ruta de acceso de la red")
        End If
    End Sub

    'Private Sub ReportError(ByVal ex As Exception)
    '    ErrorControl.WriteEntry(ex.Message, EventLogEntryType.Error)
    'End Sub

    Private Sub EnviaFolioFiscal()
        ' Try
        Console.WriteLine("Iniciando")
        Dim c As Integer = 0
        Dim Serie As String = ""
        Dim Porc As Double = 0
        Dim fecha As Date = Today
        fecha = fecha.AddHours(-172)
        Dim Folios As New FinagilDSTableAdapters.CFDI_EncabezadoTableAdapter
        Dim FoliosFin As New FinagilDSTableAdapters.FoliosFinagilTableAdapter
        Dim Historia As New FinagilDSTableAdapters.HistoriaTableAdapter
        Dim Hisgin As New FinagilDSTableAdapters.HisginTableAdapter
        Dim T As New FinagilDS.CFDI_EncabezadoDataTable
        Dim TT As New FinagilDS.FoliosFinagilDataTable
        Dim total As Integer = 0

        'fecha = "01/03/2017"

        Folios.Fill(T, fecha.Month, fecha.Year)
        total = T.Rows.Count

        'pone el folios fiscal en la historia
        For Each r As FinagilDS.CFDI_EncabezadoRow In T.Rows
            'If r.Factura = "10A150034" Then
            '    r.Factura = "10A150034"
            'End If
            c += 1
            Porc = (c / total) * 100
            Serie = r._27_Serie_Comprobante
            If Mid(r._27_Serie_Comprobante, 1, 2) = "AA" Then Serie = "A"
            'If Mid(r._27_Serie_Comprobante, 1, 1) = "C" Then Serie = "C"
            Historia.FolioFiscal(r.Guid, Serie, r._1_Folio)
            Console.Clear()
            Console.WriteLine("Proceso 1: " & MonthName(fecha.Month) & " " & fecha.Year & " " & Porc.ToString("n2") & "%")
        Next
        c = 0
        Dim Xx As String = fecha.ToString("yyyyMM")
        Dim Factura As String
        FoliosFin.Fill(TT, Xx)
        total = TT.Rows.Count
        'pone el folios fiscal en la hisgin
        For Each rr As FinagilDS.FoliosFinagilRow In TT.Rows
            c += 1
            Porc = (c / total) * 100
            Serie = rr.Serie.Trim
            If Mid(rr.Serie, 1, 2) = "AA" Then Serie = "A"
            If Mid(rr.Serie, 1, 1) = "C" Then Serie = "C"
            Factura = Serie & rr.Numero
            If rr.Numero = 159369 Then
                Porc = (c / total) * 100
            End If
            Hisgin.FolioFiscal(Trim(rr.Cheque) & "-" & rr.FolioFiscal, rr.Anexo, Factura)
            Console.Clear()
            Console.WriteLine("Proceso 2: " & MonthName(fecha.Month) & " " & fecha.Year & " " & Porc.ToString("n2") & "%")
        Next
        'Next
        'Catch ex As Exception
        'EnviaError("Ecacerest@lamoderna.com.mx", ex.Message, "error de Folio Fiscal")
        'End Try
    End Sub

    Private Sub LlenaPolizas()
        Dim HisginX As New FinagilDSTableAdapters.HisginTableAdapter
        Dim D As New System.IO.DirectoryInfo(My.Computer.FileSystem.CurrentDirectory)
        Dim F As System.IO.FileInfo() = D.GetFiles("*.txt")
        Dim f1 As System.IO.StreamReader
        Dim f2 As System.IO.StreamWriter
        Dim Linea As String
        Dim LineaX As String
        Dim Referencia As String
        Dim sFecha As String = ""
        Dim Anexo As String
        '++++++++++++++++++++++++++++++++++++LIGA POLIZAS CON Su FOLIO FISCAL
        For i As Integer = 0 To F.Length - 1
            Console.WriteLine(F(i).FullName)
            f1 = New System.IO.StreamReader(F(i).FullName, Text.Encoding.GetEncoding(1252))
            f2 = New System.IO.StreamWriter(My.Computer.FileSystem.CurrentDirectory & "\Nuevo-" & F(i).Name, False, Text.Encoding.GetEncoding(1252))
            While Not f1.EndOfStream
                Linea = f1.ReadLine

                If Mid(Linea, 1, 1) <> "M" Then
                    f2.WriteLine(Linea)
                    Console.WriteLine(Mid(Linea, 41, 10))
                    sFecha = Trim(Mid(Linea, 4, 8))
                Else
                    Console.WriteLine(Mid(Linea, 35, 10))
                    Anexo = Mid(Linea, 35, 5) & Mid(Linea, 41, 4)
                    Referencia = Trim(Mid(Linea, 101, 100))
                    LineaX = Mid(Linea, 1, 110)
                    If Referencia <> "" Then
                        Referencia = Mid(HisginX.BuscaConcepto(Anexo, sFecha, Referencia), 1, 102)
                    Else
                        Referencia = Space(102)
                    End If
                    LineaX += Referencia
                    LineaX += Mid(Linea, 213, 99)
                    f2.WriteLine(LineaX)
                End If
            End While
            f1.Close()
            f2.Close()
        Next
        '++++++++++++++++++++++++++++++++++++LIGA POLIZAS CON SU FOLIO FISCAL
    End Sub

    Private Sub FoliosAvios()
        Try
            Console.WriteLine("Corrije AVIO")

            Console.WriteLine("Iniciando AVIO")
            Dim c As Integer = 0
            Dim GUID As String = ""
            Dim Folio As String = ""
            Dim Porc As Double = 0
            Dim fecha As Date = Today

            fecha = fecha.AddHours(-172)
            Dim Folios As New FinagilDSTableAdapters.FacturaSinFolioTableAdapter
            Dim DetalleFinagil As New FinagilDSTableAdapters.DetalleFINAGILTableAdapter
            Dim T As New FinagilDS.FacturaSinFolioDataTable
            Dim total As Integer = 0
            Dim Serie As String
            Folios.Fill(T, fecha.ToString("yyyyMM01"))
            total = T.Rows.Count
            For Each r As FinagilDS.FacturaSinFolioRow In T.Rows
                c += 1
                Porc = (c / total) * 100
                Serie = ""
                Folio = ""
                For x As Integer = 1 To r.Factura.Trim.Length
                    If IsNumeric(Mid(r.Factura, x, 1)) Then
                        Folio += Mid(r.Factura, x, 1)
                    Else
                        Serie += Mid(r.Factura, x, 1)
                    End If
                Next
                If Folio = "" Then Folio = "0"
                GUID = Folios.SacaGUID(Folio, Serie)
                DetalleFinagil.UpdateFolio(GUID, r.Factura.Trim)
                Console.Clear()
                Console.WriteLine("Proceso 3 Avio: " & MonthName(fecha.Month) & " " & fecha.Year & " " & Porc.ToString("n2") & "%")

            Next
        Catch ex As Exception
            EnviaError("Ecacerest@lamoderna.com.mx", ex.Message, "error de Folio Fiscal")
        End Try
    End Sub

    Private Sub FoliosFactor100()
        '        Try
        Console.WriteLine("Iniciando")
        Dim c As Integer = 0
        Dim Serie As String = ""
        Dim Porc As Double = 0
        Dim fecha As Date = Today
        fecha = fecha.AddHours(-172)
        Dim Folios As New CFDIdsTableAdapters.FacturaTableAdapter
        Dim Factor100 As New Factor100DSTableAdapters.FACT_EMITIDASTableAdapter
        Dim T As New CFDIds.FacturaDataTable
        Dim total As Integer = 0
        Folios.FillByB(T, fecha.Month, fecha.Year)
        total = T.Rows.Count
        'pone el folios en Factor 100
        For Each r As CFDIds.FacturaRow In T.Rows
            c += 1
            Porc = (c / total) * 100
            Serie = r.Serie
            Serie = "F"
            Factor100.FolioFiscal(r.FolioFiscal, Serie, r.Referencia)
            Console.Clear()
            Console.WriteLine("Proceso 1: " & MonthName(fecha.Month) & " " & fecha.Year & " " & Porc.ToString("n2") & "%")
        Next
        '        Catch ex As Exception
        '        EnviaError("Ecacerest@lamoderna.com.mx", ex.Message, "error de Folio Fiscal")
        '        End Try
    End Sub

End Module
