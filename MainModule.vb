Imports System.IO
Imports System.Net
Imports System.Security.Policy
Imports System.Threading
Imports System.Web.Script.Serialization

Imports RestSharp

Public Enum NamesJSON As Integer
   ParentCategories = 1
   ChildCategories = 2
   UploadBrands = 3
   UploadItems = 4
   UploadVariations = 5
End Enum

Module MainModule
   Private htParentCatJSON As New Hashtable
   Private htChildCatJSON As New Hashtable
   Private htBrandsJSON As New Hashtable
   Private htUploadItemJSON As New Hashtable
   Private htUploadVariationJSON As New Hashtable

   Private iParentCatFiles As Integer = 0
   Private iChildCatFiles As Integer = 0
   Private iBrandFiles As Integer = 0
   Private iItemFiles As Integer = 0
   Private iVariationFiles As Integer = 0

   Private iRowLimit As Integer = CInt(GetValueFromConfig("RowLimit"))

   Sub Main(ByVal args() As String)
      'HACK: per debuggarlo, giustamente, bisogna passargli un parametro
      '#If DEBUG Then
      '       ReDim args(0)
      '      args(0) = "/Export"
      '#End If
      If args.Length >= 1 Then
         For Each par As String In args
            If par.StartsWith("/?") Then
               ShowInfo()

            ElseIf par.StartsWith("/TestErr") Then
               TestErr()
            ElseIf par.StartsWith("/ExportParent") Then
               UploadParentCategories()

               CallLogJSON(NamesJSON.ParentCategories, htParentCatJSON)

            ElseIf par.StartsWith("/ExportChild") Then
               UploadChildCategories()

               CallLogJSON(NamesJSON.ChildCategories, htChildCatJSON)

            ElseIf par.StartsWith("/ExportBrands") Then
               UploadBrands()

               CallLogJSON(NamesJSON.UploadBrands, htBrandsJSON)

            ElseIf par.StartsWith("/ExportProducts") Then
               UploadProducts()

               CallLogJSON(NamesJSON.UploadItems, htUploadItemJSON)

            ElseIf par.StartsWith("/ExportVariations") Then
               UploadVariations()

               CallLogJSON(NamesJSON.UploadVariations, htUploadVariationJSON)

            ElseIf par.StartsWith("/ExportLight") Then
               StartExport(True)

            ElseIf par.StartsWith("/Export") Then
               StartExport()

            ElseIf par.StartsWith("/SendFiles") Then
               If CBool(GetValueFromConfig("SendFilesToWS")) Then
                  Dim sUrl As String = GetValueFromConfig("urlBase") & GetValueFromConfig("urlPath")
                  SendFiles(sUrl)
                  Dim sUrl2 As String = GetValueFromConfig("urlBase2") & GetValueFromConfig("urlPath2")
                  SendFiles(sUrl2)
               Else
                  Console.WriteLine("Invio files non abilitato!")
               End If

            Else
               ShowInfo()

            End If
         Next
      End If
   End Sub

   Private Sub TestErr()
      Try
         Dim i As Integer = "a"
      Catch ex As Exception
         ThrowException("TEST ERRORI:" & ex.ToString, True)
      End Try

      If IsAppInError And CInt(GetValueFromConfig("enable_err_mail")) = 1 Then
         SendEmail(GetValueFromConfig("err_mail_to_address"), GetValueFromConfig("err_mail_cc_address"), GetValueFromConfig("err_mail_ccn_address"), "Errore Connettore PrestaShop", Util.sError, False)
      End If
   End Sub

   Private Sub ShowInfo()
      Dim sInfo As String = ""

      sInfo &= "Elenco comandi accettati:" & vbNewLine
      sInfo &= "/TestErr              --> Testa l'invio di email di errore" & vbNewLine
      sInfo &= "/Export               --> Esporta tutti i dati" & vbNewLine
      sInfo &= "/ExportLight          --> Esporta tutti i dati, di cui, solo articoli movimentati nelle ultime 24h" & vbNewLine
      sInfo &= "/ExportParent         --> Esporta tutte le macro-categorie" & vbNewLine
      sInfo &= "/ExportChild          --> Esporta tutte le sotto-categorie" & vbNewLine
      sInfo &= "/ExportBrands         --> Esporta tutte le marche" & vbNewLine
      sInfo &= "/ExportProducts       --> Esporta tutti i prodotti" & vbNewLine
      sInfo &= "/ExportVariations     --> Esporta tutte le taglie" & vbNewLine
      sInfo &= "/SendFiles            --> Invia i file senza rigenerarli" & vbNewLine
      sInfo &= "/?                    --> Mostra l'elenco dei comandio accettati"

      Console.WriteLine(sInfo)
   End Sub

   Private Sub StartExport(Optional ByVal Light As Boolean = False)
      Dim dataInizio As DateTime = Now

      Dim sJsonPath As String = TerminatePath(GetValueFromConfig("PathJsonFiles"))

      For Each fileName As String In Directory.GetFiles(sJsonPath)
         File.Delete(fileName)
      Next

      UploadParentCategories()

      CallLogJSON(NamesJSON.ParentCategories, htParentCatJSON)

      UploadChildCategories()

      CallLogJSON(NamesJSON.ChildCategories, htChildCatJSON)

      UploadBrands()

      CallLogJSON(NamesJSON.UploadBrands, htBrandsJSON)

      UploadProducts(Light)

      CallLogJSON(NamesJSON.UploadItems, htUploadItemJSON)

      UploadVariations(Light)

      CallLogJSON(NamesJSON.UploadVariations, htUploadVariationJSON)

      If CBool(GetValueFromConfig("SendFilesToWS")) Then
         Dim sUrl As String = GetValueFromConfig("urlBase") & GetValueFromConfig("urlPath")
         SendFiles(sUrl)
         Dim sUrl2 As String = GetValueFromConfig("urlBase2") & GetValueFromConfig("urlPath2")
         SendFiles(sUrl2)
      Else
         SendFinalReport(dataInizio)
      End If
   End Sub

   Private Sub CallLogJSON(ByVal JSONName As Integer, ByVal Value As Hashtable)
      Dim iNumberOfFiles As Integer = 0

      Select Case JSONName
         Case NamesJSON.ParentCategories
            iNumberOfFiles = iParentCatFiles

         Case NamesJSON.ChildCategories
            iNumberOfFiles = iChildCatFiles

         Case NamesJSON.UploadBrands
            iNumberOfFiles = iBrandFiles

         Case NamesJSON.UploadItems
            iNumberOfFiles = iItemFiles

         Case NamesJSON.UploadVariations
            iNumberOfFiles = iVariationFiles

      End Select

      For i As Integer = 1 To iNumberOfFiles
         Using JSONLogger As New LogJSON(JSONName, i)
            If i = iNumberOfFiles And Right(Value(i), 1) <> "]" And Right(Value(i), 1) = "," Then Value(i) = Left(Value(i), Len(Value(i)) - 1) & "]"

            JSONLogger.AddRow(Value(i))
         End Using
      Next
   End Sub

   Private Sub SendFiles(ByVal sUrl As String)
      Dim dataInizio As DateTime = Now

      ''Dim sUrl As String = GetValueFromConfig("urlBase") & GetValueFromConfig("urlPath")
      Dim sJsonPath As String = TerminatePath(GetValueFromConfig("PathJsonFiles"))

      For Each fileName As String In Directory.GetFiles(sJsonPath)
         Try
            Dim htQS As New Hashtable

            htQS.Add("key", GetValueFromConfig("Key"))

            Select Case True
               Case fileName.StartsWith(LogJSON.getFullPathFromFileName(NamesJSON.ParentCategories & "_" & LogJSON.getFileName(NamesJSON.ParentCategories)))
                  htQS.Add("file_json_type", NamesJSON.ParentCategories)

               Case fileName.StartsWith(LogJSON.getFullPathFromFileName(NamesJSON.ChildCategories & "_" & LogJSON.getFileName(NamesJSON.ChildCategories)))
                  htQS.Add("file_json_type", NamesJSON.ChildCategories)

               Case fileName.StartsWith(LogJSON.getFullPathFromFileName(NamesJSON.UploadBrands & "_" & LogJSON.getFileName(NamesJSON.UploadBrands)))
                  htQS.Add("file_json_type", NamesJSON.UploadBrands)

               Case fileName.StartsWith(LogJSON.getFullPathFromFileName(NamesJSON.UploadItems & "_" & LogJSON.getFileName(NamesJSON.UploadItems)))
                  htQS.Add("file_json_type", NamesJSON.UploadItems)

               Case fileName.StartsWith(LogJSON.getFullPathFromFileName(NamesJSON.UploadVariations & "_" & LogJSON.getFileName(NamesJSON.UploadVariations)))
                  htQS.Add("file_json_type", NamesJSON.UploadVariations)

            End Select

            htQS.Add("only_new_items", GetValueFromConfig("onlyNewItems"))

            Dim client As New RestClient(sUrl & "?" & FromHashTableToQueryString(htQS))
            Dim request As New RestRequest(Method.POST)

            ServicePointManager.SecurityProtocol = CType((&HC0 Or &H300 Or &HC00), SecurityProtocolType)

            Using sr As New StreamReader(fileName)
               Dim sJSON As String = sr.ReadToEnd

               request.AddHeader("Accept", "application/json")
               request.AddHeader("Cache-Control", "no-cache")

               request.RequestFormat = DataFormat.Json
               request.AlwaysMultipartFormData = True

               request.AddFile("file_json", fileName, "application/json")

               '5 minuti di timeout (basteranno?)
               client.Timeout = CInt(GetValueFromConfig("RequestTimeout"))
               request.Timeout = CInt(GetValueFromConfig("RequestTimeout"))

               Dim response As RestResponse = client.Execute(request)
               Dim responseCode As Integer = response.StatusCode

               If IsEmpty(response.ErrorMessage) Then
                  Dim oResult As Object
                  Dim Deserializer As New JavaScriptSerializer()
                  Dim JSON As New Dictionary(Of String, Object)()

                  Deserializer.MaxJsonLength = 8388608 '16 MB

                  oResult = Deserializer.DeserializeObject(response.Content)
                  JSON = DirectCast(oResult, Dictionary(Of String, Object))

                  If JSON("response").ToString <> "success" Then
                     Dim sErrs As String = ""

                     For i As Integer = 0 To UBound(JSON("response"))
                        Dim sErr As String = JSON("response")(i)

                        sErrs += If(i = UBound(JSON("response")), sErr, sErr & vbNewLine)
                     Next

                     If Not IsEmpty(sErrs) Then
                        Throw New Exception("Errore nella procedura SendFiles:" & vbNewLine &
                                                        "File: " & fileName & vbNewLine &
                                                        "URL: " & sUrl & "?" & FromHashTableToQueryString(htQS) & vbNewLine &
                                                        "RequestContent: " & sJSON & vbNewLine &
                                                        "ResponseCode: " & responseCode.ToString & vbNewLine &
                                                        "ResponseBody: " & response.Content & vbNewLine & vbNewLine &
                                                        "Clear Errors: " & vbNewLine & sErrs)
                     End If
                  End If
               Else
                  Throw New Exception("Errore nella procedura SendFiles:" & vbNewLine &
                                            "File: " & fileName & vbNewLine &
                                            "URL: " & sUrl & "?" & FromHashTableToQueryString(htQS) & vbNewLine &
                                            "RequestContent: " & sJSON & vbNewLine & vbNewLine &
                                            "Message: " & response.ErrorMessage)
               End If
            End Using
         Catch ex As Exception
            ThrowException("Errore nella procedura SendFiles (errore .Net):" & vbNewLine & ex.ToString & " Url: " & sUrl, True)
         End Try

         Thread.Sleep(GetValueFromConfig("CooldownCall"))
      Next

      SendFinalReport(dataInizio)
   End Sub

   Private Sub SendFinalReport(ByVal dataInizio As DateTime)
        Dim sMessage As String = ""

        sMessage &= "L'invio dei dati alla piattaforma PrestaShop è iniziato alle " & dataInizio.ToString & " ed è terminato alle " & Now.ToString

        SendEmail(GetValueFromConfig("end_mail_to_address"), GetValueFromConfig("end_mail_cc_address"), GetValueFromConfig("end_mail_ccn_address"), "Fine invio dati a PrestaShop", sMessage, False)

        If IsAppInError Then
            SendEmail(GetValueFromConfig("err_mail_to_address"), GetValueFromConfig("err_mail_cc_address"), GetValueFromConfig("err_mail_ccn_address"), "Errore Connettore PrestaShop", Util.sError, False)
        End If
    End Sub

    Private Sub UploadParentCategories()
        Dim iRowsMade As Integer = 0

        Dim hSQL As New SQL(getConnectionString)
        Dim dtMacroCategorie As New DataTable

        Dim sSql As String = "SELECT * FROM bru_TabCategorie"

        If hSQL.OttieniValori(sSql, dtMacroCategorie) Then
            If dtMacroCategorie.Rows.Count > 0 Then
                For Each row As DataRow In dtMacroCategorie.Rows
                    Try
                        iRowsMade += 1

                        Dim sJSON As String = ""

                        sJSON &= "{"

                        addToJSON(sJSON, "id", row.Item("ID_ECWID").ToString, "S", True)
                        addToJSON(sJSON, "name", row.Item("Descrizione").ToString, "S", True)
                        addToJSON(sJSON, "description", row.Item("Descrizione").ToString, "S", True)
                        addToJSON(sJSON, "enabled", True, "B", False)

                        sJSON &= "},"

                        If iRowsMade = 1 Then
                            iParentCatFiles += 1
                            htParentCatJSON.Add(iParentCatFiles, "[" & sJSON)

                        ElseIf iRowsMade < iRowLimit Then
                            htParentCatJSON(iParentCatFiles) &= sJSON

                        ElseIf iRowsMade = iRowLimit Then
                            htParentCatJSON(iParentCatFiles) = Left(htParentCatJSON(iParentCatFiles), Len(htParentCatJSON(iParentCatFiles)) - 1) & "]"
                            iRowsMade = 0

                        End If
                    Catch ex As Exception
                        Exit For

                        ThrowException("Errore nella procedura UploadParentCategories (errore .Net):" & vbNewLine & ex.ToString, True)
                    End Try
                Next
            End If
        Else
            ThrowException("Errore nella procedura UploadParentCategories (errore hSQL):" & vbNewLine & hSQL.LeggiErrore)
        End If

        hSQL.Dispose()
        dtMacroCategorie.Dispose()
    End Sub

    Private Sub UploadChildCategories()
        Dim iRowsMade As Integer = 0

        Dim hSQL As New SQL(getConnectionString)
        Dim dtMacroCategorie As New DataTable

        Dim sSql As String

        sSql = "SELECT DISTINCT bru_ArticoliCategorieClassi.MacroCategoria, bru_ArticoliCategorieClassi.Classe, bru_ArticoliCategorieClassi.Linea, bru_ArticoliCategorieClassi.Gruppo, bru_Tabclas.Descrizione, bru_TabCategorie.ID_ECWID AS Parent_ID, bru_TabCategorieClassi.ID_ECWID AS Child_ID" & vbNewLine
        sSql &= "FROM bru_ArticoliCategorieClassi" & vbNewLine
        sSql &= "LEFT JOIN bru_TabCategorie" & vbNewLine
        sSql &= "ON bru_TabCategorie.Codice = bru_ArticoliCategorieClassi.MacroCategoria" & vbNewLine
        sSql &= "LEFT JOIN bru_TabCategorieClassi" & vbNewLine
        sSql &= "ON bru_ArticoliCategorieClassi.MacroCategoria = bru_TabCategorieClassi.MacroCategoria" & vbNewLine
        sSql &= "AND bru_ArticoliCategorieClassi.Classe = bru_TabCategorieClassi.Classe" & vbNewLine
        sSql &= "AND bru_ArticoliCategorieClassi.Linea = bru_TabCategorieClassi.Linea" & vbNewLine
        sSql &= "AND bru_ArticoliCategorieClassi.Gruppo = bru_TabCategorieClassi.Gruppo" & vbNewLine
        sSql &= "LEFT JOIN bru_Tabclas" & vbNewLine
        sSql &= "ON bru_ArticoliCategorieClassi.Classe = bru_Tabclas.Classe" & vbNewLine
        sSql &= "AND bru_ArticoliCategorieClassi.Linea = bru_Tabclas.Linea" & vbNewLine
        sSql &= "AND bru_ArticoliCategorieClassi.Gruppo = bru_Tabclas.Gruppo"

        If hSQL.OttieniValori(sSql, dtMacroCategorie) Then
            If dtMacroCategorie.Rows.Count > 0 Then
                For Each row As DataRow In dtMacroCategorie.Rows
                    Try
                        iRowsMade += 1

                        Dim sJSON As String = ""

                        sJSON &= "{"

                        addToJSON(sJSON, "id", row.Item("Child_ID").ToString, "I", True)
                        addToJSON(sJSON, "name", row.Item("Descrizione").ToString, "S", True)
                        addToJSON(sJSON, "description", row.Item("Descrizione").ToString, "S", True)
                        addToJSON(sJSON, "enabled", True, "B", True)
                        addToJSON(sJSON, "parentId", row.Item("Parent_ID").ToString, "I", False)

                        sJSON &= "},"

                        If iRowsMade = 1 Then
                            iChildCatFiles += 1
                            htChildCatJSON.Add(iChildCatFiles, "[" & sJSON)

                        ElseIf iRowsMade < iRowLimit Then
                            htChildCatJSON(iChildCatFiles) &= sJSON

                        ElseIf iRowsMade = iRowLimit Then
                            htChildCatJSON(iChildCatFiles) = Left(htChildCatJSON(iChildCatFiles), Len(htChildCatJSON(iChildCatFiles)) - 1) & "]"
                            iRowsMade = 0

                        End If
                    Catch ex As Exception
                        Exit For

                        ThrowException("Errore nella procedura UploadChildCategories (errore .Net):" & vbNewLine & ex.ToString, True)
                    End Try
                Next
            End If
        Else
            ThrowException("Errore nella procedura UploadChildCategories (errore hSQL):" & vbNewLine & hSQL.LeggiErrore)
        End If

        hSQL.Dispose()
        dtMacroCategorie.Dispose()
    End Sub

    Private Sub UploadBrands()
        Dim iRowsMade As Integer = 0

        Dim hSQL As New SQL(getConnectionString)
        Dim dtMarchi As New DataTable

        Dim sSql As String

        sSql = "SELECT * FROM bru_TabBrand"

        If hSQL.OttieniValori(sSql, dtMarchi) Then
            If dtMarchi.Rows.Count > 0 Then
                For Each row As DataRow In dtMarchi.Rows
                    Try
                        iRowsMade += 1

                        Dim sJSON As String = ""

                        sJSON &= "{"

                        addToJSON(sJSON, "sku", row.Item("Codice").ToString, "S", True)
                        addToJSON(sJSON, "name", row.Item("Descrizione").ToString, "S", True)
                        addToJSON(sJSON, "description", row.Item("Descrizione").ToString, "S", True)
                        addToJSON(sJSON, "enabled", True, "B", False)

                        sJSON &= "},"

                        If iRowsMade = 1 Then
                            iBrandFiles += 1
                            htBrandsJSON.Add(iBrandFiles, "[" & sJSON)

                        ElseIf iRowsMade < iRowLimit Then
                            htBrandsJSON(iBrandFiles) &= sJSON

                        ElseIf iRowsMade = iRowLimit Then
                            htBrandsJSON(iBrandFiles) = Left(htBrandsJSON(iBrandFiles), Len(htBrandsJSON(iBrandFiles)) - 1) & "]"
                            iRowsMade = 0

                        End If
                    Catch ex As Exception
                        Exit For

                        ThrowException("Errore nella procedura UploadBrands (errore .Net):" & vbNewLine & ex.ToString, True)
                    End Try
                Next
            End If
        End If

        hSQL.Dispose()
        dtMarchi.Dispose()
    End Sub

    Private Sub UploadProducts(Optional ByVal Light As Boolean = False)
        Dim iRowsMade As Integer = 0

        Dim hSQL As New SQL(getConnectionString)
        Dim dtProdotti As New DataTable

        Dim sSql As String = "SELECT CodArt, InviaWeb, P_ID_ECWID FROM mas_Articoli WHERE (InviaWeb = 1 OR (ISNULL(InviaWeb, 0) = 0 AND ISNULL(P_ID_ECWID, '') <> ''))"
        If Light Then
            sSql = sSql & " AND CodArt IN ("
            sSql = sSql & GetInFiltroLight()
            sSql = sSql & ") "
        End If
        If hSQL.OttieniValori(sSql, dtProdotti) Then
            If dtProdotti.Rows.Count > 0 Then
                For Each rowArt As DataRow In dtProdotti.Rows
                    Dim sCodArt As String = rowArt.Item("CodArt").ToString
                    Dim bIsEnabled As Boolean = CBool(rowArt.Item("InviaWeb").ToString)

                    Dim dtProdotto As New DataTable
                    Dim dtCategorie As New DataTable

                    sSql = "SELECT mas_Articoli.CodArt, mas_Articoli.Descrizione, p_MacroClasseWeb AS MacroClasse," & vbNewLine
                    sSql &= "CASE WHEN mas_Articoli.p_DataInizioSconto <= CONVERT(date, GETDATE()) AND mas_Articoli.p_DataFineSconto >= CONVERT(date, GETDATE()) AND mas_Articoli.p_Sconto > 0 " & vbNewLine
                    sSql &= "THEN mas_Articoli.Prezzo_Dettaglio - ((mas_Articoli.Prezzo_Dettaglio * mas_Articoli.p_Sconto) / 100) " & vbNewLine
                    sSql &= "ELSE mas_Articoli.Prezzo_Dettaglio END AS Prezzo_Dettaglio_Scontato," & vbNewLine
                    sSql &= "ISNULL(mas_Articoli.Prezzo_Dettaglio, 0) AS Prezzo_Dettaglio," & vbNewLine
                    sSql &= "ISNULL(mas_Articoli.p_Sconto, 0) AS PercSconto," & vbNewLine
                    sSql &= "CONVERT(varchar, mas_Articoli.p_DataInizioSconto, 111) AS DataInizioSconto," & vbNewLine
                    sSql &= "CONVERT(varchar, mas_Articoli.p_DataFineSconto, 111) AS DataFineSconto," & vbNewLine
               sSql &= "mas_Articoli.Prezzo_Dettaglio," & vbNewLine
               sSql &= "mas_Articoli.Costo_ultimo," & vbNewLine
               sSql &= "CASE WHEN ISNULL(mas_Articoli.p_ClasseWeb, '')  = '' THEN mas_TabClas.P_ClasseECWID ELSE mas_Articoli.p_ClasseWeb END AS ClasseWeb," & vbNewLine
                    sSql &= "CASE WHEN ISNULL(mas_Articoli.p_ClasseWeb, '')  = '' THEN mas_TabClas.P_LineaECWID ELSE mas_Articoli.p_LineaWeb END AS LineaWeb," & vbNewLine
                    sSql &= "CASE WHEN ISNULL(mas_Articoli.p_ClasseWeb, '')  = '' THEN mas_TabClas.P_GruppoECWID ELSE mas_Articoli.p_GruppoWeb END AS GruppoWeb," & vbNewLine
                    sSql &= "ISNULL(bru_TabBrand.Codice, '') AS Marca," & vbNewLine
                    sSql &= "ISNULL(p_BarCode, '') AS p_BarCode," & vbNewLine
                    sSql &= "ISNULL(Stagione, '') AS Stagione" & vbNewLine
                    sSql &= "FROM mas_Articoli  " & vbNewLine
                    sSql &= "LEFT JOIN mas_TabClas ON mas_TabClas.Classe = mas_Articoli.Classe And mas_TabClas.Linea = mas_Articoli.Linea And mas_TabClas.Gruppo = mas_Articoli.Gruppo" & vbNewLine
                    sSql &= "LEFT JOIN bru_TabBrand ON bru_TabBrand.Codice = mas_Articoli.p_Marca" & vbNewLine
                    sSql &= "WHERE mas_Articoli.CodArt = '" & sCodArt & "'"

                    If hSQL.OttieniValori(sSql, dtProdotto) Then
                        If dtProdotto.Rows.Count > 0 Then
                            For Each row As DataRow In dtProdotto.Rows
                                Try
                                    iRowsMade += 1

                                    Dim sJSON As String = ""

                                    sJSON &= "{"

                                    addToJSON(sJSON, "sku", row.Item("CodArt").ToString, "S", True)
                                    addToJSON(sJSON, "name", row.Item("Descrizione").ToString, "S", True)
                                    addToJSON(sJSON, "unlimited", True, "B", True)

                                    If Not IsEmpty(row.Item("Stagione").ToString) Then
                                        sJSON &= """attributes"":[{"

                                        addToJSON(sJSON, "name", "Stagione", "S", True)
                                        addToJSON(sJSON, "value", row.Item("Stagione").ToString, "S", True)
                                        addToJSON(sJSON, "show", "NOTSHOW", "S", False)

                                        sJSON &= "}],"
                                    End If

                           'Remmato 12/03/21 a seguito di mail di Jacopo Bruno
                           'addToJSON(sJSON, "price", row.Item("Prezzo_Dettaglio_Scontato").ToString, "I", True)
                           addToJSON(sJSON, "price", row.Item("Prezzo_Dettaglio").ToString, "I", True)
                           'richiesta di Jacopo del 10/02/2026
                           addToJSON(sJSON, "finalCost", row.Item("Costo_ultimo").ToString, "I", True)

                           'addToJSON(sJSON, "compareToPrice", row.Item("Prezzo_Dettaglio").ToString, "I", True)

                           addToJSON(sJSON, "discount", row.Item("PercSconto").ToString, "I", True)
                                    addToJSON(sJSON, "discountStartDate", row.Item("DataInizioSconto").ToString, "D", True)
                                    addToJSON(sJSON, "discountEndDate", row.Item("DataFineSconto").ToString, "D", True)

                                    If Not IsEmpty(row.Item("Marca").ToString) Then
                                        addToJSON(sJSON, "Marca", row.Item("Marca").ToString, "S", True)
                                    End If

                                    If Not IsEmpty(row.Item("p_BarCode").ToString) Then
                                        addToJSON(sJSON, "UPC", row.Item("p_BarCode").ToString, "S", True)
                                    End If

                                    sJSON &= """shipping"":{"
                                    addToJSON(sJSON, "type", "GLOBAL_METHODS", "S", True)
                                    addToJSON(sJSON, "methodMarkup", "0", "I", True)
                                    addToJSON(sJSON, "flatRate", "0", "I", True)
                                    sJSON &= """disabledMethods"":[],""enabledMethods"":[]},""categoryIds"":["

                                    sSql = "SELECT bru_TabCategorieClassi.ID_ECWID" & vbNewLine
                                    sSql &= "FROM bru_ArticoliCategorieClassi" & vbNewLine
                                    sSql &= "INNER JOIN bru_TabCategorieClassi" & vbNewLine
                                    sSql &= "ON bru_TabCategorieClassi.MacroCategoria = bru_ArticoliCategorieClassi.MacroCategoria" & vbNewLine
                                    sSql &= "AND bru_TabCategorieClassi.Classe = bru_ArticoliCategorieClassi.Classe" & vbNewLine
                                    sSql &= "AND bru_TabCategorieClassi.Linea = bru_ArticoliCategorieClassi.Linea" & vbNewLine
                                    sSql &= "AND bru_TabCategorieClassi.Gruppo = bru_ArticoliCategorieClassi.Gruppo" & vbNewLine
                                    sSql &= "WHERE CodArt = '" & row.Item("CodArt").ToString & "'"

                                    If hSQL.OttieniValori(sSql, dtCategorie) Then
                                        If dtCategorie.Rows.Count > 0 Then
                                            For Each rowCat As DataRow In dtCategorie.Rows
                                                sJSON &= rowCat.Item("ID_ECWID").ToString & ","
                                            Next

                                            sJSON = Left(sJSON, Len(sJSON) - 1)
                                        End If
                                    Else
                                        ThrowException("Errore nella procedura UploadProducts (errore hSQL):" & hSQL.LeggiErrore)
                                    End If

                                    sJSON &= "],"

                                    addToJSON(sJSON, "enabled", bIsEnabled, "B", False)

                                    sJSON &= "},"

                                    If iRowsMade = 1 Then
                                        iItemFiles += 1
                                        htUploadItemJSON.Add(iItemFiles, "[" & sJSON)

                                    ElseIf iRowsMade < iRowLimit Then
                                        htUploadItemJSON(iItemFiles) &= sJSON

                                    ElseIf iRowsMade = iRowLimit Then
                                        htUploadItemJSON(iItemFiles) = Left(htUploadItemJSON(iItemFiles), Len(htUploadItemJSON(iItemFiles)) - 1) & "]"
                                        iRowsMade = 0

                                    End If
                                Catch ex As Exception
                                    ThrowException("Errore nella procedura UploadProducts (errore .Net):" & vbNewLine & ex.ToString, True)
                                End Try
                            Next
                        End If
                    Else
                        ThrowException("Errore nella procedura UploadProducts (errore hSQL):" & hSQL.LeggiErrore)
                    End If

                    dtProdotto.Dispose()
                    dtCategorie.Dispose()
                Next
            End If
        Else
            ThrowException("Errore nella procedura UploadProducts (errore hSQL):" & hSQL.LeggiErrore)
        End If

        hSQL.Dispose()
        dtProdotti.Dispose()
    End Sub

    Private Sub UploadVariations(Optional ByVal Light As Boolean = False)
        Dim iRowsMade As Integer = 0

        Dim hSQL As New SQL(getConnectionString)
        Dim dtQtaPerTaglie As New DataTable

        Dim sSql As String = ""

        sSql &= "SELECT mas_Articoli.CodArt, mas_Articoli.P_ID_ECWID," & vbNewLine
        sSql &= "CASE WHEN mas_Articoli.p_DataInizioSconto <= CONVERT(date, GETDATE()) AND mas_Articoli.p_DataFineSconto >= CONVERT(date, GETDATE()) AND mas_Articoli.p_Sconto > 0 " & vbNewLine
        sSql &= "THEN mas_Articoli.Prezzo_Dettaglio - ((mas_Articoli.Prezzo_Dettaglio * mas_Articoli.p_Sconto) / 100) " & vbNewLine
        sSql &= "ELSE mas_Articoli.Prezzo_Dettaglio END AS Prezzo_Dettaglio_Scontato," & vbNewLine
        sSql &= "ISNULL(mas_Articoli.p_Sconto, 0) AS PercSconto," & vbNewLine
        sSql &= "CONVERT(varchar, mas_Articoli.p_DataInizioSconto, 111) AS DataInizioSconto," & vbNewLine
        sSql &= "CONVERT(varchar, mas_Articoli.p_DataFineSconto, 111) AS DataFineSconto," & vbNewLine
        sSql &= "ISNULL(mas_Articoli.Prezzo_Dettaglio, 0) AS Prezzo_Dettaglio," & vbNewLine

        For i As Integer = 1 To 20
            sSql &= "CASE WHEN ISNULL(mas_Articoli.p_Tg" & i.ToString & ", '') = '' THEN mas_TabTaglie.Tg" & i.ToString & " ELSE mas_Articoli.p_Tg" & i.ToString & " END AS Tg" & i.ToString & "," & vbNewLine
        Next

        'Mattia agguinto 27/09/22
        For i As Integer = 1 To 20
            sSql &= "mas_Articoli.p_BarCodeTg" & i.ToString & "," & vbNewLine
        Next

        For i As Integer = 1 To 20
            sSql &= "ISNULL(mag_Magazzini.GiacenzaIniziale_Tg" & i.ToString & ", 0) + ISNULL(mag_Magazzini.Carichi_Tg" & i.ToString & ", 0) - ISNULL(mag_Magazzini.Scarichi_Tg" & i.ToString & ", 0) AS Giacenza_Tg" & i.ToString & "," & vbNewLine
        Next

        sSql &= "isNull(mag_magazzini.P_ID_ECWID, '') AS ID_ECWID," & vbNewLine
        sSql &= "mas_Articoli.P_ID_ECWID AS ART_ID_ECWID" & vbNewLine
        sSql &= "FROM mas_Articoli" & vbNewLine
        sSql &= "LEFT JOIN mas_TabTaglie ON mas_Articoli.Tipo_Taglia = mas_TabTaglie.Codice" & vbNewLine
        sSql &= "LEFT JOIN mag_Magazzini ON mas_Articoli.CodArt = mag_Magazzini.CodArt" & vbNewLine
        'REMMATO - 19/07/2021 - Richiesta di Jacopo Bruno per alleggerire il flusso di scambio
        'sSql &= "WHERE mas_Articoli.Taglie = 1 AND (mas_Articoli.InviaWeb = 1 OR (ISNULL(mas_Articoli.InviaWeb, 0) = 0 AND ISNULL(mas_Articoli.P_ID_ECWID, '') <> ''))"
        sSql &= "WHERE mas_Articoli.Taglie = 1 AND mas_Articoli.InviaWeb = 1"
        If Light Then
            sSql = sSql & " AND mas_Articoli.CodArt IN ("
            sSql = sSql & GetInFiltroLight()
            sSql = sSql & ") "
        End If

        If hSQL.OttieniValori(sSql, dtQtaPerTaglie) Then
            If dtQtaPerTaglie.Rows.Count > 0 Then
                For Each row As DataRow In dtQtaPerTaglie.Rows
                    For i As Integer = 1 To 20
                        If Not IsEmpty(row.Item("Tg" & i.ToString).ToString) Then
                            If row.Item("Tg" & i.ToString).ToString <> "@" Then 'Refuso (se ci sono giacenze negative non le porterebbe a 0) --> And CDbl(row.Item("Giacenza_Tg" & i.ToString).ToString) >= 0 Then
                                Try
                                    iRowsMade += 1

                                    Dim sJSON As String

                                    Dim sCodArt As String = row.Item("CodArt").ToString
                                    Dim sTgX As String = row.Item("Tg" & i.ToString).ToString
                                    Dim iGiacenza As Integer = row.Item("Giacenza_Tg" & i.ToString).ToString
                                    Dim iPrezzoScontato As Double = CDbl(row.Item("Prezzo_Dettaglio_Scontato").ToString)
                                    Dim iPrezzoOriginale As Double = CDbl(row.Item("Prezzo_Dettaglio").ToString)

                                    sJSON = "{""options"":[{"

                                    addToJSON(sJSON, "name", "TAGLIA", "S", True)
                                    addToJSON(sJSON, "value", sTgX, "S", False)

                                    sJSON &= "}],"

                                    addToJSON(sJSON, "sku", sTgX.Replace(" ", "_"), "S", True)
                                    addToJSON(sJSON, "parent_sku", sCodArt, "S", True)

                                    If iGiacenza > 0 Then
                                        addToJSON(sJSON, "quantity", iGiacenza.ToString, "I", True)
                                    Else
                                        addToJSON(sJSON, "quantity", 0.ToString, "I", True)
                                    End If

                                    addToJSON(sJSON, "unlimited", False, "B", True)

                                    'Remmato 12/03/21 a seguito di mail di Jacopo Bruno
                                    'addToJSON(sJSON, "price", iPrezzoScontato.ToString, "I", True)
                                    addToJSON(sJSON, "price", iPrezzoOriginale.ToString, "I", True)

                                    'addToJSON(sJSON, "compareToPrice", iPrezzoOriginale.ToString, "I", False)

                                    addToJSON(sJSON, "discount", row.Item("PercSconto").ToString, "I", True)
                                    addToJSON(sJSON, "discountStartDate", row.Item("DataInizioSconto").ToString, "D", True)
                                    'addToJSON(sJSON, "discountEndDate", row.Item("DataFineSconto").ToString, "D", False)

                                    'MATTIA 27/09/22 Aggiunota su ricihiesta di Jacopo Bruno
                                    If Not IsEmpty(row.Item("p_BarCodeTg" & i.ToString).ToString) Then
                                        addToJSON(sJSON, "discountEndDate", row.Item("DataFineSconto").ToString, "D", True)
                                        addToJSON(sJSON, "UPC", row.Item("p_BarCodeTg" & i.ToString).ToString, "S", False)
                                    Else
                                        addToJSON(sJSON, "discountEndDate", row.Item("DataFineSconto").ToString, "D", False)
                                    End If

                                    sJSON &= "},"

                                    If iRowsMade = 1 Then
                                        iVariationFiles += 1
                                        htUploadVariationJSON.Add(iVariationFiles, "[" & sJSON)

                                    ElseIf iRowsMade < iRowLimit Then
                                        htUploadVariationJSON(iVariationFiles) &= sJSON

                                    ElseIf iRowsMade = iRowLimit Then
                                        htUploadVariationJSON(iVariationFiles) = Left(htUploadVariationJSON(iVariationFiles), Len(htUploadVariationJSON(iVariationFiles)) - 1) & "]"

                                        iRowsMade = 0
                                    End If
                                Catch ex As Exception
                                    ThrowException("Errore nella procedura UploadVariation (errore .Net):" & vbNewLine & ex.ToString, True)
                                End Try
                            End If
                        End If
                    Next
                Next
            End If
        Else
            ThrowException("Errore nella procedura UploadVariation (errore hSQL):" & vbNewLine & hSQL.LeggiErrore)
        End If

        hSQL.Dispose()
        dtQtaPerTaglie.Dispose()
    End Sub

    Private Function GetInFiltroLight() As String
        Dim sSQL As String = ""
        sSQL = sSQL & "SELECT CodArt" & vbNewLine
        sSQL = sSQL & "FROM mas_Articoli" & vbNewLine
        sSQL = sSQL & "WHERE DataMod >= DATEADD(day, -1, GETDATE())" & vbNewLine
        sSQL = sSQL & "" & vbNewLine
        sSQL = sSQL & "UNION" & vbNewLine
        sSQL = sSQL & "" & vbNewLine
        sSQL = sSQL & "SELECT DISTINCT mag_DMovimenti.CodArt" & vbNewLine
        sSQL = sSQL & "FROM mag_DMovimenti" & vbNewLine
        sSQL = sSQL & "INNER JOIN mag_TMovimenti" & vbNewLine
        sSQL = sSQL & "ON mag_TMovimenti.AnnoOper = mag_DMovimenti.AnnoOper" & vbNewLine
        sSQL = sSQL & "AND mag_TMovimenti.NumeroOper = mag_DMovimenti.NumeroOper" & vbNewLine
        sSQL = sSQL & "WHERE mag_TMovimenti.DataOper >= DATEADD(day, -1, GETDATE())" & vbNewLine
        sSQL = sSQL & "   OR mag_DMovimenti.DataIns >= DATEADD(day, -1, GETDATE())" & vbNewLine
        sSQL = sSQL & "   OR mag_DMovimenti.DataMod >= DATEADD(day, -1, GETDATE())" & vbNewLine
        sSQL = sSQL & "" & vbNewLine
        sSQL = sSQL & "UNION" & vbNewLine
        sSQL = sSQL & "" & vbNewLine
        sSQL = sSQL & "SELECT DISTINCT CodArt" & vbNewLine
        sSQL = sSQL & "FROM bru_CodArtEliminati" & vbNewLine
        sSQL = sSQL & "WHERE DataOra >= DATEADD(day, -1, GETDATE())" & vbNewLine
        sSQL = sSQL & "" & vbNewLine

        Return sSQL
    End Function
End Module