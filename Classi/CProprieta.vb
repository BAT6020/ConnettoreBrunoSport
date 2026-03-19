Public Class CProprieta
    Dim sProprieta As String

    Public Sub New()
        sProprieta = ""
    End Sub

    Public Sub Elimina(ByVal sNomeProprieta As String)
        Dim sStartString As String
        Dim sStopString As String
        Dim nPosStart As Long
        Dim nPosStop As Long
        Dim sNewProprieta As String


        sNomeProprieta = UCase(sNomeProprieta)

        sStartString = "<" & sNomeProprieta & ">"
        sStopString = "</" & sNomeProprieta & ">"

        'Gianluca 31/03/2015, aggiunto il .ToUpper su sProprieta, perchè la funzione _
        'InStr è case sensitive.

        nPosStart = InStr(sProprieta.ToUpper, sStartString)
        If nPosStart > 0 Then

            nPosStop = InStr(sProprieta.ToUpper, sStopString)
            nPosStop = nPosStop + Len(sStopString)

            sNewProprieta = ""
            sNewProprieta = sNewProprieta + Mid(sProprieta, 1, CInt(nPosStart - 1))
            sNewProprieta = sNewProprieta + Mid(sProprieta, CInt(nPosStop))
            sProprieta = sNewProprieta
        End If


    End Sub

    Public Function Leggi(ByVal sNomeProprieta As String) As String
        Dim sValore As Object
        Dim sStartString As String
        Dim sStopString As String
        Dim nPosStart As Long
        Dim nPosStop As Long

        If IsEmpty(sNomeProprieta) Then
            sValore = sProprieta
        Else
            'If TypeName(sNomeProprieta) <> "String" And TypeName(sNomeProprieta) <> "Field" Then
            '    sNomeProprieta = NomeProprieta(sNomeProprieta)
            'End If

            sNomeProprieta = UCase(sNomeProprieta)

            sStartString = "<" & sNomeProprieta & ">"
            sStopString = "</" & sNomeProprieta & ">"

            'Gianluca 31/03/2015, aggiunto il .ToUpper su sProprieta, perchè la funzione _
            'InStr è case sensitive.
            'sProprieta = sProprieta.Replace(sStartString.ToLower, sStartString)
            'sProprieta = sProprieta.Replace(sStopString.ToLower, sStopString)

            nPosStart = InStr(sProprieta.ToUpper, sStartString)
            If nPosStart > 0 Then
                nPosStop = InStr(sProprieta.ToUpper, sStopString)

                nPosStart = nPosStart + Len(sStartString)

                sValore = Mid(sProprieta, CInt(nPosStart), CInt(nPosStop - nPosStart))

            Else
                sValore = Nothing
            End If
        End If

        Leggi = CStr(sValore)
    End Function

    Public Sub Scrivi(ByVal sNomeProprieta As String, ByVal sValore As String)
        Dim nPosStart As Long
        Dim nPosStop As Long

        Dim sStartString As String
        Dim sStopString As String
        Dim sStringa As String

        If Len(sNomeProprieta) > 0 Then
            sNomeProprieta = UCase(sNomeProprieta)

            sStartString = "<" & sNomeProprieta & ">"
            sStopString = "</" & sNomeProprieta & ">"

            'Gianluca 31/03/2015, aggiunto il .ToUpper su sProprieta, perchè la funzione _
            'InStr è case sensitive.

            nPosStart = InStr(sProprieta.ToUpper, sStartString)
            If nPosStart > 0 Then

                nPosStop = InStr(sProprieta.ToUpper, sStopString)

                'Dario 19/06/2002 modificato perchè ogni volta che scriveva una proprietà la spostava alla fine della stringa
                'sStringa = Left(sProprieta, nPosStart - 1)
                'sStringa = sStringa & Mid(sProprieta, nPosStop + Len(sStopString))

                sStringa = Left(sProprieta, CInt(nPosStart + Len(sStartString) - 1))
                sStringa = sStringa & Mid(sProprieta, CInt(nPosStop))

                'Gianluca 10/07/2019
                'Aggiunto il CompareMethod.Text sulla Replace per renderla non case sensitive.
                sStringa = Replace(sStringa, sStartString & sStopString, sStartString & CStr(sValore) & sStopString, , , CompareMethod.Text)

                sProprieta = sStringa
            Else
                sProprieta = sProprieta & "<" & sNomeProprieta & ">"
                sProprieta = sProprieta & CStr(sValore)
                sProprieta = sProprieta & "</" & sNomeProprieta & ">"
            End If
        Else
            sProprieta = CStr(sValore)
        End If
    End Sub

    Public Function Count() As Integer
        Dim sStopString As String
        Dim nPosStart As Integer

        sStopString = "</"

        Count = 0
        nPosStart = 1
        While nPosStart <> 0
            nPosStart = InStr(nPosStart, sProprieta, sStopString)
            If nPosStart <> 0 Then
                Count = Count + 1
                nPosStart = nPosStart + 1
            End If
        End While
    End Function

    Public Function Contains(ByVal NomeProprieta As String) As Boolean
        If sProprieta.ToString.ToUpper.Contains("<" & NomeProprieta.ToUpper & ">") And sProprieta.Contains("</" & NomeProprieta.ToUpper & ">") Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function Length() As Integer
        Return sProprieta.Length
    End Function

    Public Sub Clear()
        sProprieta = ""
    End Sub

    Public Function NomeProprieta(ByVal PosizioneCampo As Integer) As String
        Dim nPosStart As Integer
        Dim nPosStop As Integer
        Dim xi As Integer

        nPosStart = 0
        For xi = 1 To PosizioneCampo
            nPosStart = InStr(nPosStart + 1, sProprieta, "</")
        Next xi
        If nPosStart <> 0 Then
            nPosStart = nPosStart + 2
            nPosStop = InStr(nPosStart, sProprieta, ">")
            NomeProprieta = Mid(sProprieta, nPosStart, nPosStop - nPosStart)
        Else
            NomeProprieta = ""
        End If

    End Function

    Public Sub setCProprietaFromString(ByVal sStringa As String)
        sProprieta = sStringa
    End Sub

    Public Sub GetHashTableFromCProprieta(ByVal Stringone As CProprieta, ByRef ht As Hashtable)

        Dim myDelims As String() = New String() {"<"}
        Dim Arr() As String = Stringone.Leggi("").ToString.Split(myDelims, StringSplitOptions.None)
        For i As Integer = 0 To Arr.Length - 1
            Dim SubS As String = Arr(i)
            If SubS <> "" Then
                If Stringone.Contains(SubS.Replace("/", "").Replace(">", "").Replace("<", "")) Then
                    Dim key As String = ""
                    Dim value As String = ""
                    key = SubS.Replace("/", "").Replace(">", "").Replace("<", "")
                    value = Stringone.Leggi(key)
                    If Not ht.ContainsKey(key) Then
                        ht.Add(key, Stringone.Leggi(key))
                    End If
                End If
            End If
        Next

    End Sub

End Class