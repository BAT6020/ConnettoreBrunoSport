Imports System.IO

Public Class LogJSON
    Implements IDisposable

    Private sFullFilePath As String
    Private sw As StreamWriter

    Public Shared Function getFileName(ByVal name As Integer) As String
        Select Case name
            Case NamesJSON.ParentCategories
                Return GetValueFromConfig("ParentCategoryFileName")

            Case NamesJSON.ChildCategories
                Return GetValueFromConfig("ChildCategoryFileName")

            Case NamesJSON.UploadBrands
                Return GetValueFromConfig("UploadBrandFileName")

            Case NamesJSON.UploadItems
                Return GetValueFromConfig("UploadItemFileName")

            Case NamesJSON.UploadVariations
                Return GetValueFromConfig("UploadVariationFileName")

            Case Else
                Return ""

        End Select
    End Function

    Public Shared Function getFullPathFromFileName(ByVal name As String) As String
        Return TerminatePath(GetValueFromConfig("PathJsonFiles")) & name
    End Function

    Public Sub New(ByVal name As Integer, ByVal FileNumber As Integer)
        Dim sPath As String = TerminatePath(GetValueFromConfig("PathJsonFiles"))

        If Not Directory.Exists(sPath) Then Directory.CreateDirectory(sPath)

        sPath &= name.ToString & "_" & getFileName(name)

        sFullFilePath = sPath & "_" & FileNumber.ToString & ".json"

        If File.Exists(sFullFilePath) Then File.Delete(sFullFilePath)

        sw = File.CreateText(sFullFilePath)
    End Sub

    Public Sub AddRow(ByVal sRow As String)
        sw.Write(sRow)
        sw.Flush()
    End Sub

#Region "Dispose"
    Private disposedValue As Boolean

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: eliminare lo stato gestito (oggetti gestiti)

                sw.Close()
            End If

            ' TODO: liberare risorse non gestite (oggetti non gestiti) ed eseguire l'override del finalizzatore
            ' TODO: impostare campi di grandi dimensioni su Null
            disposedValue = True
        End If
    End Sub

    ' ' TODO: eseguire l'override del finalizzatore solo se 'Dispose(disposing As Boolean)' contiene codice per liberare risorse non gestite
    ' Protected Overrides Sub Finalize()
    '     ' Non modificare questo codice. Inserire il codice di pulizia nel metodo 'Dispose(disposing As Boolean)'
    '     Dispose(disposing:=False)
    '     MyBase.Finalize()
    ' End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        ' Non modificare questo codice. Inserire il codice di pulizia nel metodo 'Dispose(disposing As Boolean)'
        Dispose(disposing:=True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region
End Class
