Imports System.IO
Imports System.Threading
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Public Class WordSaver

    Dim oWordApp As Object = Nothing
    Dim oWordDoc As Object
    Dim oWordSel As Object
    Dim oWordSec As Object

    Public Sub New()
        If (Me.oWordApp Is Nothing) Then
            Try
                Me.oWordApp = CreateObject("Word.Application")
            Catch ex As Exception
                MessageBox.Show("Erro ao salvar o documento. É necessario instalar o Microsoft Word", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End Try

            Me.oWordDoc = Me.oWordApp.Documents.Add()
            Me.oWordDoc.Activate()
            Me.oWordSel = Me.oWordApp.Selection
            Me.oWordSec = Me.oWordDoc.Sections(1)
            Me.oWordSec.PageSetup.DifferentFirstPageHeaderFooter = True 'C# true = -1 and false = 0
        End If
    End Sub

    Public Function IsCreated() As Boolean
        If (Me.oWordApp Is Nothing) Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Sub InsertParagraph()
        Me.oWordSel.TypeParagraph()
    End Sub

    Public Sub InsertText(ByVal text As String)
        Me.oWordSel.TypeText(text)
    End Sub

    Public Sub Collapse(ByVal value As Integer)
        Me.oWordSel.Collapse(value)
    End Sub

    Public Sub Save(ByVal outputFile As String, Optional ByVal tipo As Integer = 0)
        Me.oWordDoc.SaveAs(outputFile, tipo)
        Me.oWordApp.Quit()
    End Sub

    Public Sub InsertFileHidden(ByVal fileName As String)
        Dim newStart As Integer = Me.oWordDoc.Paragraphs.Count
        Dim count As Integer = Me.oWordDoc.Paragraphs.Count

        Me.oWordSel.SetRange(Me.oWordDoc.Paragraphs(count).Range.End, Me.oWordDoc.Paragraphs(count).Range.End)

        If Me.InsertFile(fileName) Then
            count = Me.oWordDoc.Paragraphs.Count
            Me.oWordSel.SetRange(Me.oWordDoc.Paragraphs(newStart - 2).Range.Start, Me.oWordDoc.Paragraphs(count).Range.End)
            Me.oWordSel.Font.Hidden = True
        End If

        count = Me.oWordDoc.Paragraphs.Count
        Me.oWordSel.SetRange(Me.oWordDoc.Paragraphs(count).Range.End, Me.oWordDoc.Paragraphs(count).Range.End)
        Me.oWordSel.Font.Hidden = False
    End Sub

    Function InsertFile(ByVal fileName As String) As Boolean
        If File.Exists(fileName) Then
            Me.oWordSel.InsertFile(fileName)
            Return True
        End If
        Return False
    End Function
End Class
