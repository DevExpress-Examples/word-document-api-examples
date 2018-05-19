﻿Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.IO
Imports System.Text
Imports System.Threading.Tasks

Namespace RichEditDocumentServerAPIExample.CodeUtils
    #Region "CodeExampleDemoUtils"
    Public NotInheritable Class CodeExampleDemoUtils

        Private Sub New()
        End Sub

        Public Shared Function GatherExamplesFromProject(ByVal examplesPath As String, ByVal language As ExampleLanguage) As Dictionary(Of String, FileInfo)
            Dim result As New Dictionary(Of String, FileInfo)()
            For Each fileName As String In Directory.GetFiles(examplesPath, "*" & GetCodeExampleFileExtension(language))
                result.Add(Path.GetFileNameWithoutExtension(fileName), New FileInfo(fileName))
            Next fileName
            Return result
        End Function
        Public Shared Function GetCodeExampleFileExtension(ByVal language As ExampleLanguage) As String
            If language = ExampleLanguage.VB Then
                Return ".vb"
            End If
            Return ".cs"
        End Function
        Public Shared Function DeleteLeadingWhiteSpaces(ByVal lines() As String, ByVal stringToDelete As String) As String()
            Dim result(lines.Length - 1) As String
            Dim stringToDeleteLength As Integer = stringToDelete.Length

            For i As Integer = 0 To lines.Length - 1
                Dim index As Integer = lines(i).IndexOf(stringToDelete)
                result(i) = If(index >= 0, lines(i).Substring(index + stringToDeleteLength), lines(i))
            Next i
            Return result
        End Function
        Public Shared Function ConvertStringToMoreHumanReadableForm(ByVal exampleName As String) As String
            Dim result As String = SplitCamelCase(exampleName)
            result = result.Replace(" In ", " in ")
            result = result.Replace(" And ", " and ")
            result = result.Replace(" To ", " to ")
            result = result.Replace(" From ", " from ")
            result = result.Replace(" With ", " with ")
            result = result.Replace(" By ", " by ")
            Return result
        End Function
        Private Shared Function SplitCamelCase(ByVal exampleName As String) As String
            Dim length As Integer = exampleName.Length
            If length = 1 Then
                Return exampleName
            End If

            Dim result As New StringBuilder(length * 2)
            Dim position As Integer = 0
            Do While position < length - 1
                Dim current As Char = exampleName.Chars(position)
                Dim [next] As Char = exampleName.Chars(position + 1)
                result.Append(current)
                If Char.IsLower(current) AndAlso Char.IsUpper([next]) Then
                    result.Append(" "c)
                End If
                position += 1
            Loop
            result.Append(exampleName.Chars(length - 1))
            Return result.ToString()
        End Function
        Public Shared Function GetExamplePath(ByVal exampleFolderName As String) As String '"CodeExamples"
            Dim examplesPath2 As String = Path.Combine(Directory.GetCurrentDirectory() & "\..\..\", exampleFolderName)
            If Directory.Exists(examplesPath2) Then
                Return examplesPath2
            End If
            Dim examplesPathInInsallation As String = GetRelativeDirectoryPath(exampleFolderName)
            Return examplesPathInInsallation
        End Function

        Public Shared Function GetRelativeDirectoryPath(ByVal name As String) As String
            name = "Data\" & name
            Dim path As String = System.Windows.Forms.Application.StartupPath
            Dim s As String = "\"
            For i As Integer = 0 To 10
                If System.IO.Directory.Exists(path & s & name) Then
                    Return (path & s & name)
                Else
                    s &= "..\"
                End If
            Next i
            Return ""
        End Function
        Public Shared Function FindExamples(ByVal examplePath As String, ByVal examplesCS As Dictionary(Of String, FileInfo), ByVal examplesVB As Dictionary(Of String, FileInfo)) As List(Of CodeExampleGroup)

            Dim result As New List(Of CodeExampleGroup)()

            Dim current As Dictionary(Of String, FileInfo) = Nothing
            Dim csExampleFinder As ExampleFinder = New ExampleFinderCSharp()
            Dim vbExampleFinder As ExampleFinder = New ExampleFinderVB()

            current = If(examplesCS.Count <> 0, examplesCS, examplesVB)

            For Each sourceCodeItem As KeyValuePair(Of String, FileInfo) In current
                Dim key As String = sourceCodeItem.Key

                Dim foundExamplesCS As New List(Of CodeExample)()
                If examplesCS.Count <> 0 Then
                    foundExamplesCS = csExampleFinder.Process(examplesCS(key))
                End If

                Dim foundExamplesVB As New List(Of CodeExample)()
                If examplesVB.Count <> 0 Then
                    foundExamplesVB = vbExampleFinder.Process(examplesVB(key))
                End If

                Dim mergedExamplesCollection As New CodeExampleCollection()

                mergedExamplesCollection.Merge(foundExamplesCS)
                mergedExamplesCollection.Merge(foundExamplesVB)

                If mergedExamplesCollection.Count = 0 Then
                    Continue For
                End If

                Dim group As New CodeExampleGroup() With { _
                    .Name = mergedExamplesCollection(0).HumanReadableGroupName, _
                    .Examples = mergedExamplesCollection _
                }
                result.Add(group)
            Next sourceCodeItem
            Return result
        End Function

        Public Shared Function DetectExampleLanguage(ByVal solutionFileNameWithoutExtenstion As String) As ExampleLanguage
            Dim projectPath As String = Directory.GetCurrentDirectory() & "\..\..\"

            Dim csproject() As String = Directory.GetFiles(projectPath, "*.csproj")
            If csproject.Length <> 0 AndAlso csproject(0).EndsWith(solutionFileNameWithoutExtenstion & ".csproj") Then
                Return ExampleLanguage.Csharp
            End If
            Dim vbproject() As String = Directory.GetFiles(projectPath, "*.vbproj")
            If vbproject.Length <> 0 AndAlso vbproject(0).EndsWith(solutionFileNameWithoutExtenstion & ".vbproj") Then
                Return ExampleLanguage.VB
            End If
            Return ExampleLanguage.Csharp
        End Function
    End Class
    #End Region

End Namespace
