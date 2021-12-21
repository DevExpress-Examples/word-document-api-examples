Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Text.RegularExpressions

Namespace RichEditDocumentServerAPIExample.CodeUtils

#Region "ExampleFinder"
    Public MustInherit Class ExampleFinder

        Public MustOverride ReadOnly Property RegexRegionPattern As String

        Public MustOverride ReadOnly Property RegionStartPattern As String

        Public MustOverride ReadOnly Property RegionHelperStartPattern As String

        Public Function Process(ByVal fileWithExample As FileInfo) As List(Of CodeExample)
            If fileWithExample Is Nothing Then Return New List(Of CodeExample)()
            Dim groupName As String = Path.GetFileNameWithoutExtension(fileWithExample.Name)
            Dim code As String
            Using stream As FileStream = File.Open(fileWithExample.FullName, FileMode.Open, FileAccess.Read)
                Dim sr As StreamReader = New StreamReader(stream)
                code = sr.ReadToEnd()
            End Using

            Dim foundExamples As List(Of CodeExample) = ParseSouceFileAndFindRegionsWithExamples(groupName, code)
            Return foundExamples
        End Function

        Public Function ParseSouceFileAndFindRegionsWithExamples(ByVal groupName As String, ByVal sourceCode As String) As List(Of CodeExample)
            Dim result As List(Of CodeExample) = New List(Of CodeExample)()
            Dim matches = Regex.Matches(sourceCode, RegexRegionPattern, RegexOptions.Singleline)
            For Each match In matches
                Dim lines As String() = match.ToString().Split(New String() {Microsoft.VisualBasic.Constants.vbCrLf, Microsoft.VisualBasic.Constants.vbLf}, StringSplitOptions.None)
                If lines.Length <= 2 Then Continue For
                lines = DeleteLeadingWhiteSpacesFromSourceCode(lines)
                Dim regionName As String = String.Empty
                Dim regionIsValid As Boolean = ValidateRegionName(lines, regionName)
                If Not regionIsValid Then Continue For
                Dim exampleCode As String = String.Join(Microsoft.VisualBasic.Constants.vbCrLf, lines, 1, lines.Length - 2)
                result.Add(CreateRichEditExample(groupName, regionName, exampleCode))
            Next

            Return result
        End Function

        Protected Function CreateRichEditExample(ByVal exampleGroup As String, ByVal regionName As String, ByVal exampleCode As String) As CodeExample
            Dim result As CodeExample = New CodeExample()
            SetExampleCode(exampleCode, result)
            result.RegionName = regionName
            result.HumanReadableGroupName = ConvertStringToHumanReadableForm(exampleGroup)
            Return result
        End Function

        Protected MustOverride Sub SetExampleCode(ByVal exampleCode As String, ByVal newExample As CodeExample)

        Protected Overridable Function DeleteLeadingWhiteSpacesFromSourceCode(ByVal lines As String()) As String()
            Return DeleteLeadingWhiteSpaces(lines, "        ")
        End Function

        Protected Overridable Function ValidateRegionName(ByVal lines As String(), ByRef regionName As String) As Boolean
            Dim keepHashMark As Integer = 0 ' "#example" if value is -1 or "example" if value will be 0
            Dim region As String = lines(0)
            Dim regionIndex As Integer = region.IndexOf(RegionHelperStartPattern)
            If regionIndex = 0 Then
                regionName = ConvertStringToHumanReadableForm(region.Substring(regionIndex + RegionHelperStartPattern.Length + keepHashMark))
            End If

            If regionIndex < 0 Then
                regionIndex = region.IndexOf(RegionStartPattern)
                If regionIndex < 0 Then
                    regionName = String.Empty
                    Return False
                End If

                regionName = ConvertStringToHumanReadableForm(region.Substring(regionIndex + RegionStartPattern.Length + keepHashMark))
            End If

            Return True
        End Function
    End Class

#End Region
#Region "ExampleFinderVB"
    Public Class ExampleFinderVB
        Inherits ExampleFinder

        Public Overrides ReadOnly Property RegexRegionPattern As String
            Get
                Return "#Region.*?#End Region"
            End Get
        End Property

        Public Overrides ReadOnly Property RegionStartPattern As String
            Get
                Return "#Region ""#"
            End Get
        End Property

        Public Overrides ReadOnly Property RegionHelperStartPattern As String
            Get
                Return "#Region ""#@"
            End Get
        End Property

        Protected Overrides Function DeleteLeadingWhiteSpacesFromSourceCode(ByVal lines As String()) As String()
            Dim result As String() = MyBase.DeleteLeadingWhiteSpacesFromSourceCode(lines)
            Return CodeExampleUtils.DeleteLeadingWhiteSpaces(result, Microsoft.VisualBasic.Constants.vbTab & Microsoft.VisualBasic.Constants.vbTab)
        End Function

        Protected Overrides Function ValidateRegionName(ByVal lines As String(), ByRef regionName As String) As Boolean
            Dim result As Boolean = MyBase.ValidateRegionName(lines, regionName)
            If Not result Then Return result
            regionName = regionName.TrimEnd(""""c)
            Return True
        End Function

        Protected Overrides Sub SetExampleCode(ByVal code As String, ByVal newExample As CodeExample)
            newExample.CodeVB = code
        End Sub
    End Class

#End Region
#Region "ExampleFinderCSharp"
    Public Class ExampleFinderCSharp
        Inherits ExampleFinder

        Public Overrides ReadOnly Property RegexRegionPattern As String
            Get
                Return "#region.*?#endregion"
            End Get
        End Property

        Public Overrides ReadOnly Property RegionStartPattern As String
            Get
                Return "#region #"
            End Get
        End Property

        Public Overrides ReadOnly Property RegionHelperStartPattern As String
            Get
                Return "#region #@"
            End Get
        End Property

        Protected Overrides Sub SetExampleCode(ByVal code As String, ByVal newExample As CodeExample)
            newExample.CodeCS = code
        End Sub
    End Class
#End Region
End Namespace
