Imports System
Imports System.CodeDom.Compiler
Imports System.Reflection

Namespace RichEditDocumentServerAPIExample.CodeUtils
	Public MustInherit Class ExampleCodeEvaluator
		Protected MustOverride ReadOnly Property CodeStart() As String
		Protected MustOverride ReadOnly Property CodeBeforeClasses() As String
		Protected MustOverride ReadOnly Property CodeEnd() As String
		Protected MustOverride Function GetCodeDomProvider() As CodeDomProvider
		Protected MustOverride Function GetModuleAssembly() As String
		Protected MustOverride Function GetExampleClassName() As String

		Public Function ExecuteCodeAndGenerateDocument(ByVal args As CodeEvaluationEventArgs) As Boolean
			Dim theCode As String = String.Concat(CodeStart, args.Code, CodeBeforeClasses, args.CodeClasses, CodeEnd)
			Dim linesOfCode() As String = { theCode }
			Return CompileAndRun(linesOfCode, args.EvaluationParameter)
		End Function

		Protected Friend Function CompileAndRun(ByVal linesOfCode() As String, ByVal evaluationParameter As Object) As Boolean
			Dim CompilerParams As New CompilerParameters()

			CompilerParams.GenerateInMemory = True
			CompilerParams.TreatWarningsAsErrors = False
			CompilerParams.GenerateExecutable = False

			Dim referencesSystem() As String = { "System.dll", "System.Windows.Forms.dll", "System.Data.dll", "System.Xml.dll", "System.Core.dll", "System.Drawing.dll" }

			Dim referencesDX() As String = { AssemblyInfo.SRAssemblyData, GetModuleAssembly(), AssemblyInfo.SRAssemblyOfficeCore, AssemblyInfo.SRAssemblyPrintingCore, AssemblyInfo.SRAssemblyPrinting, AssemblyInfo.SRAssemblyDocs, AssemblyInfo.SRAssemblyUtils, AssemblyInfo.SRAssemblyRichEdit }
			Dim references((referencesSystem.Length + referencesDX.Length) - 1) As String

			For referenceIndex As Integer = 0 To referencesSystem.Length - 1
				references(referenceIndex) = referencesSystem(referenceIndex)
			Next referenceIndex

			Dim i As Integer = 0
			Dim initial As Integer = referencesSystem.Length
			Do While i < referencesDX.Length
				Dim assembly As System.Reflection.Assembly = System.Reflection.Assembly.Load(referencesDX(i) & AssemblyInfo.FullAssemblyVersionExtension)
				If assembly IsNot Nothing Then
					references(i + initial) = assembly.Location
				End If
				i += 1
			Loop
			CompilerParams.ReferencedAssemblies.AddRange(references)


			Dim provider As CodeDomProvider = GetCodeDomProvider()
			Dim compile As CompilerResults = provider.CompileAssemblyFromSource(CompilerParams, linesOfCode)

			If compile.Errors.HasErrors Then
				Dim text As String = "Compile error: "
				For Each ce As CompilerError In compile.Errors
					text &= "rn" & ce.ToString()
				Next ce
				System.Windows.Forms.MessageBox.Show(text)
				Return False
			End If

			Dim [module] As System.Reflection.Module = Nothing
			Try
				[module] = compile.CompiledAssembly.GetModules()(0)
			Catch
			End Try
			Dim moduleType As Type = Nothing
			If [module] Is Nothing Then
				Return False
			End If
			moduleType = [module].GetType(GetExampleClassName())

			Dim methInfo As MethodInfo = Nothing
			If moduleType Is Nothing Then
				Return False
			End If
			methInfo = moduleType.GetMethod("Process")

			If methInfo IsNot Nothing Then
				Try
					methInfo.Invoke(Nothing, New Object() { evaluationParameter })
				Catch e1 As Exception
					Return False ' an error
				End Try
				Return True
			End If
			Return False
		End Function
	End Class

	Public MustInherit Class RichEditExampleCodeEvaluator
		Inherits ExampleCodeEvaluator


		Protected Overrides Function GetModuleAssembly() As String
			Return AssemblyInfo.SRAssemblyRichEditCore
		End Function
		Protected Overrides Function GetExampleClassName() As String
			Return "RichEditCodeResultViewer.ExampleItem"
		End Function
	End Class
	#Region "RichEditCSExampleCodeEvaluator"
	Public Class RichEditCSExampleCodeEvaluator
		Inherits RichEditExampleCodeEvaluator

		Protected Overrides Function GetCodeDomProvider() As CodeDomProvider
			Return New Microsoft.CSharp.CSharpCodeProvider()
		End Function
'INSTANT VB NOTE: The field codeStart was renamed since Visual Basic does not allow fields to have the same name as other class members:
		Private Const codeStart_Conflict As String = "using System;" & vbCrLf & "using System.Text;" & vbCrLf & "using DevExpress.XtraPrinting;" & vbCrLf & "using DevExpress.XtraRichEdit;" & vbCrLf & "using DevExpress.XtraRichEdit.API.Native;" & vbCrLf & "using DevExpress.XtraRichEdit.Export;" & vbCrLf & "using DevExpress.XtraRichEdit.Import;" & vbCrLf & "using DevExpress.Office.Utils;" & vbCrLf & "using System.Drawing;" & vbCrLf & "using System.Windows.Forms;" & vbCrLf & "using DevExpress.Utils;" & vbCrLf & "using System.IO;" & vbCrLf & "using System.Diagnostics;" & vbCrLf & "using System.Xml;" & vbCrLf & "using System.Data;" & vbCrLf & "using System.Collections.Generic;" & vbCrLf & "using System.Linq;" & vbCrLf & "using System.Globalization;" & vbCrLf & "using Document = DevExpress.XtraRichEdit.API.Native.Document;" & vbCrLf & "using TableRow = DevExpress.XtraRichEdit.API.Native.TableRow;" & vbCrLf & "namespace RichEditCodeResultViewer { " & vbCrLf & "public class ExampleItem { " & vbCrLf & "        public static void Process(RichEditDocumentServer wordProcessor) { " & vbCrLf & vbCrLf

'INSTANT VB NOTE: The field codeBeforeClasses was renamed since Visual Basic does not allow fields to have the same name as other class members:
		Private Const codeBeforeClasses_Conflict As String = "       " & vbCrLf & " }" & vbCrLf & "    }" & vbCrLf

'INSTANT VB NOTE: The field codeEnd was renamed since Visual Basic does not allow fields to have the same name as other class members:
		Private Const codeEnd_Conflict As String = vbCrLf & "    }" & vbCrLf

		Protected Overrides ReadOnly Property CodeStart() As String
			Get
				Return codeStart_Conflict
			End Get
		End Property
		Protected Overrides ReadOnly Property CodeBeforeClasses() As String
			Get
				Return codeBeforeClasses_Conflict
			End Get
		End Property
		Protected Overrides ReadOnly Property CodeEnd() As String
			Get
				Return codeEnd_Conflict
			End Get
		End Property
	End Class
	#End Region
	#Region "RichEditVbExampleCodeEvaluator"
	Public Class RichEditVbExampleCodeEvaluator
		Inherits RichEditExampleCodeEvaluator

		Protected Overrides Function GetCodeDomProvider() As CodeDomProvider
			Return New VBCodeProvider()
		End Function
'INSTANT VB NOTE: The field codeStart was renamed since Visual Basic does not allow fields to have the same name as other class members:
		Private Const codeStart_Conflict As String = "Imports Microsoft.VisualBasic" & vbCrLf & "Imports System" & vbCrLf & "Imports DevExpress.XtraRichEdit" & vbCrLf & "Imports DevExpress.XtraRichEdit.API.Native" & vbCrLf & "Imports System.Drawing" & vbCrLf & "Imports System.Windows.Forms" & vbCrLf & "Imports DevExpress.Utils" & vbCrLf & "Imports System.IO" & vbCrLf & "Imports System.Diagnostics" & vbCrLf & "Imports System.Xml" & vbCrLf & "Imports System.Data" & vbCrLf & "Imports System.Collections.Generic" & vbCrLf & "Imports System.Globalization" & vbCrLf & "Imports Document = DevExpress.XtraRichEdit.API.Native.Document" & vbCrLf & "Imports TableRow = DevExpress.XtraRichEdit.API.Native.TableRow" & vbCrLf & "Namespace RichEditCodeResultViewer" & vbCrLf & "	Public Class ExampleItem" & vbCrLf & "		Public Shared Sub Process(ByVal wordProcessor As RichEditDocumentServer)" & vbCrLf & vbCrLf

'INSTANT VB NOTE: The field codeBeforeClasses was renamed since Visual Basic does not allow fields to have the same name as other class members:
		Private Const codeBeforeClasses_Conflict As String = vbCrLf & "		End Sub" & vbCrLf & "	End Class" & vbCrLf

'INSTANT VB NOTE: The field codeEnd was renamed since Visual Basic does not allow fields to have the same name as other class members:
		Private Const codeEnd_Conflict As String = vbCrLf & "End Namespace" & vbCrLf

		Protected Overrides ReadOnly Property CodeStart() As String
			Get
				Return codeStart_Conflict
			End Get
		End Property
		Protected Overrides ReadOnly Property CodeBeforeClasses() As String
			Get
				Return codeBeforeClasses_Conflict
			End Get
		End Property
		Protected Overrides ReadOnly Property CodeEnd() As String
			Get
				Return codeEnd_Conflict
			End Get
		End Property
	End Class
	#End Region

	Public MustInherit Class ExampleEvaluatorByTimer
		Implements IDisposable

		Private leakSafeCompileEventRouter As LeakSafeCompileEventRouter
		Private compileExampleTimer As System.Windows.Forms.Timer
		Private compileComplete As Boolean = True
		Private Const CompileTimeIntervalInMilliseconds As Integer = 2000

		Public Sub New(ByVal enableTimer As Boolean)
			Me.leakSafeCompileEventRouter = New LeakSafeCompileEventRouter(Me)

			If enableTimer Then
				Me.compileExampleTimer = New System.Windows.Forms.Timer()
				Me.compileExampleTimer.Interval = CompileTimeIntervalInMilliseconds

				AddHandler compileExampleTimer.Tick, AddressOf leakSafeCompileEventRouter.OnCompileExampleTimerTick 'OnCompileTimerTick
				Me.compileExampleTimer.Enabled = True
			End If
		End Sub
		Public Sub New()
			Me.New(True)
		End Sub

		#Region "Events"
		Public Event QueryEvaluate As CodeEvaluationEventHandler

		Protected Friend Overridable Function RaiseQueryEvaluate() As CodeEvaluationEventArgs
			If QueryEvaluateEvent IsNot Nothing Then
				Dim args As New CodeEvaluationEventArgs()
				RaiseEvent QueryEvaluate(Me, args)
				Return args
			End If
			Return Nothing
		End Function
		Public Event OnBeforeCompile As EventHandler

		Private Sub RaiseOnBeforeCompile()
			RaiseEvent OnBeforeCompile(Me, New EventArgs())
		End Sub

		Public Event OnAfterCompile As OnAfterCompileEventHandler

		Private Sub RaiseOnAfterCompile(ByVal result As Boolean)
			RaiseEvent OnAfterCompile(Me, New OnAfterCompileEventArgs() With {.Result = result})
		End Sub
		#End Region

		Public Sub CompileExample(ByVal sender As Object, ByVal e As EventArgs)
			If Not compileComplete Then
				Return
			End If
			Dim args As CodeEvaluationEventArgs = RaiseQueryEvaluate()
			If Not args.Result Then
				Return
			End If

			ForceCompile(args)
		End Sub
		Public Sub ForceCompile(ByVal args As CodeEvaluationEventArgs)
			compileComplete = False
			If Not String.IsNullOrEmpty(args.Code) Then
				CompileExampleAndShowPrintPreview(args)
			End If

			compileComplete = True
		End Sub
		Private Sub CompileExampleAndShowPrintPreview(ByVal args As CodeEvaluationEventArgs)
			Dim evaluationSucceed As Boolean = False
			Try
				RaiseOnBeforeCompile()

				evaluationSucceed = Evaluate(args)
			Finally
				RaiseOnAfterCompile(evaluationSucceed)
			End Try
		End Sub

		Public Function Evaluate(ByVal args As CodeEvaluationEventArgs) As Boolean
			Dim richeditExampleCodeEvaluator As ExampleCodeEvaluator = GetExampleCodeEvaluator(args.Language)
			Return richeditExampleCodeEvaluator.ExecuteCodeAndGenerateDocument(args)
		End Function

		Protected MustOverride Function GetExampleCodeEvaluator(ByVal language As ExampleLanguage) As ExampleCodeEvaluator

		Public Sub Dispose() Implements IDisposable.Dispose
			If compileExampleTimer IsNot Nothing Then
				compileExampleTimer.Enabled = False
				If leakSafeCompileEventRouter IsNot Nothing Then
					RemoveHandler compileExampleTimer.Tick, AddressOf leakSafeCompileEventRouter.OnCompileExampleTimerTick 'OnCompileTimerTick
				End If
				compileExampleTimer.Dispose()
				compileExampleTimer = Nothing
			End If
		End Sub
	End Class

	#Region "RichEditExampleEvaluatorByTimer"
	Public Class RichEditExampleEvaluatorByTimer
		Inherits ExampleEvaluatorByTimer

		Public Sub New()
			MyBase.New()
		End Sub

		Protected Overrides Function GetExampleCodeEvaluator(ByVal language As ExampleLanguage) As ExampleCodeEvaluator
			If language = ExampleLanguage.VB Then
				Return New RichEditVbExampleCodeEvaluator()
			End If
			Return New RichEditCSExampleCodeEvaluator()
		End Function
	End Class
	#End Region

	#Region "LeakSafeCompileEventRouter"
	Public Class LeakSafeCompileEventRouter
		Private ReadOnly weakControlRef As WeakReference

		Public Sub New(ByVal [module] As ExampleEvaluatorByTimer)
			'Guard.ArgumentNotNull(module, "module");
			Me.weakControlRef = New WeakReference([module])
		End Sub
		Public Sub OnCompileExampleTimerTick(ByVal sender As Object, ByVal e As EventArgs)
			Dim [module] As ExampleEvaluatorByTimer = DirectCast(weakControlRef.Target, ExampleEvaluatorByTimer)
			If [module] IsNot Nothing Then
				[module].CompileExample(sender, e)
			End If
		End Sub
	End Class
	Public Class CodeEvaluationEventArgs
		Inherits EventArgs

		Public Property Result() As Boolean
		Public Property Code() As String
		Public Property CodeClasses() As String
		Public Property Language() As ExampleLanguage
		Public Property EvaluationParameter() As Object
	End Class
	Public Delegate Sub CodeEvaluationEventHandler(ByVal sender As Object, ByVal e As CodeEvaluationEventArgs)

	Public Class OnAfterCompileEventArgs
		Inherits EventArgs

		Public Property Result() As Boolean
	End Class
	Public Delegate Sub OnAfterCompileEventHandler(ByVal sender As Object, ByVal e As OnAfterCompileEventArgs)
	#End Region

End Namespace
