Imports System
Imports Microsoft.VisualBasic
Imports Amos
Imports AmosEngineLib
Imports AmosEngineLib.AmosEngine.TMatrixID
Imports MiscAmosTypes
Imports MiscAmosTypes.cDatabaseFormat
Imports System.Xml

<System.ComponentModel.Composition.Export(GetType(Amos.IPlugin))>
Public Class CustomCode
    Implements IPlugin

    'This plugin was written January 2018 by John Lim for James Gaskin.
    Public Function Name() As String Implements IPlugin.Name
        Return "Indirect Effects"
    End Function

    Public Function Description() As String Implements IPlugin.Description
        Return "Creates matrices of all possible standardized and unstandardized indirect effects in the model."
    End Function

    Public Function Mainsub() As Integer Implements IPlugin.MainSub

        'Settings to get bootstrap estimates.
        pd.GetCheckBox("AnalysisPropertiesForm", "StandardizedCheck").Checked = True
        pd.GetCheckBox("AnalysisPropertiesForm", "DoBootstrapCheck").Checked = True
        pd.GetTextBox("AnalysisPropertiesForm", "BootstrapText").Text = "2000"
        pd.GetCheckBox("AnalysisPropertiesForm", "ConfidenceBCCheck").Checked = True
        pd.GetTextBox("AnalysisPropertiesForm", "ConfidenceBCText").Text = "90"

        'Create the html output that combines standardized/unstandardized indirect effects.
        CreateOutput()

    End Function

    'Create the html output that combines standardized/unstandardized indirect effects.
    Sub CreateOutput()

        pd.AnalyzeCalculateEstimates()

        'Get regression weights, standardized regression weights, and user-defined estimands w/bootstrap confidence intervals xml tables from the output.
        Dim tableRegression As XmlElement = GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Regression Weights:']/table/tbody")
        Dim tableStandardized As XmlElement = GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Standardized Regression Weights:']/table/tbody")
        Dim tableBootstrap As XmlElement = GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='bootstrap']/div[@ntype='bootstrapconfidence']/div[@ntype='biascorrected']/div[@ntype='scalars']/div[@nodecaption='User-defined estimands:']/table/tbody")
        Dim numRegression As Integer = GetNodeCount(tableRegression)
        Dim numBootstrap As Integer = GetNodeCount(tableBootstrap)
        Dim standardizedIndirectEffects(numBootstrap - 1) As Double 'Array to hold standardized indirect effects

        For x = 1 To numBootstrap 'Iterate through the table of bootstrap estimates.
            Dim paths As String() = Strings.Split(MatrixName(tableBootstrap, x, 0), " --> ") 'Split variable names into array.
            For y = 1 To numRegression 'Iterate through Standardized Regression Weights to match first path.
                If paths(0) = MatrixName(tableStandardized, y, 2) And paths(1) = MatrixName(tableStandardized, y, 0) Then
                    Dim firstPath As Double = MatrixElement(tableStandardized, y, 3)
                    For z = 1 To numRegression 'Iterate through Standardized Regression Weights again to match second path.
                        If paths(1) = MatrixName(tableStandardized, z, 2) And paths(2) = MatrixName(tableStandardized, z, 0) Then
                            Dim secondPath As Double = MatrixElement(tableStandardized, z, 3)
                            standardizedIndirectEffects(x - 1) = firstPath * secondPath 'Multiply standardized estimates together to get indirect effect
                        End If
                    Next
                End If
            Next
        Next

        'Delete the output file if it exists
        If (System.IO.File.Exists("IndirectEffects.html")) Then
            System.IO.File.Delete("IndirectEffects.html")
        End If

        'Start the debugger for the html output
        Dim debug As New AmosDebug.AmosDebug
        Dim resultWriter As New TextWriterTraceListener("IndirectEffects.html")
        Trace.Listeners.Clear()
        Trace.Listeners.Add(resultWriter)

        debug.PrintX("<html><body><h1>Indirect Effects</h1><hr/>")

        'Populate model fit measures in data table
        debug.PrintX("<table><tr><th>Indirect Path</th><th>Unstandardized Estimate</th><th>Lower</th><th>Upper</th><th>P-Value</th><th>Standardized Estimate</th></tr><tr>")

        For i = 1 To numBootstrap
            debug.PrintX("<td>" + MatrixName(tableBootstrap, i, 0) + "</td>") 'Name of indirect path
            debug.PrintX("<td>" + MatrixElement(tableBootstrap, i, 3).ToString("#0.000")) 'Estimate
            debug.PrintX("<td>" + MatrixElement(tableBootstrap, i, 4).ToString("#0.000")) 'Lower
            debug.PrintX("<td>" + MatrixElement(tableBootstrap, i, 5).ToString("#0.000")) 'Upper
            debug.PrintX("<td>" + MatrixElement(tableBootstrap, i, 6).ToString("#0.000")) 'P-Value

            'Output the significance significance with the standardized estimate
            If MatrixName(tableBootstrap, i, 6) = "***" Then
                debug.PrintX("<td>" + standardizedIndirectEffects(i - 1).ToString("#0.000") + "***</td>")
            ElseIf MatrixName(tableBootstrap, i, 6) = "" Then
                debug.PrintX("<td>" + standardizedIndirectEffects(i - 1).ToString("#0.000") + "</td>")
            ElseIf MatrixElement(tableBootstrap, i, 6) = 0 Then
                debug.PrintX("<td>" + standardizedIndirectEffects(i - 1).ToString("#0.000") + "</td>")
            ElseIf MatrixElement(tableBootstrap, i, 6) < 0.001 Then
                debug.PrintX("<td>" + standardizedIndirectEffects(i - 1).ToString("#0.000") + "***</td>")
            ElseIf MatrixElement(tableBootstrap, i, 6) < 0.01 Then
                debug.PrintX("<td>" + standardizedIndirectEffects(i - 1).ToString("#0.000") + "**</td>")
            ElseIf MatrixElement(tableBootstrap, i, 6) < 0.05 Then
                debug.PrintX("<td>" + standardizedIndirectEffects(i - 1).ToString("#0.000") + "*</td>")
            ElseIf MatrixElement(tableBootstrap, i, 6) < 0.1 Then
                debug.PrintX("<td>" + standardizedIndirectEffects(i - 1).ToString("#0.000") + "&#x271D;</td>")
            Else
                debug.PrintX("<td>" + standardizedIndirectEffects(i - 1).ToString("#0.000") + "</td>")
            End If

            debug.PrintX("</tr>")
        Next

        'References
        debug.PrintX("</table><h3>References</h3>Significance of Estimates:<br>*** p < 0.001<br>** p < 0.010<br>* p < 0.050<br>&#x271D; p < 0.100<br>")
        debug.PrintX("<p>--If you would like to cite this tool directly, please use the following:")
        debug.PrintX("Gaskin, J., James, M., & Lim, J. (2020), ""Indirect Effects"", AMOS Plugin. <a href=""http://statwiki.kolobkreations.com"">Gaskination's StatWiki</a>.</p>")

        'Write style And close
        debug.PrintX("<style>table{border:1px solid black;border-collapse:collapse;}td{border:1px solid black;text-align:center;padding:5px;}th{text-weight:bold;padding:10px;border: 1px solid black;}</style>")
        debug.PrintX("</body></html>")

        'Take down our debugging, release file, open html
        Trace.Flush()
        Trace.Listeners.Remove(resultWriter)
        resultWriter.Close()
        resultWriter.Dispose()
        Process.Start("IndirectEffects.html")

    End Sub

    'Get a string element from an xml table.
    Function MatrixName(eTableBody As XmlElement, row As Long, column As Long) As String

        Dim e As XmlElement

        Try
            e = eTableBody.ChildNodes(row - 1).ChildNodes(column) 'This means that the rows are not 0 based.
            MatrixName = e.InnerText
        Catch ex As Exception
            MatrixName = ""
        End Try

    End Function

    'Get a number from an xml table.
    Function MatrixElement(eTableBody As XmlElement, row As Long, column As Long) As Double

        Dim e As XmlElement

        Try
            e = eTableBody.ChildNodes(row - 1).ChildNodes(column) 'This means that the rows are not 0 based.
            MatrixElement = CDbl(e.GetAttribute("x"))
        Catch ex As Exception
            MatrixElement = 0
        End Try

    End Function

    'Use an output table path to get the xml version of the table.
    Function GetXML(path As String) As XmlElement

        Dim doc As Xml.XmlDocument = New Xml.XmlDocument()
        doc.Load(Amos.pd.ProjectName & ".AmosOutput")
        Dim nsmgr As XmlNamespaceManager = New XmlNamespaceManager(doc.NameTable)
        Dim eRoot As Xml.XmlElement = doc.DocumentElement

        GetXML = eRoot.SelectSingleNode(path, nsmgr)

    End Function

    'Get the number of rows in an xml table.
    Function GetNodeCount(table As XmlElement) As Integer

        Dim nodeCount As Integer = 0

        'Handles a model with zero correlations
        Try
            nodeCount = table.ChildNodes.Count
        Catch ex As NullReferenceException
            nodeCount = 0
        End Try

        GetNodeCount = nodeCount

    End Function

    'Set all the path values to null.
    Sub ClearPaths()
        'Set paths back to null.
        For Each variable As PDElement In pd.PDElements 'Iterate through the paths in the model
            If variable.IsPath Then
                If (variable.Variable1.IsLatentVariable And variable.Variable2.IsLatentVariable) Or (variable.Variable1.IsObservedVariable And variable.Variable2.IsObservedVariable) Then
                    variable.Value1 = variable.Variable1.NameOrCaption + " | " + variable.Variable2.NameOrCaption 'Change path to names of the connected variables.
                End If
            End If
        Next
    End Sub

    'Set path values to concatenated names of connected variables.
    Sub NamePaths()
        For Each variable As PDElement In pd.PDElements 'Iterate through the paths in the model
            If variable.IsPath Then
                If (variable.Variable1.IsLatentVariable And variable.Variable2.IsLatentVariable) Or (variable.Variable1.IsObservedVariable And variable.Variable2.IsObservedVariable) Then
                    variable.Value1 = variable.Variable1.NameOrCaption + " | " + variable.Variable2.NameOrCaption 'Change path to names of the connected variables.
                End If
            End If
        Next
    End Sub

End Class


