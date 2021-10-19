Imports System
Imports pfcls

Module Program

    Public Class pfcGeometryExamples

        'Analyzing Interference
        '======================================================================
        'Function   :   showInterferences
        'Purpose    :   This function finds the interference in an assembly, 
        '               highlights the interfering surfaces, and calculates 
        '               the interference volume.    
        '======================================================================
        Public Sub showInterferences(ByRef session As IpfcBaseSession)

            Dim model As IpfcModel
            Dim assembly As IpfcAssembly
            Dim globalEval As IpfcGlobalEvaluator
            Dim globalInterferences As IpfcGlobalInterferences
            Dim globalInterference As IpfcGlobalInterference
            Dim selectionPair As IpfcSelectionPair
            Dim selection1, selection2 As IpfcSelection
            Dim interVolume As IpfcInterferenceVolume
            Dim totalVolume As Double
            Dim noInterferences As Integer
            Dim i As Integer

            Try
                '======================================================================
                'Get the current solid
                '======================================================================
                model = session.CurrentModel
                If model Is Nothing Then
                    Throw New Exception("Model not present")
                End If
                If (Not model.Type = EpfcModelType.EpfcMDL_ASSEMBLY) Then
                    Throw New Exception("Model is not an assembly")
                End If
                assembly = CType(model, IpfcAssembly)

                globalEval = (New CMpfcInterference).CreateGlobalEvaluator(assembly)

                '======================================================================
                'Select the list of interferences in the assembly
                'Setting parameter to true will select only solid geometry
                'Setting it to false will through an exception
                '======================================================================
                globalInterferences = globalEval.ComputeGlobalInterference(True)

                If globalInterferences Is Nothing Then
                    Throw New Exception("No interference detected in assembly : " + assembly.FullName)
                    Exit Sub
                End If

                '======================================================================
                'For each interference display interfering surfaces and calculate the
                'interfering volume
                '======================================================================
                noInterferences = globalInterferences.Count
                For i = 0 To noInterferences - 1
                    globalInterference = globalInterferences.Item(i)

                    selectionPair = globalInterference.SelParts
                    selection1 = selectionPair.Sel1
                    selection2 = selectionPair.Sel2
                    selection1.Highlight(EpfcStdColor.EpfcCOLOR_HIGHLIGHT)
                    selection2.Highlight(EpfcStdColor.EpfcCOLOR_HIGHLIGHT)

                    interVolume = globalInterference.Volume
                    totalVolume = interVolume.ComputeVolume()

                    Console.WriteLine("Interference " + i.ToString + " Volume : " + totalVolume.ToString)
                    interVolume.Highlight(EpfcStdColor.EpfcCOLOR_ERROR)

                Next

            Catch ex As Exception
                Console.WriteLine(ex.Message.ToString + Chr(13) + ex.StackTrace.ToString)
                Exit Sub
            End Try
        End Sub

    End Class

    Public Class modelNames
        Public Sub showModelNames(ByRef session As IpfcBaseSession)
            Dim model As IpfcModel
            Dim FullModelName As String

            model = session.CurrentModel
            FullModelName = model.FullName
            Console.WriteLine(FullModelName.ToString)
        End Sub

    End Class

    Sub Main()

        'Connectin to existing Creo session
        Dim connection As IpfcAsyncConnection
        Dim classAsyncConnection As New CCpfcAsyncConnection
        connection = classAsyncConnection.Connect(DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value)

        Dim session As IpfcBaseSession
        session = connection.Session

        'asking function pfcGeometryExamples to show all interferences in model
        Dim test As New pfcGeometryExamples
        test.showInterferences(session)

        'write all model names to command window
        Dim modelNameme As New modelNames
        modelNameme.showModelNames(session)



    End Sub

End Module
'this comment wa not been here