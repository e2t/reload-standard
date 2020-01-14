Attribute VB_Name = "Main"
'Written in 2015-2016 by Eduard E. Tikhenko <eduard@tikhenko.com>
'
'To the extent possible under law, the author(s) have dedicated all copyright
'and related and neighboring rights to this software to the public domain
'worldwide. This software is distributed without any warranty.
'You should have received a copy of the CC0 Public Domain Dedication along
'with this software.
'If not, see <http://creativecommons.org/publicdomain/zero/1.0/>

Dim swApp As Object
Dim modelExt As ModelDocExtension

Sub Main()
    Dim stdFile As String
    Dim doc As ModelDoc2
    Dim docIsDrawing As Boolean
    
    Set swApp = Application.SldWorks
    If swApp.GetDocumentCount > 0 Then
        Set doc = swApp.ActiveDoc
        docIsDrawing = (doc.GetType = swDocDRAWING)
        Set modelExt = doc.Extension
        stdFile = ReadLinesFrom("settings.ini")(0) & "\" & _
                  IIf(docIsDrawing, "Чертежный", "Модельный") & _
                  " стандарт.sldstd"
        ReloadStandard stdFile
        If docIsDrawing Then
            ChangeLineStyles
        End If
    End If
End Sub

Function GetListFiles(dirname As String, Optional template As String = "*") As String()
    Dim fso As FileSystemObject
    Dim aFolder As folder
    Dim f As Variant
    Dim result() As String
    Dim i As Integer
    
    Set fso = New FileSystemObject
    Set aFolder = fso.GetFolder(dirname)
    If aFolder.Files.count > 0 Then
        ReDim result(aFolder.Files.count - 1)
        i = -1
        For Each f In aFolder.Files
            If f.Name Like template Then
                i = i + 1
                result(i) = f.Name
            End If
        Next
        If i >= 0 Then
            ReDim Preserve result(i)
            GetListFiles = result
        End If
    End If
End Function

Private Sub ReloadStandard(stdFile As String)
    If IsFileExists(stdFile) Then
        modelExt.LoadDraftingStandard (stdFile)
    Else
        MsgBox ("Not Found" + stdFile)
    End If
End Sub

Function IsFileExists(fullname As String) As Boolean
    IsFileExists = False
    If Dir(fullname) <> "" Then
        IsFileExists = True
    End If
End Function

Function ReadLinesFrom(filename As String) As String()
    Dim aStrings() As String
    Dim fullFilename As String
    Dim skip As String
    
    fullFilename = swApp.GetCurrentMacroPathFolder + "\" + filename
    If IsFileExists(fullFilename) Then
        Open fullFilename For Binary As #1
        skip = InputB(2, #1) ' always FF FE (skip first 2 bytes)
        aStrings = Split(InputB(LOF(1), #1), vbNewLine)
        Close #1
    Else
        MsgBox fullFilename
        ReDim aStrings(0)
        aStrings(0) = ""
    End If
    ReadLinesFrom = aStrings
End Function

Sub SetLineStyle(object_type As swUserPreferenceIntegerValue_e, value As Integer)
    'only for drawing extension!
    modelExt.SetUserPreferenceInteger object_type, swDetailingNoOptionSpecified, value
End Sub

Function ChangeLineStyles() 'mask for button
    SetLineStyle swLineFontVisibleEdgesStyle, swLineCONTINUOUS
    SetLineStyle swLineFontVisibleEdgesThickness, swLW_NORMAL
    
    SetLineStyle swLineFontHiddenEdgesStyle, swLineHIDDEN
    SetLineStyle swLineFontHiddenEdgesThickness, swLW_THIN
    
    SetLineStyle swLineFontSketchCurvesStyle, swLineCONTINUOUS
    SetLineStyle swLineFontSketchCurvesThickness, swLW_THIN
    
    SetLineStyle swLineFontConstructionCurvesStyle, swLinePHANTOM
    SetLineStyle swLineFontConstructionCurvesThickness, swLW_THIN
    
    SetLineStyle swLineFontCrosshatchStyle, swLineCONTINUOUS
    SetLineStyle swLineFontCrosshatchThickness, swLW_THIN
    
    SetLineStyle swLineFontTangentEdgesStyle, swLinePHANTOM
    SetLineStyle swLineFontTangentEdgesThickness, swLW_THIN
    
    SetLineStyle swLineFontCosmeticThreadStyle, swLineCONTINUOUS
    SetLineStyle swLineFontCosmeticThreadThickness, swLW_THIN
    
    SetLineStyle swLineFontHideTangentEdgeStyle, swLineHIDDEN
    SetLineStyle swLineFontHideTangentEdgeThickness, swLW_THIN
    
    SetLineStyle swLineFontExplodedLinesStyle, swLineCHAINTHICK
    SetLineStyle swLineFontExplodedLinesThickness, swLW_THICK
    
    SetLineStyle swLineFontBreakLineStyle, swLineCONTINUOUS
    SetLineStyle swLineFontBreakLineThickness, swLW_THIN
    
    SetLineStyle swLineFontSpeedPakDrawingsModelEdgesStyle, swLineCONTINUOUS
    SetLineStyle swLineFontSpeedPakDrawingsModelEdgesThickness, swLW_NORMAL
    
    SetLineStyle swLineFontAdjoiningComponentStyle, swLineCENTER
    SetLineStyle swLineFontAdjoiningComponent, swLW_THIN
    
    SetLineStyle swLineFontBendLineUpStyle, swLinePHANTOM
    SetLineStyle swLineFontBendLineUpThickness, swLW_THIN
    
    SetLineStyle swLineFontBendLineDownStyle, swLinePHANTOM
    SetLineStyle swLineFontBendLineDownThickness, swLW_THIN
    
    SetLineStyle swLineFontEnvelopeComponentStyle, swLineCONTINUOUS
    SetLineStyle swLineFontEnvelopeComponentThickness, swLW_THIN
End Function
