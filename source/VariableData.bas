Attribute VB_Name = "VariableData"
'===============================================================================
'   Макрос          : VariableData
'   Версия          : 2023.03.13
'   Сайты           : https://vk.com/elvin_macro
'                     https://github.com/elvin-nsk
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

Public Const RELEASE As Boolean = True

Public Const APP_NAME As String = "VariableData"

'===============================================================================

Private Const ExportAsExt As String = "jpg"

'===============================================================================

Sub Start()

    If RELEASE Then On Error GoTo Catch
    
    Dim Source As InitData
    Set Source = InitData.GetDocumentOrPage
    If Source.IsError Then Exit Sub
    
    Dim Cfg As Config
    Set Cfg = Config.Bind
    CheckCfg Cfg
    
    Dim Table As Dictionary
    Set Table = _
        FileToKeyedColumns(Cfg.TableFile, Cfg.CsvCharset, Cfg.CsvSeparator)
    
    Application.Optimization = RELEASE
    
    MainRoutine Source.Page, Table, Cfg
    
Finally:
    Application.Optimization = False
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally

End Sub

'===============================================================================

Private Sub MainRoutine( _
                ByVal Page As Page, _
                ByVal TableDic As Dictionary, _
                ByVal Cfg As Config _
            )
    Dim TotalRows As Long
    TotalRows = TableDic(TableDic.Keys(0)).Count
    If TotalRows < 1 Then Throw "Пустая таблица"
    Dim PBar As IProgressBar
    Set PBar = ProgressBar.CreateNumeric(TotalRows)
    Dim Doc As Document
    Dim Name As String
    Dim Row As Long
    For Row = 1 To TotalRows
        Set Doc = Page.Shapes.All.CreateDocumentFrom
        Doc.Unit = cdrPixel
        ProcessDocument Doc.ActivePage, TableDic, Row, Cfg
        Name = FileNameFromCell(TableDic(Cfg.FileNameColumn)(Row))
        With FileSpec.Create
            .NameWithoutExt = Name
            .Ext = ExportAsExt
            .Path = Cfg.TargetFolder
            .Path = MakeDir(.Path & Name)
            With Doc.ExportBitmap( _
                     .ToString, cdrJPEG, cdrCurrentPage, cdrRGBColorImage, _
                     Doc.ActivePage.SizeWidth, Doc.ActivePage.SizeHeight _
                 )
                 .Compression = Cfg.Compression
                 .Finish
            End With
        End With
        PBar.Update
        Doc.Close
    Next Row
End Sub

Private Property Get FileNameFromCell(ByVal Cell As String) As String
    Dim Arr As Variant
    Arr = VBA.Split(Cell, "\")
    FileNameFromCell = Arr(UBound(Arr))
End Property

Private Sub ProcessDocument( _
                ByVal Page As Page, _
                ByVal TableDic As Dictionary, _
                ByVal Row As Long, _
                ByVal Cfg As Config _
            )
    Dim ShapesDic As Dictionary
    Set ShapesDic = FindShapesByNames(Page.Shapes.All, TableDic.Keys)
    Dim Shapes As ShapeRange
    Dim Tag As Variant
    Dim Shape As Shape
    For Each Tag In ShapesDic.Keys
        For Each Shape In ShapesDic(Tag)
            ProcessShape Shape, TableDic(Tag)(Row), Cfg
        Next Shape
    Next Tag
End Sub

Private Function ProcessShape( _
                     ByVal Shape As Shape, _
                     ByVal Data As String, _
                     ByVal Cfg As Config _
                 )
    If Shape.Type = cdrTextShape Then
        Shape.Text.Story.Text = Data
    Else
        ImportAndComposeImages Shape, Data, Cfg
    End If
End Function

Private Function ImportAndComposeImages( _
                     ByVal Shape As Shape, _
                     ByVal Path As String, _
                     ByVal Cfg As Config _
                 )
    Dim FSO As New FileSystemObject
    If Not FSO.FolderExists(Path) Then
        If Not RELEASE Then Show "Не найдён путь " & Path
        Exit Function
    End If
    Dim Shapes As New ShapeRange
    Dim File As Scripting.File
    For Each File In FSO.GetFolder(Path).Files
        Shape.Layer.Import File.Path
        Shapes.Add ActiveShape
        ComposeImages Shapes, Shape.BoundingBox, Cfg
    Next File
End Function

Private Function ComposeImages( _
                     ByVal Shapes As ShapeRange, _
                     ByVal Box As Rect, _
                     ByVal Cfg As Config _
                 )
    If Shapes.Count = 1 Then
        FitInside Shapes.FirstShape, Box
        Exit Function
    End If
    Dim QuarterSize As Rect
    Set QuarterSize = CalcQuarterSize(Box, Cfg.Space)
    Dim Index As Long
    For Index = 1 To Shapes.Count
        Shapes(Index).SizeWidth = QuarterSize.Width
        Shapes(Index).SizeHeight = QuarterSize.Height
        Select Case Index
            Case 1
                Shapes(Index).LeftX = Box.Left
                Shapes(Index).TopY = Box.Top
            Case 2
                Shapes(Index).LeftX = Box.Left + QuarterSize.Width + Cfg.Space
                Shapes(Index).TopY = Box.Top
            Case 3
                Shapes(Index).LeftX = Box.Left
                Shapes(Index).TopY = Box.Top - QuarterSize.Height - Cfg.Space
            Case 4
                Shapes(Index).LeftX = Box.Left + QuarterSize.Width + Cfg.Space
                Shapes(Index).TopY = Box.Top - QuarterSize.Height - Cfg.Space
        End Select
    Next Index
End Function

Private Function CalcQuarterSize( _
                     ByVal Box As Rect, _
                     ByVal Space As Double _
                 ) As Rect
    Set CalcQuarterSize = CreateRect
    CalcQuarterSize.Width = CalcHalfLength(Box.Width, Space)
    CalcQuarterSize.Height = CalcHalfLength(Box.Height, Space)
End Function

Private Function CalcHalfLength( _
                     ByVal Length As Double, _
                     ByVal Space As Double _
                 ) As Double
    CalcHalfLength = (Length - Space) / 2
End Function

Private Sub CheckCfg(ByVal Cfg As Config)
    If Not FileExists(Cfg.TableFile) Then _
        Throw "Не найден файл таблицы " & Cfg.TableFile
    If Not FSO.FolderExists(Cfg.TargetFolder) Then _
        Throw "Не найдена целевая папка для сохранения " & Cfg.TargetFolder
End Sub
