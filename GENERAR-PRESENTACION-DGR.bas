Attribute VB_Name = "Módulo1"
Option Explicit
Sub planillaExtractosBancarios()

'--------- IMPORTANTE ---------
'   PARA CADA CARGA SIEMPRE
'   REVISAR LAS COLUMNAS
'   DEL EXTRACTO BANCARIO
'   Y QUE COINCIDAN
'   CON LAS DEL MACRO
'----------           ---------


Dim r As Range
Dim wb As Workbook, ws As Worksheet, ivaWs As Worksheet
Dim i& ' --> Lo mismo que decir "Dim i As Long"
Dim checkSheetName As String

Set wb = ThisWorkbook
Set ws = wb.Worksheets(1)

'Creo una nueva hoja para los datos que me interesan, si ya existe no se crea

On Error Resume Next
checkSheetName = wb.Worksheets(2).Name
If checkSheetName = "Carga HOLISTOR" Then
Else
    Worksheets.Add(After:=wb.Worksheets(1)).Name = "Carga HOLISTOR"
End If

Set ivaWs = wb.Worksheets(2)

ivaWs.Cells(1, 1) = "Fecha"
ivaWs.Cells(1, 2) = "Concepto"
ivaWs.Cells(1, 3) = "Importe"

'Buso el último elemento del extracto bancario

Dim finEx As Long
finEx = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

'Creo un índice de referencia de la siguiente fila a llenar para la segunda hoja

Dim ultPos As Integer
ultPos = 2

' Marcado de las filas que contienen IVA o Percepción IVA
Debug.Print ""
Debug.Print ""
Debug.Print ""
Debug.Print ""

For i = 1 To finEx
    Set r = ws.Cells(i, 4) 'O podría ser: Set r = ws.Range("C" & i)
    If Trim(r.MergeArea.Cells(1, 1).Value) Like "*Debito Fiscal 21*" Or Trim(r.MergeArea.Cells(1, 1).Value) Like "*Percepcion IVA*" Then
        ivaWs.Cells(ultPos, 1) = CDate(ws.Cells(i, 1))
        ivaWs.Cells(ultPos, 2) = ws.Cells(i, 4)
        ivaWs.Cells(ultPos, 3) = CCur(ws.Range("S" & i))
        ultPos = ultPos + 1
        'Range("A" & i & ":Y" & i).Interior.Color = RGB(240, 180, 200)
        Debug.Print r.MergeArea.Cells(1, 1).Value
    End If
Next i

'Armo una tabla con los datos que me interesan

Dim tbl As ListObject
finEx = ivaWs.Cells(ivaWs.Rows.Count, "A").End(xlUp).Row

Set tbl = ivaWs.ListObjects.Add( _
    SourceType:=xlSrcRange, _
    Source:=Range("A1:C" & finEx), _
    XlListObjectHasHeaders:=xlYes)

ivaWs.Cells.Columns.AutoFit

End Sub



