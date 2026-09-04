Option Explicit
Attribute VB_Name = "Ret-Percp a DGR"
Sub GenerarTxtDGR()
    ' ==========================================================
    ' MACRO FINAL: EXPORTAR A DOS TXT
    ' - Retenciones: SIN número de comprobante
    ' - Percepciones: CON número de comprobante
    ' - VALORES: Respeta el signo negativo (para Notas de Crédito)
    ' ==========================================================
    
    ' --- ALÍCUOTAS ---
    Const ALICUOTA_PERCEPCION As String = "3.60"
    Const ALICUOTA_RETENCION As String = "2.50"
	    
    
    Dim ws As Worksheet
    Dim fso As Object
    
    ' Objetos para DOS archivos
    Dim fileRet As Object
    Dim filePer As Object
    Dim rutaRet As String
    Dim rutaPer As String
    
    Dim ultimaFila As Long, i As Long
    
    Dim selector As String
    
    
    ' --- CONFIGURACIÓN DE COLUMNAS ---
    Const COL_TIPO As String = "A"
    Const COL_DENOM As String = "B"
    Const COL_CUIT As String = "C"
    Const COL_CONCEPTO As String = "D"
    Const COL_ORDEN As String = "E"
    Const COL_NETO As String = "F"
    Const COL_IB As String = "G"
    
    ' Variables de datos
    Dim sTipo As String, sDenominacion As String, sCuit As String
    Dim sConcepto As String, vOrden As Variant, sOrdenLimpio As String
    Dim dNeto As Double, dMontoIB As Double
    
    ' Variables de texto final
    Dim txtTipoComp As String, txtPtoVenta As String, txtNumComp As String
    Dim txtNeto As String, txtAlicuota As String, txtIB As String
    Dim linea As String
    
    ' Constantes fijas
    Const SEPARADOR As String = vbTab
    
    Set ws = ActiveSheet
    
    ' 1. PREPARAR ARCHIVOS

	ultimaFila = ws.Range(COL_CUIT & ws.Rows.Count).End(xlUp).Row

	selector = InputBox("Ingrese fecha de período (dd/mm/aaaa):", "Fecha", Date)

	If selector = "" Then
	    Exit Sub
	End If

	While Not IsDate(selector)
	    MsgBox "Fecha Incorrecta", vbCritical
	    selector = InputBox("Ingrese fecha de período (dd/mm/aaaa):", "Fecha", Date)
	    
	    If selector = "" Then
		Exit Sub
	    End If
	Wend

	Dim FECHA_FIJA As String
	Dim FECHA_COMP As String

	FECHA_FIJA = selector
	FECHA_COMP = Mid(selector, 1, 2) & Mid(selector, 4, 2) & Mid(selector, 7, 4)

	rutaRet = ThisWorkbook.Path & "\DGR_Retenciones.txt"
	rutaPer = ThisWorkbook.Path & "\DGR_Percepciones.txt"

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set fileRet = fso.CreateTextFile(rutaRet, True)
	Set filePer = fso.CreateTextFile(rutaPer, True)
        
    ' --- INICIO DEL BUCLE ---
    For i = 2 To ultimaFila
        
        ' --- A. OBTENER DATOS ---
        sTipo = Trim(ws.Range(COL_TIPO & i).Value)
        sDenominacion = UCase(Trim(ws.Range(COL_DENOM & i).Value))
        sCuit = Trim(ws.Range(COL_CUIT & i).Value)
        sConcepto = Trim(ws.Range(COL_CONCEPTO & i).Value)
        vOrden = ws.Range(COL_ORDEN & i).Value
        
        ' Validar números (Neto e IB)
        If IsNumeric(ws.Range(COL_NETO & i).Value) And ws.Range(COL_NETO & i).Value <> "" Then
            dNeto = CDbl(ws.Range(COL_NETO & i).Value)
        Else
            dNeto = 0
        End If
        
        If IsNumeric(ws.Range(COL_IB & i).Value) And ws.Range(COL_IB & i).Value <> "" Then
            dMontoIB = CDbl(ws.Range(COL_IB & i).Value)
        Else
            dMontoIB = 0
        End If
        
        ' --- B. DAR FORMATO GENERAL ---
        
        ' Tipo Comprobante (01 o 03)
        If InStr(1, sTipo, "Factura", vbTextCompare) > 0 Then
            txtTipoComp = "01"
        ElseIf InStr(1, sTipo, "Nota", vbTextCompare) > 0 Or InStr(1, sTipo, "Crédito", vbTextCompare) > 0 Then
            txtTipoComp = "03"
        Else
            txtTipoComp = "01"
        End If
        
        ' Nro Orden y Comprobante (SOLO SE USA EN PERCEPCIONES)
        If IsNumeric(vOrden) And vOrden <> "" Then
            sOrdenLimpio = CStr(vOrden)
        Else
            sOrdenLimpio = "0"
        End If
        txtPtoVenta = Right("00000" & sOrdenLimpio, 5)
        txtNumComp = txtPtoVenta & "-" & FECHA_COMP
        
        ' --- ACÁ ESTÁ EL CAMBIO PARA LOS NEGATIVOS ---
        ' Eliminé la función Abs(). Ahora respeta el signo que tenga la celda.
        txtNeto = Replace(Format(dNeto, "0.00"), ",", ".")
        txtIB = Replace(Format(dMontoIB, "0.00"), ",", ".")
        
        ' Alícuota
	If InStr(1, sConcepto, "Percepción", vbTextCompare) > 0 Then
	    txtAlicuota = ALICUOTA_PERCEPCION
	ElseIf InStr(1, sConcepto, "Retención", vbTextCompare) > 0 Then
	    txtAlicuota = ALICUOTA_RETENCION
	Else
	    txtAlicuota = "0.00"
	End If
        
        ' --- C. ESCRITURA CONDICIONAL ---
        
        If InStr(1, sConcepto, "Retención", vbTextCompare) > 0 Then
            ' === FORMATO RETENCIONES ===
            ' SIN el campo de número de comprobante/pto venta
            linea = sCuit & SEPARADOR & _
                    sDenominacion & SEPARADOR & _
                    FECHA_FIJA & SEPARADOR & _
                    txtTipoComp & SEPARADOR & _
                    txtNeto & SEPARADOR & _
                    txtAlicuota & SEPARADOR & _
                    txtIB & SEPARADOR & _
                    "0.00"
            
            fileRet.WriteLine linea
            
        ElseIf InStr(1, sConcepto, "Percepción", vbTextCompare) > 0 Then
            ' === FORMATO PERCEPCIONES ===
            ' CON el campo de número de comprobante
            linea = sCuit & SEPARADOR & _
                    sDenominacion & SEPARADOR & _
                    FECHA_FIJA & SEPARADOR & _
                    txtTipoComp & SEPARADOR & _
                    txtNumComp & SEPARADOR & _
                    txtNeto & SEPARADOR & _
                    txtAlicuota & SEPARADOR & _
                    txtIB & SEPARADOR & _
                    "0.00"
            
            filePer.WriteLine linea
        End If
        
    Next i
    
    ' --- CIERRE ---
    fileRet.Close
    filePer.Close
    
    Set fileRet = Nothing
    Set filePer = Nothing
    Set fso = Nothing
    
    MsgBox "Proceso Terminado Exitosamente." & vbNewLine & _
           "Se han generado DGR_Retenciones.txt y DGR_Percepciones.txt" & vbNewLine & _
           "Los valores negativos se han mantenido.", vbInformation

End Sub
