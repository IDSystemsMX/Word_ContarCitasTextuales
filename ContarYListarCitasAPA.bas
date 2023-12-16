'WordMacro CONTARYLISTARCITASAPA
'Por: Mtro. LSC. Sergio Hugo Sanchez Olivares
'Fecha: 16/Dic/2023
' 
' Cuenta las citas textuales en formato APA 7a. edicion.
' La estructura es: (Autor, Fecha) o (Autor, Fecha, pag 0)
' Si la cita esta mal escrita segun esta estructura en el documento, no la contara.
'

Sub ContarYListarCitasAPA()
    Dim contadorCitas As Integer
    Dim listaCitas As String
    contadorCitas = 0
    listaCitas = ""
    
    ' Expresión regular para buscar citas APA
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "\(\w+,\s\d{4}(, pag \d+)?\)"
    
    ' Recorre todos los párrafos del documento
    For Each p In ActiveDocument.Paragraphs
        ' Busca coincidencias con la expresión regular en el texto del párrafo
        If regex.Test(p.Range.Text) Then
            ' Incrementa el contador
            contadorCitas = contadorCitas + 1
            
            ' Agrega la cita a la lista con el número de página
            listaCitas = listaCitas & "Cita " & contadorCitas & ": " & " (Página " & p.Range.Information(wdActiveEndPageNumber) & ")" & vbCrLf
        End If
    Next p
    
    ' Muestra el resultado en una caja de diálogo
    MsgBox "Total de citas APA encontradas: " & contadorCitas & vbCrLf & vbCrLf & "Lista de Citas:" & vbCrLf & listaCitas, vbInformation, "Contador y Lista de Citas APA"
End Sub


