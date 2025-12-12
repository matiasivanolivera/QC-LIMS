Attribute VB_Name = "exportarTodoElVBA"
Sub exportarTodoElVBA()

    Dim comp As Object
    Dim ruta As String
    Dim nombre As String
    
    ' Carpeta donde se exportarán los archivos
    ruta = ThisWorkbook.Path & "\ExportVBA\"
    
    ' Crear carpeta si no existe
    If Dir(ruta, vbDirectory) = "" Then
        MkDir ruta
    End If

    ' Recorrer cada componente del proyecto VBA
    For Each comp In ThisWorkbook.VBProject.VBComponents
        
        Select Case comp.Type
            
            Case 1 ' Módulo estándar (.BAS)
                nombre = comp.Name & ".bas"
                comp.Export ruta & nombre
                
            Case 2 ' Módulo de clase (.CLS)
                nombre = comp.Name & ".cls"
                comp.Export ruta & nombre
                
            Case 3 ' Formulario (.FRM + .FRX)
                nombre = comp.Name & ".frm"
                comp.Export ruta & nombre
                
            Case Else
                ' Otros tipos no exportables
        End Select

    Next comp

    MsgBox "Exportación completa en: " & ruta, vbInformation

End Sub

