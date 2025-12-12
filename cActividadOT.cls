VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cActividadOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Ensayo As String
Public Muestra As String      ' lo vamos a usar para mostrar la especialidad
Public NPLote As String
Public tecnica As String
Public Fecha As Date
Public Analista As String
Public Hoja As String
Public Celda As String
Public TextoCrudo As String

' ?? NUEVOS CAMPOS PARA EL INDUSTRIAL
Public Especialidad As String
Public Presentacion As String
Public FormaFF As String      ' (C, CR, S, J, MP, etc.)
Public TipoProducto As String ' PT / PI / MP
Public aliasAuto As String
Public Variante As String


