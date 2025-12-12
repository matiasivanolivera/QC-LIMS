Attribute VB_Name = "abreFromDsdExcel"
Sub RevisarActividadesPendientes_OT()

    Load frmOTReview              ' Carga el form en memoria
    frmOTReview.CargarActividades ' Parser + maestros + combo + ListView
    frmOTReview.Show              ' Mostrar ya lleno

End Sub


