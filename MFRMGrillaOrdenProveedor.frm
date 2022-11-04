Option Explicit
Public PrimeraActivacion         As Boolean
Dim Loc_Accion                   As String
Dim LOC_CampoOK                  As Boolean
Dim TxtSQL                       As String

Dim OBJ_BusLogic                 As CLS_Buslogic
Dim Uen                          As CLS_Uen
Dim OrdenProveedorCabeza         As CLS_OrdenProveedorCabeza

Dim RS_Registros                 As ADODB.Recordset

Dim Loc_Bookmark                 As Variant
Dim Loc_Identificacion_Bookmark  As Long
Dim Loc_Estado                   As String
Dim LOC_Baja                     As Boolean

Const Loc_ColNroOrden = 1
Const Loc_ColFecha = 2
Const Loc_ColProveedor = 3
Const Loc_ColAutorizo = 4
Const Loc_ColConfecciono = 5
Const Loc_ColBaja = 6
Const Loc_ColImpresa = 7
Const Loc_ColIdentificacion = 8

Private Sub Form_Load()
   SFRMContenedor.Visible = False
   SFRMBotones.Visible = False
End Sub

Private Sub Form_Activate()
   If PrimeraActivacion Then
      GLO_ClaveUEN = 0
      GLO_ClaveCuenta = 0
      Loc_Accion = GLO_Accion
      LOC_CampoOK = True
      CrearObjetos
      FSPRGrilla.RowHeight(-1) = GLO_AlturaFila
      Select Case Loc_Accion
           Case "S"
                   SCMDAgregar.Enabled = False
                   SCMDModificar.Enabled = False
                   SCMDBorrar.Enabled = False
                   SCMDSalir.Enabled = False
           Case Else
                   SCMDAgregar.Enabled = True
                   SCMDModificar.Enabled = True
                   SCMDBorrar.Enabled = True
                   SCMDSalir.Enabled = True
      End Select
      InicializarControles

      PrimeraActivacion = False
      SFRMContenedor.Visible = True
      SFRMBotones.Visible = True
      FSPRGrilla.SetFocus
   End If
End Sub

Private Sub CrearObjetos()
   'Crear un Objeto BusLogic
   Set OBJ_BusLogic = New CLS_Buslogic
   Set Uen = New CLS_Uen
   Set OrdenProveedorCabeza = New CLS_OrdenProveedorCabeza
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   'Transformar el {TAB} en {ENTER}
   Tecla_ENTER (KeyAscii)
End Sub

Private Sub FPDTFechaDesde_Validate(Cancel As Boolean)
   If Not FPDTFechaDesde.IsValid Or FPDTFechaDesde.Text = "" Then
      MsgBox "Fecha Inválida, Ingrese el Dato Nuevamente", vbInformation + vbOKOnly, "Atención"
      Cancel = True
   End If
End Sub

Private Sub FPDTFechaHasta_Validate(Cancel As Boolean)
   If Not FPDTFechaHasta.IsValid Or FPDTFechaHasta.Text = "" Then
      MsgBox "Fecha Inválida, Ingrese el Dato Nuevamente", vbInformation + vbOKOnly, "Atención"
      Cancel = True
   End If
End Sub

Private Sub FPLIUen_GotFocus()
   FSPRGrilla.MaxRows = 0
End Sub

Private Sub FPLIUen_Validate(Cancel As Boolean)
   If FPLIUen = 0 Then
      FTXTNombreuen = "Todas las Uen"
   Else
      TxtSQL = "Select UEN_NOMBRE From UEN Where UEN_CODIGO = " & FPLIUen.Value
      If OBJ_BusLogic.Get_RS_ADO_Aux(RS_Registros, TxtSQL, adUseClient, adOpenStatic, adLockReadOnly, False) Then
         If RS_Registros.EOF Then
            GLO_Accion = "S"
            MFRMGrillaUen.PrimeraActivacion = True
            MFRMGrillaUen.Show vbModal
            FPLIUen = MFRMGrillaUen.Numerouen
            FTXTNombreuen.Text = MFRMGrillaUen.NombreUen
            Unload MFRMGrillaUen
          Else
            FTXTNombreuen = RS_Registros!UEN_NOMBRE
          End If
      End If
    End If
End Sub

Private Sub FSPRGrilla_DblClick(ByVal Col As Long, ByVal row As Long)
   SCMDModificar_Click
End Sub

Private Sub FSPRGrilla_LostFocus()
   If Me.ActiveControl.Name <> "SCMDSalir" And LOC_CampoOK Then
      SCMDSalir.Default = False
      SCMDModificar.Default = False
   End If
End Sub

Private Sub FSPRGrilla_GotFocus()
   If LOC_CampoOK Then
      SCMDModificar.Default = True
   End If
   
   If ArmarTextoSelect = True Then
      FSPRGrilla.MaxRows = RS_Registros.RecordCount
      Set FSPRGrilla.DataSource = RS_Registros
      Call Setear_BookMark(RS_Registros, Loc_Bookmark, "D", GLO_Accion, Loc_Identificacion_Bookmark)
          
      FSPRGrilla.ReDraw = True
      FSPRGrilla.Refresh
         
      If RS_Registros.EOF Then
         SCMDAgregar.Enabled = True
         SCMDSalir.Enabled = True
      End If
   Else
      MsgBox "Error al Conectar con la Base de Datos, Avise al Sector Sistemas.-", vbCritical + vbOKOnly, "Atención"
   End If
End Sub

Private Sub SCMDAgregar_Click()
   If LOC_CampoOK Then
      LOC_CampoOK = False
      GLO_Accion = "A"
      MFRMOrdenCompraProveedor.PrimeraActivacion = True
      MFRMOrdenCompraProveedor.Show vbModal
'      If MFRMOrdenCompraProveedor.GraboRegistro = True Then
'               InicializaGrilla
'      End If
      Unload MFRMOrdenCompraProveedor
'       MFRMOrdenCompraProveedor.PrimeraActivacion = True
'       MFRMOrdenCompraProveedor.Show vbModal
'       Unload MFRMOrdenCompraProveedor
'       Loc_Identificacion_Bookmark = MFRMOrdenCompraProveedor.Loc_Identificador
      FSPRGrilla.SetFocus
      LOC_CampoOK = True
   End If
End Sub

Private Sub SCMDBorrar_Click()
     
   If InicializarGlobales Then
      Call Setear_BookMark(RS_Registros, Loc_Bookmark, "G")
      GLO_Accion = "B"
      If LOC_Baja Then
         LOC_Baja = False
         MsgBox "La orden de carga fue dada de baja, No se puede modificar", vbInformation + vbOKOnly, "Atención"
      Else
         If Loc_Estado = "I" Then
            MsgBox "La orden de carga ya fue impresa, No puede ser modificada", vbInformation + vbOKOnly, "Atención"
         Else
            MFRMOrdenCompraProveedor.PrimeraActivacion = True
            MFRMOrdenCompraProveedor.Show vbModal
            If MFRMOrdenCompraProveedor.GraboRegistro = True Then
               InicializaGrilla
            End If
            Unload MFRMOrdenCompraProveedor
         End If
      End If
      
'      If GLO_ClaveUEN <> GLO_CodigoUen Then
'         MsgBox "No se puede borrar una orden de carga hecha en otra U.E.N.", vbInformation + vbOKOnly, "Atención"
'      Else
'         If LOC_Baja Then
'            MsgBox "La orden de carga fue dada de baja", vbInformation + vbOKOnly, "Atención"
'         Else
'            If Loc_Estado = "I" Then
'               GLO_Respuesta = MsgBox("La orden de carga ya fue impresa, desea borrarla de todas maneras?", vbInformation + vbYesNo + vbDefaultButton2, "Atención")
'               If GLO_Respuesta = vbYes Then
'                  MFRMOrdenCompraProveedor.PrimeraActivacion = True
'                  MFRMOrdenCompraProveedor.Show vbModal
'                  Unload MFRMOrdenCompraProveedor
'               End If
'            Else
'               MFRMOrdenCompraProveedor.PrimeraActivacion = True
'               MFRMOrdenCompraProveedor.Show vbModal
'               Unload MFRMOrdenCompraProveedor
'            End If
'         End If
'      End If
   End If
   FSPRGrilla.SetFocus
End Sub

Private Sub InicializaGrilla()
   'Loc_ActualizaGrilla = True
   FSPRGrilla.MaxRows = 0
End Sub

Private Sub SCMDConsultar_Click()
   If InicializarGlobales() Then
      Call Setear_BookMark(RS_Registros, Loc_Bookmark, "G")
      GLO_Accion = "C"
      If LOC_Baja Then
         LOC_Baja = False
         MsgBox "La orden de carga fue dada de baja, No se puede modificar", vbInformation + vbOKOnly, "Atención"
      Else
         If Loc_Estado = "I" Then
            MsgBox "La orden de carga ya fue impresa, No puede ser modificada", vbInformation + vbOKOnly, "Atención"
         Else
            MFRMOrdenCompraProveedor.PrimeraActivacion = True
            MFRMOrdenCompraProveedor.Show vbModal
            Unload MFRMOrdenCompraProveedor
         End If
      End If
   End If
   FSPRGrilla.SetFocus
End Sub

Private Sub SCMDFechaDesde_GotFocus()
   FPDTFechaDesde.SetFocus
End Sub

Private Sub SCMDFechaHasta_GotFocus()
   FPDTFechaHasta.SetFocus
End Sub

Private Sub SCMDImprimir_Click()
On Error GoTo ManejadorError
   Call Setear_BookMark(RS_Registros, Loc_Bookmark, "G")
   If InicializarGlobales() Then
      If LOC_Baja Then
         MsgBox "La orden de carga fue dada de baja, No se puede imprimir", vbInformation + vbOKOnly, "Atención"
      Else
         If Loc_Estado = "I" Then
            MsgBox "La orden de carga ya fue impresa, No se puede reimprimir", vbInformation + vbOKOnly, "Atención"
         Else
            'RegistrarDataSource
'            TxtSQL = "Exec List_OrdenCompraProveedor " & GLO_ClaveIdCodigo
'            If Not OBJ_BusLogic.Get_RS_ADO_Aux(RS_DatosReporte, TxtSQL, adUseServer, adOpenStatic, adLockReadOnly, False) Then
'               Exit Sub
'            End If
'            Call Imprimir_CR10(GLO_PathReportes + "OrdenCompraProveedor.rpt", RS_DatosReporte, "I", 2, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, 0, 0, "", "S")

            ModificaEstado

         End If
      End If
   End If
   FSPRGrilla.SetFocus
   Exit Sub
ManejadorError:
        Set GLO_ADO_Parametro = Nothing
        MsgBox "Error al Intentar Imprimir la Orden Compra a Proveedor - " & vbCr & "Función Imprimir " & vbCr & "Error Nº: " & CStr(Err.Number) & "-" & Err.Description, vbOKOnly + vbCritical, "Atención"
End Sub



'Private Sub SCMDImprimir_Click()
'On Error GoTo ManejadorError
'   Call Setear_BookMark(RS_Registros, Loc_Bookmark, "G")
'   If InicializarGlobales() Then
'      If GLO_ClaveUEN <> GLO_CodigoUen Then
'         MsgBox "No se puede imprimir una orden de carga hecha en otra U.E.N.", vbInformation + vbOKOnly, "Atención"
'      Else
'         If LOC_Baja Then
'            MsgBox "La orden de carga fue dada de baja, No se puede imprimir", vbInformation + vbOKOnly, "Atención"
'         Else
'            If Loc_Estado = "I" Then
'               MsgBox "La orden de carga ya fue impresa, No se puede reimprimir", vbInformation + vbOKOnly, "Atención"
'            Else
'               RegistrarDataSource
'               TxtSQL = "Exec List_ComprobanteOrdenCargaCombustible " & GLO_ClaveUEN & ", " & GLO_ClaveCuenta
'               If Not OBJ_BusLogic.Get_RS_ADO_Aux(RS_DatosReporte, TxtSQL, adUseServer, adOpenStatic, adLockReadOnly, False) Then
'                  Exit Sub
'               End If
'               Call Imprimir_CR10(GLO_PathReportes + "ComprobanteOrdenCargaCombustible.rpt", RS_DatosReporte, "I", 2, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, 0, 0, "", "S")
'
'               ModificaEstado
'
'            End If
'         End If
'      End If
'   End If
'   FSPRGrilla.SetFocus
'   Exit Sub
'ManejadorError:
'        Set GLO_ADO_Parametro = Nothing
'        MsgBox "Error al Intentar Imprimir la Orden Carga de Combustible - " & vbCr & "Función Imprimir " & vbCr & "Error Nº: " & CStr(Err.Number) & "-" & Err.Description, vbOKOnly + vbCritical, "Atención"
'End Sub

Private Sub SCMDModificar_Click()
   If InicializarGlobales() Then
      Call Setear_BookMark(RS_Registros, Loc_Bookmark, "G")
      GLO_Accion = "M"
      If LOC_Baja Then
         LOC_Baja = False
         MsgBox "La orden de carga fue dada de baja, No se puede modificar", vbInformation + vbOKOnly, "Atención"
      Else
         If Loc_Estado = "I" Then
            MsgBox "La orden de carga ya fue impresa, No puede ser modificada", vbInformation + vbOKOnly, "Atención"
         Else
            MFRMOrdenCompraProveedor.PrimeraActivacion = True
            MFRMOrdenCompraProveedor.Show vbModal
            If MFRMOrdenCompraProveedor.GraboRegistro = True Then
               InicializaGrilla
            End If
            Unload MFRMOrdenCompraProveedor
         End If
      End If
   End If
  FSPRGrilla.SetFocus
End Sub

Private Function InicializarGlobales() As Boolean
   Dim Contenido   As Variant
   Dim Retorno     As Boolean

   InicializarGlobales = False
   GLO_ClaveIdCodigo = 0

   If FSPRGrilla.MaxRows > 0 Then
      Retorno = FSPRGrilla.GetText(Loc_ColNroOrden, FSPRGrilla.ActiveRow, Contenido)
      If Retorno Then
         GLO_ClaveIdCodigo = CLng(Contenido)
         Retorno = FSPRGrilla.GetText(Loc_ColIdentificacion, FSPRGrilla.ActiveRow, Contenido)
         If Retorno Then
            Loc_Identificacion_Bookmark = Contenido
            Retorno = FSPRGrilla.GetText(Loc_ColBaja, FSPRGrilla.ActiveRow, Contenido)
            If Retorno Then
               If CInt(Contenido) = 1 Then
                  LOC_Baja = True
               Else
                  LOC_Baja = False
               End If
               Retorno = FSPRGrilla.GetText(Loc_ColImpresa, FSPRGrilla.ActiveRow, Contenido)
               If Retorno Then
                  If CInt(Contenido) = 1 Then
                     Loc_Estado = "I"
                  Else
                     Loc_Estado = "N"
                  End If
               End If
            End If
         End If
         InicializarGlobales = Retorno
      End If
   End If

End Function

Private Sub SCMDPortaPapeles_Click()
   FSPRGrilla.OperationMode = OperationModeNormal
   FSPRGrilla.row = 0
   FSPRGrilla.Col = 1
   FSPRGrilla.row2 = FSPRGrilla.MaxRows
   FSPRGrilla.col2 = FSPRGrilla.MaxCols
   FSPRGrilla.Action = ActionSelectBlock '2
   FSPRGrilla.Action = ActionClipboardCopy '14
   FSPRGrilla.OperationMode = OperationModeRow
   FSPRGrilla.SetFocus
End Sub

Private Sub SCMDProveedor_GotFocus()
   FTXTProveedor.SetFocus
End Sub

Private Sub SCMDSalir_Click()
   'Borro las firmas de la carpeta Reportes
   
   If LOC_CampoOK Then
       LOC_CampoOK = False
       Set OBJ_BusLogic = Nothing
       Set RS_Registros = Nothing
       Salida
       LOC_CampoOK = True
   End If
End Sub

Private Sub Salida()
   GLO_Accion = Loc_Accion
   Me.Hide
End Sub

Private Function ArmarTextoSelect() As Boolean
   Dim TxtUen          As String
   ArmarTextoSelect = False
   TxtSQL = "Exec Cons_OrdenProveedorCabeza " & CInt(FPLIUen.Value) & ", '%" & Trim(FTXTProveedor) & "%', " & Format(FPDTFechaDesde, "yyyymmdd") & ", " & Format(FPDTFechaHasta, "yyyymmdd") '& ", " & CInt(FSPRGrilla.ActiveRow)
   If OBJ_BusLogic.Get_RS_ADO_Aux(RS_Registros, TxtSQL, adUseClient, adOpenStatic, adLockReadOnly, False) Then
      ArmarTextoSelect = True
   End If
End Function

Private Sub InicializarControles()
   If GLO_Unidad = "C" Then
      If FPLIUen.Value = 0 Then
         FTXTNombreuen = "Todas las Uen"
      End If
   Else
      FPLIUen.Value = GLO_CodigoUen
      If Uen.Get_Uen(FPLIUen.Value, False) Then
         FTXTNombreuen.Text = Uen.Nombre
      End If
      FPLIUen.ControlType = ControlTypeStatic
      FPLIUen.BackColor = GLO_BackColorVerde
   End If
   FPDTFechaHasta = Date
   FPDTFechaDesde.Value = FPDTFechaHasta.Value - 30
End Sub

Private Function ModificaEstado()
   If OBJ_BusLogic.ConectarASql(True) Then
      Screen.MousePointer = vbHourglass
      GLO_ADO_Conexion.BeginTrans
      If GrabarRegistro Then
         ' Confirma Transacción
         GLO_ADO_Conexion.CommitTrans
'         FSPRGrilla.SetText Loc_ColEstado, FSPRGrilla.ActiveRow, "I" '1
      Else
         ' Deshace Transacción
         GLO_ADO_Conexion.RollbackTrans
      End If
      If Not OBJ_BusLogic.DesconectarDeSql() Then
      End If
      Screen.MousePointer = vbNormal
   End If
End Function
'
'Private Function GrabarRegistro() As Boolean
'On Error GoTo ManejadorError
'               GrabarRegistro = False
'               Set GLO_ADO_Comando = Nothing
'               Set GLO_ADO_Comando = New ADODB.Command
'               GLO_ADO_Comando.ActiveConnection = GLO_ADO_Conexion
'
'               GLO_ADO_Comando.CommandText = "Modi_OrdenCargaCombustibleEstado"
'               GLO_ADO_Comando.CommandType = adCmdStoredProc
'
''              @SUCURSALCOMPROBANTE      smallint,
'               Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("SUCURSALCOMPROBANTE", adInteger, adParamInput, , GLO_CodigoUen)
'               GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
''              @NUMEROCOMPROBANTE      int,
'               Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("NUMEROCOMPROBANTE", adInteger, adParamInput, , GLO_ClaveCuenta)
'               GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
''              @ESTADO
'               Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("ESTADO", adVarChar, adParamInput, 1, "I")
'               GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
''              @OPERACION                   varchar (1),
'               Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("OPERACION", adVarChar, adParamInput, 1, "M")
'               GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
''              @SUCURSAL                    smallint
'               Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("SUCURSAL", adInteger, adParamInput, , GLO_CodigoUen)
'               GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
'
'               GLO_ADO_Comando.Execute
'               Set GLO_ADO_Parametro = Nothing
'               GrabarRegistro = True
'               Exit Function
'ManejadorError:
'        Set GLO_ADO_Parametro = Nothing
'        MsgBox "Error al Intentar Imprimir la Orden Carga de Combustible - " & vbCr & "Función Imprimir " & vbCr & "Error Nº: " & CStr(Err.Number) & "-" & Err.Description, vbOKOnly + vbCritical, "Atención"
'
'End Function



Private Sub SCMDUen_GotFocus()
   If FPLIUen.ControlType = ControlTypeNormal Then
      FPLIUen.SetFocus
   Else
      FSPRGrilla.SetFocus
   End If
End Sub

Private Function LeerRegistro() As Boolean
   LeerRegistro = False
   If OrdenProveedorCabeza.Get_OrdenProveedorCabeza(GLO_ClaveIdCodigo, False) Then
      If Uen.Get_Uen(OrdenProveedorCabeza.Uen, False) Then
         LeerRegistro = True
      End If
   End If
End Function

Private Function GrabarRegistro() As Boolean

   GrabarRegistro = False
   
   If LeerRegistro Then
      OrdenProveedorCabeza.Impresa = "S"
      If OrdenProveedorCabeza.Modi_OrdenProveedorCabeza Then
         GrabarRegistro = True
      End If
   End If
   Exit Function

End Function
