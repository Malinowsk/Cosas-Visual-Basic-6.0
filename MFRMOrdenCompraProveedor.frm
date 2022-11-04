Option Explicit

Dim TxtSQL                          As String
Dim Loc_Accion                      As String
Dim RS_Registros                    As ADODB.Recordset

Public RDO_OrdenProveedorRenglon    As ADODB.Recordset
Public RDO_OrdenProveedorCC         As ADODB.Recordset

Public PrimeraActivacion            As Boolean
Dim Contenido                       As Variant
Dim Retorno                         As Boolean

Public GraboRegistro                As Boolean 'variable para saber si se grabo un registro con algun alta, modi y baja

Dim Uen                             As CLS_Uen
Dim MaestroCuentas                  As CLS_MaestroCuentas
Dim OrdenProveedorCabeza            As CLS_OrdenProveedorCabeza
Dim OrdenProveedorRenglon           As CLS_OrdenProveedorRenglon
Dim OrdenProveedorCC                As CLS_OrdenProveedorCC
Dim OBJ_BusLogic                    As CLS_Buslogic

Dim SumaPorcentaje                  As Double
Dim SeModificoGrillaArticulo        As Boolean
Dim SeModificoGrillaUenPorcentaje   As Boolean

'constantes
Const Loc_SolapaCargarArticulo = 0
Const Loc_SolapaCargarUen = 1

Const LOC_ColDescripcion = 1
Const Loc_ColCantidad = 2
Const Loc_ColUnidad = 3

Const Loc_ColUEN = 1
Const Loc_ColNombre = 2
Const Loc_ColPorcentaje = 3

Private Sub Form_Activate()
   If PrimeraActivacion Then
      SFRMBotones.Refresh
      SFRMContenedor.Refresh
      PrimeraActivacion = False
      Loc_Accion = GLO_Accion
      
      CrearObjetos
      InicializarControles
      FSPRGrilla.RowHeight(-1) = GLO_AlturaFila
      FSPRGrillaPorcentaje.RowHeight(-1) = GLO_AlturaFila

      Select Case Loc_Accion
         Case "A"
            InhabilitarCampos
         Case "B", "C", "M"
            If MoverCamposAControles Then
               InhabilitarCampos
            End If
      End Select
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Loc_Accion = "A" Or Loc_Accion = "M" Then
         Select Case KeyCode
            Case vbKeyF12  'Fin de editar renglones
               FPMEObservaciones.Enabled = True
               FPMEObservaciones.SetFocus
               LimpiarRenglon
               LimpiarRenglonUen
            Case vbKeyF11  'Moverse de solapa
               If STABGrillas.Tab = Loc_SolapaCargarArticulo Then
                  FPLIPorcentajeUen.SetFocus
               Else
                  FTXTArticulo.SetFocus
               End If
         End Select
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   'Transformo el {TAB} en {ENTER}
   Tecla_ENTER (KeyAscii)
End Sub

Private Sub InicializarControles()
   FPMEObservaciones.Enabled = False
   FPMEObservaciones.Text = ""
   FSPRGrilla.MaxRows = 0
   FSPRGrillaPorcentaje.MaxRows = 0
   SumaPorcentaje = 0
   SeModificoGrillaArticulo = False
   SeModificoGrillaUenPorcentaje = False
   STABGrillas.Tab = Loc_SolapaCargarArticulo
   FTXTCreo.Text = GLO_UsuarioNombre
   
   Select Case Loc_Accion
      Case "A", "M"
         MLBLSalirGrilla.Caption = "<Fin> Editar Renglon  <Ins> Insertar Renglón  <Supr> Borrar Renglón  <F11> Solapa Sig  <F12> Fin Ingresar Renglones"
         MLBLSalirGrillaPorcentaje.Caption = "<Ins> Insertar Uen  <Supr> Borrar Uen  <F11> Solapa Ant  <F12> Fin Ingresar Uen"
      Case "B", "C"
         MLBLSalirGrilla.Caption = ""
         MLBLSalirGrillaPorcentaje.Caption = ""
   End Select
   
   MFRALineaCargaRenglon.Enabled = True
   MFRABotonRenglon.Enabled = True
   LimpiarRenglon
   LimpiarRenglonUen
End Sub

Private Sub CrearObjetos()
   Set OBJ_BusLogic = New CLS_Buslogic
   Set Uen = New CLS_Uen
   Set MaestroCuentas = New CLS_MaestroCuentas
   Set OrdenProveedorCabeza = New CLS_OrdenProveedorCabeza
   Set OrdenProveedorRenglon = New CLS_OrdenProveedorRenglon
   Set OrdenProveedorCC = New CLS_OrdenProveedorCC
End Sub

Private Sub InhabilitarCampos()
   Select Case Loc_Accion
      Case "A", "M"
         FPLIProveedorCodigo.SetFocus
         FPLIProveedorUen.SetFocus
         SCMDGrabar.Caption = "&Grabar"
         If Loc_Accion = "M" Then
            FPMEObservaciones.Enabled = True
            If FPLIProveedorUen.Value = 0 Then
               FPLIProveedorCodigo.BackColor = GLO_BackColorVerde
               FPLIProveedorCodigo.ControlType = ControlTypeStatic
            Else
               FTXTProveedorNombre.BackColor = GLO_BackColorVerde
               FTXTProveedorNombre.ControlType = ControlTypeStatic
            End If
         End If
      Case "B", "C"
         FPLIProveedorUen.Enabled = False
         FPLIProveedorUen.BackColor = GLO_BackColorVerde
         FPLIProveedorCodigo.Enabled = False
         FPLIProveedorCodigo.BackColor = GLO_BackColorVerde
         FTXTProveedorNombre.Enabled = False
         FTXTProveedorNombre.BackColor = GLO_BackColorVerde
         FTXTRecibe.Enabled = False
         FTXTRecibe.BackColor = GLO_BackColorVerde
         FTXTAutoriza.Enabled = False
         FTXTAutoriza.BackColor = GLO_BackColorVerde
         FTXTCreo.Enabled = False
         FPLIId.Enabled = False
         FPDTFecha.Enabled = False
         MFRALineaCargaRenglon.Enabled = False
         MFRABotonRenglon.Enabled = False
         MFRALineaCargaUen.Enabled = False
         MFRABotonUen.Enabled = False
         FTXTArticulo.BackColor = GLO_BackColorVerde
         FPDSCantidad.BackColor = GLO_BackColorVerde
         FTXTUnidad.BackColor = GLO_BackColorVerde
         FPLIPorcentajeUen.BackColor = GLO_BackColorVerde
         FPDSPorcentajeGasto.BackColor = GLO_BackColorVerde
         If Loc_Accion = "B" Then
            SCMDSalir.SetFocus
            SCMDGrabar.Caption = "&Borrar"
         Else
            SCMDGrabar.Enabled = False
            FSPRGrilla.SetFocus
         End If

   End Select
End Sub

'uen proveedor
Private Sub FPLIProveedorUen_Change()
   FPLIProveedorCodigo.Value = 0
   FTXTProveedorNombre.Text = ""
   If FPLIProveedorCodigo.Enabled = False Then
      FPLIProveedorCodigo.Enabled = True
      FPLIProveedorCodigo.BackColor = GLO_BackColorBlanco
   End If
   If FTXTProveedorNombre.ControlType = ControlTypeStatic Then
      FTXTProveedorNombre.BackColor = GLO_BackColorBlanco
      FTXTProveedorNombre.ControlType = ControlTypeNormal
   End If
End Sub

Private Sub FPLIProveedorUen_Validate(Cancel As Boolean)
   If FPLIProveedorUen.Value <> 0 Then
      If Not Uen.Get_Uen(FPLIProveedorUen.Value, False) Then
         GLO_Accion = "S"
         MFRMGrillaUen.PrimeraActivacion = True
         MFRMGrillaUen.Show vbModal
         FPLIProveedorUen.Value = MFRMGrillaUen.Numerouen
         Unload MFRMGrillaUen
      End If
   Else
      FPLIProveedorCodigo.Enabled = False
      FPLIProveedorCodigo.BackColor = GLO_BackColorVerde
      FTXTProveedorNombre.SetFocus
   End If
End Sub

'codigo proveedor
Private Sub FPLIProveedorCodigo_Change()
   FTXTProveedorNombre.Text = ""
End Sub

Private Sub FPLIProveedorCodigo_Validate(Cancel As Boolean)
   If Not MaestroCuentas.Get_MaestroCuentas(FPLIProveedorUen.Value, FPLIProveedorCodigo.Value, False) Then
      GLO_Accion = "S"
      MFRMGrillaMaestroCuentas.PrimeraActivacion = True
      MFRMGrillaMaestroCuentas.FPLIUen.Value = Me.FPLIProveedorUen.Value
      MFRMGrillaMaestroCuentas.Show vbModal
      Unload MFRMGrillaMaestroCuentas
      FPLIProveedorUen.Value = GLO_ClaveUEN
      FPLIProveedorCodigo.Value = GLO_ClaveCuenta
      FTXTProveedorNombre.Text = GLO_RazonSocial
   Else
      FTXTProveedorNombre.Text = MaestroCuentas.RazonSocial
   End If
   FTXTProveedorNombre.BackColor = GLO_BackColorVerde
   FTXTProveedorNombre.ControlType = ControlTypeStatic
   FTXTRecibe.SetFocus
End Sub

'grilla articulo
Private Sub FSPRGrilla_KeyDown(KeyCode As Integer, Shift As Integer)
   If Loc_Accion = "A" Or Loc_Accion = "M" Then
      Select Case KeyCode
         Case vbKeyDelete                 'para borrar fila
            If BorrarRenglonEnGrilla Then
               FSPRGrilla.SetFocus
            End If
         Case vbKeyInsert                 'para insertar fila
            SCMDProcesaRenglon_Click
         Case vbKeyEnd                    'para editar fila
            EditarRenglon
      End Select
   End If
End Sub

'grilla gasto uen
Private Sub FSPRGrillaPorcentaje_KeyDown(KeyCode As Integer, Shift As Integer)
   If Loc_Accion = "A" Or Loc_Accion = "M" Then
      Select Case KeyCode
         Case vbKeyDelete                 'para borrar fila
            If BorrarUenEnGrilla Then
               FSPRGrillaPorcentaje.SetFocus
            End If
         Case vbKeyInsert                 'para insertar fila
               SCMDProcesaUen_Click
      End Select
   End If
End Sub

Private Sub FTXTProveedorNombre_Validate(Cancel As Boolean)
   If Len(Trim(FTXTProveedorNombre.Text)) = 0 Then
      MsgBox "Debe Cargar El Nombre del Proveedor.-", vbInformation + vbOKOnly, "Atención"
      Cancel = True
   End If
End Sub

Private Sub FTXTRecibe_Validate(Cancel As Boolean)
   If Len(Trim(FTXTRecibe.Text)) = 0 Then
      MsgBox "Debe Cargar a Quién Va Dirigida la Orden de Compra.-", vbInformation + vbOKOnly, "Atención"
      Cancel = True
   End If
End Sub

Private Sub FTXTAutoriza_Validate(Cancel As Boolean)
   If Len(Trim(FTXTAutoriza.Text)) = 0 Then
      MsgBox "Debe Cargar a Quién autoriza la Orden de Compra.-", vbInformation + vbOKOnly, "Atención"
      Cancel = True
   End If
End Sub

Private Sub FTXTArticulo_GotFocus()
      STABGrillas.Tab = Loc_SolapaCargarArticulo
End Sub

'uen porcentaje
Private Sub FPLIPorcentajeUen_Change()
   FTXTPorcentajeNombre.Text = ""
End Sub

Private Sub FPLIPorcentajeUen_GotFocus()
   STABGrillas.Tab = Loc_SolapaCargarUen
End Sub

Private Sub FPLIPorcentajeUen_Validate(Cancel As Boolean)
   If Loc_Accion = "A" Or Loc_Accion = "M" Then
      If SumaPorcentaje <> 100 Then
         If Uen.Get_Uen(FPLIPorcentajeUen.Value, False) Then
            FTXTPorcentajeNombre.Text = Uen.Nombre
         Else
            GLO_Accion = "S"
            MFRMGrillaUen.PrimeraActivacion = True
            MFRMGrillaUen.Show vbModal
            FPLIPorcentajeUen.Value = MFRMGrillaUen.Numerouen
            FTXTPorcentajeNombre.Text = MFRMGrillaUen.NombreUen
            Unload MFRMGrillaUen
         End If
      End If
   End If
End Sub

Private Sub SCMDGrabar_Click()
   'Con ADO
   If ValidarControles Then
      GraboRegistro = False
      SCMDGrabar.Enabled = False
      Screen.MousePointer = vbHourglass
      ' Abrir Transacción
      If Not OBJ_BusLogic.ConectarASql(True) Then
         Screen.MousePointer = vbNormal
         SCMDGrabar.Enabled = True
         Salida
      Else
         GLO_ADO_Conexion.BeginTrans
         If Not GrabarRegistro Then
            ' Anular Transacción
            GLO_ADO_Conexion.RollbackTrans
            MsgBox "Error al Grabar el Registro, Avise al Sector Sistemas.-", vbInformation + vbOKOnly, "Atención"
         Else
            ' Cerrar Transacción
            GLO_ADO_Conexion.CommitTrans
            GraboRegistro = True
            Salida
         End If
         Screen.MousePointer = vbNormal
         SCMDGrabar.Enabled = True
      End If
   End If
End Sub

Private Function ValidarControles() As Boolean
   ValidarControles = False
    If Loc_Accion <> "B" And Loc_Accion <> "C" Then
      'Reemplazo Los TAB POR UN " " en el campo Memo
      FPMEObservaciones.Text = Replace(FPMEObservaciones.Text, Chr(9), " ")
   
      If FPLIProveedorUen.Value <> 0 And Not Uen.Get_Uen(FPLIProveedorUen.Value, False) Then
            MsgBox "El Código de Uen No es Valido.-", vbInformation + vbOKOnly, "Atención"
            FPLIProveedorUen.SetFocus
            Exit Function
      Else
         If FPLIProveedorUen.Value <> 0 And Not MaestroCuentas.Get_MaestroCuentas(FPLIProveedorUen.Value, FPLIProveedorCodigo.Value, False) Then
            MsgBox "El Código de Provedor No es Correcto.-", vbInformation + vbOKOnly, "Atención"
            FPLIProveedorCodigo.SetFocus
            Exit Function
         End If
      End If
   
      If Len(Trim(FTXTProveedorNombre.Text)) = 0 Then
         MsgBox "Debe Cargar El Nombre del Proveedor.-", vbInformation + vbOKOnly
         If FTXTProveedorNombre.ControlType = ControlTypeNormal Then
            FTXTProveedorNombre.SetFocus
         Else
            FPLIProveedorUen.SetFocus
         End If
         Exit Function
      End If
      
      If Len(Trim(FTXTRecibe.Text)) = 0 Then
         MsgBox "Debe Cargar a Quién Va Dirigida la Orden de Compra.-", vbInformation + vbOKOnly
         FTXTRecibe.SetFocus
         Exit Function
      End If
      
      If Len(Trim(FTXTAutoriza.Text)) = 0 Then
         MsgBox "Debe Cargar a Quién autoriza la Orden de Compra.-", vbInformation + vbOKOnly
         FTXTAutoriza.SetFocus
         Exit Function
      End If
         
      If FSPRGrilla.MaxRows = 0 Then
          GLO_Respuesta = MsgBox("Falta Cargar los Renglones de la Orden de Compra.-", vbExclamation + vbOKOnly, "Atención")
          STABGrillas.Tab = Loc_SolapaCargarArticulo
          FTXTArticulo.SetFocus
          Exit Function
      End If
      
      If FSPRGrillaPorcentaje.MaxRows = 0 Then
          GLO_Respuesta = MsgBox("Falta Cargar Distribución de Gastos Por Uen.-", vbExclamation + vbOKOnly, "Atención")
          FPLIPorcentajeUen.SetFocus
          Exit Function
      Else
         If SumaPorcentaje <> 100 Then
            GLO_Respuesta = MsgBox("Falta Completar Porcentaje de Gastos Por Uen.-", vbExclamation + vbOKOnly, "Atención")
            FPLIPorcentajeUen.SetFocus
            Exit Function
         End If
      End If
   End If

   ValidarControles = True
End Function

Private Function GrabarRegistro() As Boolean

   GrabarRegistro = False
   Select Case Loc_Accion
      Case "A"
         If Not ProcesarAlta Then
            Exit Function
         End If
      Case "B"
         If Not ProcesarBaja Then
            Exit Function
         End If

      Case "M"
         If Not ProcesarModi Then
            Exit Function
         End If
   End Select

   If Not GrabarAuditoria_ADO(GLO_CodigoUen, GLO_Usuario, Format(FPLIId.Value, "00000000") & " " & FTXTRecibe.Text, Date, "ORDENPROVEEDORCABEZA", Loc_Accion, GLO_Subsistema) Then
      Exit Function
   End If

   GrabarRegistro = True
End Function

'ProcesarAlta
Private Function ProcesarAlta() As Boolean

   ProcesarAlta = False
   
   If Not InstanciarObjetoCabecera Then
      Exit Function
   Else
      If Not OrdenProveedorCabeza.Alta_OrdenProveedorCabeza Then
         Exit Function
      Else
         If GLO_ClaveId > 0 Then
            FPLIId.Value = GLO_ClaveId
            'renglones
            If Not DarDeAltaOrdenProveedorRenglones Then
               Exit Function
            Else
               'uen_porcentaje
               If Not DarDeAltaOrdenProveedorRenglonesCC Then
                  Exit Function
               Else
                  ProcesarAlta = True
               End If
            End If
         
         Else
            Exit Function
         End If
      End If
   End If
End Function

'ProcesarBaja
Private Function ProcesarBaja() As Boolean

   ProcesarBaja = False
      
   OrdenProveedorCabeza.Estado = "B"
      
   If Not OrdenProveedorCabeza.Modi_OrdenProveedorCabeza Then
        Exit Function
   Else
      ProcesarBaja = True
   End If

' la logica para borrar en las 3 tablas todos los registros de la orden de compra
'   If Not Baja_SolicitudCabeza Then
'      Exit Function
'   Else
'      If Not Baja_SolicitudRenglones Then
'         Exit Function
'      Else
'         If Not Baja_SolicitudRenglonesUen Then
'            Exit Function
'         Else
'            ProcesarBaja = True
'         End If
'      End If
'   End If
   
   Exit Function
End Function

'ProcesarModi
Private Function ProcesarModi() As Boolean

   ProcesarModi = False
   
   If Not InstanciarObjetoCabecera Then
      Exit Function
   Else
      If Not OrdenProveedorCabeza.Modi_OrdenProveedorCabeza Then
         Exit Function
      Else
         GLO_ClaveId = FPLIId.Value
         If GLO_ClaveId > 0 Then
   
            'renglones
            If SeModificoGrillaArticulo = True Then
               If Baja_SolicitudRenglones Then
                  If Not DarDeAltaOrdenProveedorRenglones Then
                     Exit Function
                  End If
               Else
                  Exit Function
               End If
            End If
            
            'uen_porcentaje
            If SeModificoGrillaUenPorcentaje = True Then
               If Baja_SolicitudRenglonesUen Then
                  If Not DarDeAltaOrdenProveedorRenglonesCC Then
                     Exit Function
                  End If
               Else
                  Exit Function
               End If
            End If
         
         Else
            Exit Function
         End If
      End If
   End If
   
   ProcesarModi = True
   
End Function

'se usa para la alta y la modificacion

Private Function DarDeAltaOrdenProveedorRenglones() As Boolean
   Dim i    As Integer
   
   DarDeAltaOrdenProveedorRenglones = False
   
   For i = 1 To FSPRGrilla.MaxRows
      If Not InstanciarObjetoRenglon(i) Then
         Exit Function
      End If
      If Not OrdenProveedorRenglon.Alta_OrdenProveedorRenglon Then
         Exit Function
      End If
   Next
   
   DarDeAltaOrdenProveedorRenglones = True
End Function

'se usa para la alta y la modificacion

Private Function DarDeAltaOrdenProveedorRenglonesCC() As Boolean
   Dim i    As Integer
   
   DarDeAltaOrdenProveedorRenglonesCC = False
   
   For i = 1 To FSPRGrillaPorcentaje.MaxRows
      If Not InstanciarObjetoCC(i) Then
         Exit Function
      End If
      If Not OrdenProveedorCC.Alta_OrdenProveedorCC Then
         Exit Function
      End If
   Next
   
   DarDeAltaOrdenProveedorRenglonesCC = True
End Function

'Metodo que se usa para dar de baja de forma fisica
'
'Private Function Baja_SolicitudCabeza() As Boolean
'
'   Baja_SolicitudCabeza = False
'
'   If OrdenProveedorCabeza.Baja_OrdenProveedorCabeza Then
'      Baja_SolicitudCabeza = True
'      Exit Function
'   End If
'End Function

' se usa para la modificacion la baja fisica
Private Function Baja_SolicitudRenglones() As Boolean

   Baja_SolicitudRenglones = False
   
   OrdenProveedorRenglon.Id = OrdenProveedorCabeza.Id
   
   If Not OrdenProveedorRenglon.Baja_OrdenProveedorRenglones Then
      Exit Function
   End If
   
   Baja_SolicitudRenglones = True
End Function

' se usa para la modificacion y la baja fisica
Private Function Baja_SolicitudRenglonesUen() As Boolean

   Baja_SolicitudRenglonesUen = False
   
   OrdenProveedorCC.Id = OrdenProveedorCabeza.Id
   
   If OrdenProveedorCC.Baja_OrdenProveedorCCs Then
      Baja_SolicitudRenglonesUen = True
      Exit Function
   End If

End Function

Private Function InstanciarObjetoCabecera() As Boolean
   InstanciarObjetoCabecera = False
      
   GLO_ClaveId = 0
   OrdenProveedorCabeza.Id = CLng(FPLIId.Value)
   OrdenProveedorCabeza.Fecha = FPDTFecha.Text
   OrdenProveedorCabeza.Uen = GLO_CodigoUen
   OrdenProveedorCabeza.CuentaUen = CInt(FPLIProveedorUen.Value)
   OrdenProveedorCabeza.CuentaCodigo = CLng(FPLIProveedorCodigo.Value)
   OrdenProveedorCabeza.RazonSocial = FTXTProveedorNombre.Text
   OrdenProveedorCabeza.Recibe = FTXTRecibe.Text
   OrdenProveedorCabeza.Autorizo = FTXTAutoriza.Text
   OrdenProveedorCabeza.Usuario = GLO_UsuarioNombre
   OrdenProveedorCabeza.Observaciones = FPMEObservaciones.Text
   OrdenProveedorCabeza.Estado = "A" 'atributo para saber si esta borrado o no
   OrdenProveedorCabeza.Impresa = "N" 'atributo para saber si esta impreso o no "S"
   
   InstanciarObjetoCabecera = True
End Function

'se setea los valores del objeto renglon
Private Function InstanciarObjetoRenglon(i As Integer) As Boolean
   InstanciarObjetoRenglon = False
   
   OrdenProveedorRenglon.Id = GLO_ClaveId
   OrdenProveedorRenglon.Renglon = i
   
   Retorno = FSPRGrilla.GetText(LOC_ColDescripcion, i, Contenido)
   If Retorno Then
      OrdenProveedorRenglon.Detalle = CStr(Contenido)
   End If
   
   Retorno = FSPRGrilla.GetText(Loc_ColCantidad, i, Contenido)
   If Retorno Then
      OrdenProveedorRenglon.Cantidad = CDbl(Contenido)
   End If
   
   Retorno = FSPRGrilla.GetText(Loc_ColUnidad, i, Contenido)
   If Retorno Then
      OrdenProveedorRenglon.Unidad = CStr(Contenido)
   End If

   InstanciarObjetoRenglon = True
End Function

'se setea los valores del objeto uen porcentaje
Private Function InstanciarObjetoCC(i As Integer) As Boolean
   InstanciarObjetoCC = False
   
   OrdenProveedorCC.Id = GLO_ClaveId
   
   Retorno = FSPRGrillaPorcentaje.GetText(Loc_ColUEN, i, Contenido)
   If Retorno Then
      OrdenProveedorCC.Uen = CInt(Contenido)
   End If
   
   Retorno = FSPRGrillaPorcentaje.GetText(Loc_ColPorcentaje, i, Contenido)
   If Retorno Then
      OrdenProveedorCC.Porcentaje = CDbl(Contenido)
   End If

   InstanciarObjetoCC = True
End Function

Private Sub SCMDProcesaRenglon_Click()
   If ValidarRenglon Then
      If CargarRenglonEnGrilla Then
         LimpiarRenglon
         FTXTArticulo.SetFocus
      End If
   End If
End Sub

Private Sub SCMDProcesaUen_Click()
   If ValidarRenglonUen Then
      If CargarRenglonUenEnGrilla Then
         SumaPorcentaje = SumaPorcentaje + FPDSPorcentajeGasto.Value
         LimpiarRenglonUen
         FPLIPorcentajeUen.SetFocus
      End If
   End If
End Sub

Private Sub SCMDSalir_Click()
   If Loc_Accion = "A" Or Loc_Accion = "M" Then
      GLO_Respuesta = MsgBox("Está seguro que desea descartar las modificaciones..?", vbQuestion + vbYesNo + vbDefaultButton2, "Atención")
      If GLO_Respuesta = vbYes Then
         Salida
      Else
         FSPRGrilla.Col = 1
         FSPRGrilla.row = FSPRGrilla.MaxRows
         If FSPRGrilla.Enabled Then
             FSPRGrilla.SetFocus
         End If
      End If
   Else
      Salida
   End If
End Sub

Private Sub Salida()
   Set OBJ_BusLogic = Nothing
   GLO_Accion = Loc_Accion
   Me.Hide
End Sub

'ValidarRenglon
Private Function ValidarRenglon() As Boolean
   ValidarRenglon = False

   If Len(Trim(FTXTArticulo.Text)) = 0 Then
      MsgBox "Debe Cargar El Artículo.-", vbInformation + vbOKOnly, "Atención"
      FTXTArticulo.SetFocus
      Exit Function
   End If

   If FPDSCantidad.Value <= 0 Then
      MsgBox "Falta Cargar la Cantidad del Artículo.-", vbInformation + vbOKOnly, "Atención"
      FPDSCantidad.SetFocus
      Exit Function
   End If

   If Len(Trim(FTXTUnidad.Text)) = 0 Then
      MsgBox "Debe Cargar a Las Unidad del Producto.-", vbInformation + vbOKOnly, "Atención"
      FTXTUnidad.SetFocus
      Exit Function
   End If

   ValidarRenglon = True
End Function

'ValidarRenglonUen
Private Function ValidarRenglonUen() As Boolean
   
   ValidarRenglonUen = False
   
   If Not Uen.Get_Uen(FPLIPorcentajeUen.Value, False) Then
      MsgBox "Debe Cargar El Num de Uen.-", vbInformation + vbOKOnly, "Atención"
      FPLIPorcentajeUen.SetFocus
      Exit Function
   End If

   If Len(Trim(FTXTPorcentajeNombre.Text)) = 0 Then
      MsgBox "Debe Apretar Enter para generar Nombre de Uen.-", vbInformation + vbOKOnly, "Atención"
      FPLIPorcentajeUen.SetFocus
      Exit Function
   End If

   If FPDSPorcentajeGasto.Value <= 0 Or FPDSPorcentajeGasto.Value > 100 Then
      MsgBox "El Porcentaje Ingresado es Invalido.-", vbInformation + vbOKOnly, "Atención"
      FPDSPorcentajeGasto.SetFocus
      Exit Function
   Else
      If (SumaPorcentaje + FPDSPorcentajeGasto.Value) > 100 Then
         MsgBox "El Porcentaje Ingresado Supera La Distribucion del 100%.-", vbInformation + vbOKOnly, "Atención"
         FPDSPorcentajeGasto.SetFocus
         Exit Function
      End If
   End If
   
   If Not NoHayUnaUenRepetida Then
      MsgBox "No puede Ingresar Dos Veces La Misma Uen.-", vbInformation + vbOKOnly, "Atención"
      FPLIPorcentajeUen.SetFocus
      Exit Function
   End If
     
   ValidarRenglonUen = True
End Function

' chequea si la uen que desea ingresar en la grilla de distribucion de gastos no esta repetida
Private Function NoHayUnaUenRepetida() As Boolean
   Dim i           As Integer
   
   NoHayUnaUenRepetida = False
   
   i = 1
   Do While (i <= FSPRGrillaPorcentaje.MaxRows)
      Retorno = FSPRGrillaPorcentaje.GetText(Loc_ColUEN, i, Contenido)
      If Retorno Then
         If (FPLIPorcentajeUen.Value = CInt(Contenido)) Then
            Exit Function
         End If
      Else
         MsgBox "Error al Comprobar Si Hay Uen Repetida, Intente de nuevo.-", vbInformation + vbOKOnly, "Atención"
         FPLIPorcentajeUen.SetFocus
         Exit Function
      End If
      i = i + 1
   Loop
   
   NoHayUnaUenRepetida = True

End Function

Private Function CargarRenglonEnGrilla() As Boolean
   CargarRenglonEnGrilla = False
   FSPRGrilla.MaxRows = FSPRGrilla.MaxRows + 1
   FSPRGrilla.SetText LOC_ColDescripcion, FSPRGrilla.MaxRows, FTXTArticulo.Text
   FSPRGrilla.SetText Loc_ColCantidad, FSPRGrilla.MaxRows, FPDSCantidad.Value
   FSPRGrilla.SetText Loc_ColUnidad, FSPRGrilla.MaxRows, FTXTUnidad.Text
   SeModificoGrillaArticulo = True
   CargarRenglonEnGrilla = True
End Function

Private Function CargarRenglonUenEnGrilla() As Boolean
   CargarRenglonUenEnGrilla = False
   FSPRGrillaPorcentaje.MaxRows = FSPRGrillaPorcentaje.MaxRows + 1
   FSPRGrillaPorcentaje.SetText Loc_ColUEN, FSPRGrillaPorcentaje.MaxRows, FPLIPorcentajeUen.Value
   FSPRGrillaPorcentaje.SetText Loc_ColNombre, FSPRGrillaPorcentaje.MaxRows, FTXTPorcentajeNombre.Text
   FSPRGrillaPorcentaje.SetText Loc_ColPorcentaje, FSPRGrillaPorcentaje.MaxRows, FPDSPorcentajeGasto.Value
   SeModificoGrillaUenPorcentaje = True
   CargarRenglonUenEnGrilla = True
End Function

Private Function BorrarRenglonEnGrilla() As Boolean
   BorrarRenglonEnGrilla = False
   If FSPRGrilla.MaxRows <> 0 Then
      FSPRGrilla.row = FSPRGrilla.ActiveRow
      'FSPRGrilla.Col = -1
      FSPRGrilla.Action = ActionDeleteRow
      FSPRGrilla.MaxRows = FSPRGrilla.MaxRows - 1
      SeModificoGrillaArticulo = True
      BorrarRenglonEnGrilla = True
   End If
End Function
Private Function BorrarUenEnGrilla() As Boolean
   BorrarUenEnGrilla = False
   If FSPRGrillaPorcentaje.MaxRows <> 0 Then
      
      'descuento el porcentaje de la fila que voy a borrar
      Retorno = FSPRGrillaPorcentaje.GetText(Loc_ColPorcentaje, FSPRGrillaPorcentaje.ActiveRow, Contenido)
      If Retorno Then
         SumaPorcentaje = SumaPorcentaje - CDbl(Contenido)
      End If
      
      FSPRGrillaPorcentaje.row = FSPRGrillaPorcentaje.ActiveRow
      FSPRGrillaPorcentaje.Col = -1
      FSPRGrillaPorcentaje.Action = ActionDeleteRow
      FSPRGrillaPorcentaje.MaxRows = FSPRGrillaPorcentaje.MaxRows - 1
      SeModificoGrillaUenPorcentaje = True
      BorrarUenEnGrilla = True
   End If
End Function

Private Sub EditarRenglon()
   If ValidarRenglon Then
      If EditarRenglonEnGrilla Then
         LimpiarRenglon
         FTXTArticulo.SetFocus
      End If
   End If
End Sub

Private Function EditarRenglonEnGrilla() As Boolean
   EditarRenglonEnGrilla = False
   If FSPRGrilla.MaxRows <> 0 Then
      FSPRGrilla.SetText LOC_ColDescripcion, FSPRGrilla.ActiveRow, FTXTArticulo.Text
      FSPRGrilla.SetText Loc_ColCantidad, FSPRGrilla.ActiveRow, FPDSCantidad.Value
      FSPRGrilla.SetText Loc_ColUnidad, FSPRGrilla.ActiveRow, FTXTUnidad.Text
      SeModificoGrillaArticulo = True
      EditarRenglonEnGrilla = True
   End If
End Function

Private Sub LimpiarRenglon()
   FTXTArticulo.Text = ""
   FPDSCantidad.Value = 0
   FPDSCantidad.MaxValue = 999999999
   FTXTUnidad.Text = ""
End Sub

Private Sub LimpiarRenglonUen()
   FPLIPorcentajeUen.Value = 0
   FTXTPorcentajeNombre.Text = ""
   FPDSPorcentajeGasto.Value = 0
End Sub

Private Function LeerRegistro() As Boolean
   LeerRegistro = False
   If OrdenProveedorCabeza.Get_OrdenProveedorCabeza(GLO_ClaveIdCodigo, False) Then
      If Uen.Get_Uen(OrdenProveedorCabeza.Uen, False) Then
         LeerRegistro = True
      End If
   End If
End Function

Private Function MoverCamposAControles() As Boolean

   MoverCamposAControles = False

   If LeerRegistro Then
      FPLIId.Value = OrdenProveedorCabeza.Id
      FPDTFecha.Text = OrdenProveedorCabeza.Fecha
      FPLIProveedorUen.Value = OrdenProveedorCabeza.CuentaUen
      FPLIProveedorCodigo.Value = OrdenProveedorCabeza.CuentaCodigo
      FTXTProveedorNombre.Text = OrdenProveedorCabeza.RazonSocial
      FTXTRecibe.Text = OrdenProveedorCabeza.Recibe
      FTXTAutoriza.Text = OrdenProveedorCabeza.Autorizo
      FTXTCreo.Text = OrdenProveedorCabeza.Usuario
      FPMEObservaciones.Text = OrdenProveedorCabeza.Observaciones
   
      Set RDO_OrdenProveedorRenglon = OrdenProveedorRenglon.Get_OrdenProveedorRenglones(CLng(OrdenProveedorCabeza.Id), False)
      If (RDO_OrdenProveedorRenglon.EOF) Then
         GLO_Mensaje = MsgBox("No se Puede Leer el Registro Seleccionado", vbOKOnly + vbCritical, "Advertencia")
      Else
         Do While Not RDO_OrdenProveedorRenglon.EOF
           
               FSPRGrilla.MaxRows = FSPRGrilla.MaxRows + 1
               
               FSPRGrilla.SetText LOC_ColDescripcion, FSPRGrilla.MaxRows, RDO_OrdenProveedorRenglon!OPR_DETALLE
               FSPRGrilla.SetText Loc_ColCantidad, FSPRGrilla.MaxRows, RDO_OrdenProveedorRenglon!OPR_CANTIDAD
               FSPRGrilla.SetText Loc_ColUnidad, FSPRGrilla.MaxRows, RDO_OrdenProveedorRenglon!OPR_UNIDAD
            
            RDO_OrdenProveedorRenglon.MoveNext
         Loop
         
         RDO_OrdenProveedorRenglon.Close
         Set RDO_OrdenProveedorRenglon = Nothing
         
         'GRILLA PORCENTAJE
         Set RDO_OrdenProveedorCC = OrdenProveedorCC.Get_OrdenProveedorCCs(CLng(OrdenProveedorCabeza.Id), False)
         If (RDO_OrdenProveedorCC.EOF) Then
            GLO_Mensaje = MsgBox("No se Puede Leer el Registro Seleccionado", vbOKOnly + vbCritical, "Advertencia")
         Else
            Do While Not RDO_OrdenProveedorCC.EOF
              
                  FSPRGrillaPorcentaje.MaxRows = FSPRGrillaPorcentaje.MaxRows + 1
                  
                  FSPRGrillaPorcentaje.SetText Loc_ColUEN, FSPRGrillaPorcentaje.MaxRows, RDO_OrdenProveedorCC!OCC_UEN
                  
                  If Uen.Get_Uen(RDO_OrdenProveedorCC!OCC_UEN, False) Then
                     FSPRGrillaPorcentaje.SetText Loc_ColNombre, FSPRGrillaPorcentaje.MaxRows, CStr(Uen.Nombre)
                  End If
                  
                  FSPRGrillaPorcentaje.SetText Loc_ColPorcentaje, FSPRGrillaPorcentaje.MaxRows, RDO_OrdenProveedorCC!OCC_PORCENTAJE

               RDO_OrdenProveedorCC.MoveNext
            Loop
            RDO_OrdenProveedorCC.Close
            Set RDO_OrdenProveedorCC = Nothing
            
            If Loc_Accion = "M" Then
               SumaPorcentaje = 100
            End If
            
            FSPRGrillaPorcentaje.Refresh
            FSPRGrilla.Refresh
            MoverCamposAControles = True
         End If
      End If
   Else
      MsgBox "No Se Pudo Leer el Registro.-", vbInformation + vbOKOnly, "Atención"
      Salida
   End If
   
End Function


'codigo de uen repetida
'   Dim i           As Integer
'   i = 1
'   Do While (UenRepetida = False) And (i <= FSPRGrillaPorcentaje.MaxRows)
'      Retorno = FSPRGrillaPorcentaje.GetText(Loc_ColUEN, i, Contenido)
'      If Retorno Then
'         If (FPLIPorcentajeUen.Value = CInt(Contenido)) Then
'            UenRepetida = True
'         End If
'      Else
'         MsgBox "Error al Comprobar Si Hay Uen Repetida, Intente de nuevo.-", vbInformation + vbOKOnly, "Atención"
'         FPLIPorcentajeUen.SetFocus
'         Exit Function
'      End If
'      i = i + 1
'   Loop
'
'   If UenRepetida = True Then
'      MsgBox "No puede Ingresar Dos Veces La Misma Uen.-", vbInformation + vbOKOnly, "Atención"
'      UenRepetida = False
'      FPLIPorcentajeUen.SetFocus
'      Exit Function
'   End If
