Option Explicit

Private OBJ_BusLogic               As CLS_Buslogic
Private Loc_ADO_Registro           As ADODB.Recordset
Private mvarRegistro_Rs            As ADODB.Recordset

Private Loc_Id                     As Long
Private Loc_Uen                    As Integer
Private Loc_CuitFirmante           As String


Public Property Let Id(ByVal vData As Variant)
     Loc_Id = vData
End Property

Public Property Get Id()
     Id = Loc_Id
End Property

Public Property Let Uen(ByVal vData As Variant)
     Loc_Uen = vData
End Property

Public Property Get Uen()
     Uen = Loc_Uen
End Property

Public Property Let CuitFirmante(ByVal vData As Variant)
     Loc_CuitFirmante = vData
End Property

Public Property Get CuitFirmante()
     CuitFirmante = Loc_CuitFirmante
End Property

Public Property Let Registro_RS(ByVal vData As ADODB.Recordset)
     Set mvarRegistro_Rs = vData
End Property

Public Property Get Registro_RS() As ADODB.Recordset
     Set Registro_RS = Loc_ADO_Registro
End Property


Public Function Get_MaestroCuentasCD(ByVal Id As Long, ByVal Conectar As Boolean) As Boolean
     Get_MaestroCuentasCD = Get_MaestroCuentasCD_Local(Loc_ADO_Registro, Id, Conectar)
     If Get_MaestroCuentasCD = True Then
          Loc_Id = Loc_ADO_Registro!CCD_ID
          Loc_Uen = Loc_ADO_Registro!CCD_UEN
          Loc_CuitFirmante = Loc_ADO_Registro!CCD_CUITFIRMANTE
     End If
End Function

Public Function Get_MaestroCuentasCD_Local(Registro_RS As ADODB.Recordset, ByVal Id As Long, Optional Conectar As Boolean) As Boolean
     Dim Loc_RS As ADODB.Recordset
     Dim sErrDesc As String
     Dim lErrNo As Long

     On Error GoTo ManejadorError

     Get_MaestroCuentasCD_Local = False

     If Conectar Then
        Call Obj_buslogic.ConectarASQL_Aux(True)
     End If

     Set GLO_ADO_Comando = New ADODB.Command
     GLO_ADO_Comando.ActiveConnection = GLO_ADO_Conexion_Aux

     GLO_ADO_Comando.CommandText = "Select * From MAESTROCUENTASCD (Nolock) Where CCD_ID = " & Id & " "

     GLO_ADO_Comando.CommandType = adCmdText

     Set Loc_RS = New ADODB.Recordset
     Loc_RS.CursorLocation = adUseClient
     Loc_RS.CursorType = adOpenStatic
     Loc_RS.LockType = adLockOptimistic
     Loc_RS.Open GLO_ADO_Comando

     If Not Loc_RS.EOF Then
     Get_MaestroCuentasCD_Local = True
     End If

     Set Registro_RS = Loc_RS
     Set Loc_RS = Nothing

     If Conectar Then
     Call Obj_buslogic.DesconectarDeSql_Aux
     End If

     Exit Function

ManejadorError:

     Set Loc_RS = Nothing
     Set GLO_ADO_Conexion_Aux = Nothing
     GLO_Mensaje = MsgBox("No se Puede Leer el Registro de MaestroCuentasCD Seleccionado" & vbCr & "Error: " & CStr(Err.Number) & "-" & Err.Description, vbOKOnly + vbCritical, "Advertencia")

End Function
