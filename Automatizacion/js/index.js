let lista_Atributos = document.querySelector("#atributos");

document.querySelector("#generar-atributo").addEventListener("click",agregarAtributo);
document.querySelector("#btn-codigo").addEventListener("click",generarCodigo);

function agregarAtributo(){
    let item = document.createElement("li");
    item.innerHTML= `<div>
                        <label>Nombre de Atributo: </label>
                        <input type="text" placeholder="Nombre Atributo">
                    </div>
                    <div>
                        <label>Tipo de Atributo: </label>
                        <input type="text" placeholder="Tipo Atributo">
                    </div>
                    <div>
                        <label>Nombre de col en BBDD: </label>
                        <input type="text" placeholder="Columna Base de Datos">
                    </div>
                    <div>
                        <label> Primary Key: </label>
                        <input type="checkbox">
                    </div>`;

                    lista_Atributos.appendChild(item);
}
                
function generarCodigo(){
    
    let arreglo_de_atributos= [];
    let nombre_de_clase = document.querySelector("#nombre-clase").value;
    let prefijo = document.querySelector("#prefijo-tabla").value;

    let items = lista_Atributos.children;
    for (const iterator of items) {
        let variable={};
        variable.nombre = iterator.firstElementChild.lastElementChild.value;
        variable.tipo = iterator.firstElementChild.nextElementSibling.lastElementChild.value;
        variable.columna = iterator.firstElementChild.nextElementSibling.nextElementSibling.lastElementChild.value;
        variable.clave = iterator.lastElementChild.lastElementChild.checked;
        arreglo_de_atributos.push(variable);
    }

    let parrafo_variables = document.querySelector("#variables");
    let parrafo_SetsLets = document.querySelector("#gets-sets");
    let parrafo_Get = document.querySelector("#get");

    parrafo_variables.innerHTML = "";
    parrafo_SetsLets.innerHTML = "";
    parrafo_Get.innerHTML = "";

    let texto1_get= `Public Function Get_${nombre_de_clase}(ByVal`;
    let texto2_get= `&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Get_${nombre_de_clase} = Get_${nombre_de_clase}_Local(Loc_ADO_Registro`;
    let texto3_get= `&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;If Get_${nombre_de_clase} = True Then<br>`;
    let primeraVez= true;

    let texto1_get_local= `Public Function Get_${nombre_de_clase}_Local(Registro_RS As ADODB.Recordset`;
    let texto2_get_local= ``;

    parrafo_variables.innerHTML+= `Option Explicit <br><br>
    Private OBJ_BusLogic  &nbsp;&nbsp;&nbsp;   As CLS_Buslogic <br>
    Private Loc_ADO_Registro  &nbsp;&nbsp;&nbsp;  As ADODB.Recordset <br>
    Private mvarRegistro_Rs  &nbsp;&nbsp;&nbsp;   As ADODB.Recordset <br><br>`;

    for (const atrib of arreglo_de_atributos) {
        parrafo_variables.innerHTML+=  `Private Loc_${atrib.nombre} &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; As ${atrib.tipo}<br>`;

        parrafo_SetsLets.innerHTML+=   `Public Property Let ${atrib.nombre}(ByVal vData As Variant)<br>
                                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Loc_${atrib.nombre} = vData <br>
                                        End Property <br><br>
                                        Public Property Get ${atrib.nombre}()<br>
                                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;${atrib.nombre} = Loc_${atrib.nombre}<br>
                                        End Property<br><br>`
        ;

        if (atrib.clave == true){
            if (primeraVez){
                primeraVez = false;
                texto1_get+= ` ${atrib.nombre} As ${atrib.tipo}`;
                texto2_get_local+=`${prefijo}_${atrib.columna} = " & ${atrib.nombre} & "`;
            }
            else{
                texto1_get+= `, ByVal ${atrib.nombre} As ${atrib.tipo}`;
                texto2_get_local+=`$ AND {prefijo}_${atrib.columna} = " & ${atrib.nombre} & "`;
            }
            texto2_get+= `, ${atrib.nombre}`;
            texto1_get_local+= `, ByVal ${atrib.nombre} As ${atrib.tipo}`;
        }

        texto3_get+= `&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Loc_${atrib.nombre} = Loc_ADO_Registro!${prefijo}_${atrib.columna}<br>`;

    }
    texto1_get+= `, ByVal Conectar As Boolean) As Boolean<br>`;

    texto2_get_local+=` "<br><br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;GLO_ADO_Comando.CommandType = adCmdText<br><br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Set Loc_RS = New ADODB.Recordset<br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Loc_RS.CursorLocation = adUseClient<br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Loc_RS.CursorType = adOpenStatic<br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Loc_RS.LockType = adLockOptimistic<br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Loc_RS.Open GLO_ADO_Comando<br><br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;If Not Loc_RS.EOF Then<br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Get_${nombre_de_clase}_Local = True<br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;End If<br><br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Set Registro_RS = Loc_RS<br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Set Loc_RS = Nothing<br><br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;If Conectar Then<br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Call Obj_buslogic.DesconectarDeSql_Aux<br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;End If<br><br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Exit Function<br><br>
                        ManejadorError:<br><br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Set Loc_RS = Nothing<br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Set GLO_ADO_Conexion_Aux = Nothing<br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;GLO_Mensaje = MsgBox("No se Puede Leer el Registro de ${nombre_de_clase} Seleccionado" & vbCr & "Error: " & CStr(Err.Number) & "-" & Err.Description, vbOKOnly + vbCritical, "Advertencia")<br><br>
                        End Function<br><br>`;

    texto1_get_local+=  `, Optional Conectar As Boolean) As Boolean<br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Dim Loc_RS         As ADODB.Recordset<br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Dim sErrDesc       As String<br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Dim lErrNo         As Long<br><br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;On Error GoTo ManejadorError<br><br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Get_${nombre_de_clase}_Local = False<br><br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;If Conectar Then<br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Call Obj_buslogic.ConectarASQL_Aux(True)<br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;End If<br><br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Set GLO_ADO_Comando = New ADODB.Command<br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;GLO_ADO_Comando.ActiveConnection = GLO_ADO_Conexion_Aux<br><br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;GLO_ADO_Comando.CommandText = "Select * From ${nombre_de_clase.toUpperCase()} (Nolock) Where `;

    parrafo_variables.innerHTML+=  `<br>`;
    parrafo_SetsLets.innerHTML+=    `Public Property Let Registro_RS(ByVal vData As ADODB.Recordset)<br>
                                     &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Set mvarRegistro_Rs = vData <br>
                                     End Property <br><br>
                                     Public Property Get Registro_RS() As ADODB.Recordset<br>
                                     &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Set Registro_RS = Loc_ADO_Registro<br>
                                     End Property<br><br>`;
    
    texto2_get+=    `, Conectar)<br>`;

    texto3_get+=    `&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;End If<br>
                    End Function<br><br>`;

    parrafo_Get.innerHTML += texto1_get;
    parrafo_Get.innerHTML += texto2_get;
    parrafo_Get.innerHTML += texto3_get;
    parrafo_Get.innerHTML += texto1_get_local;
    parrafo_Get.innerHTML += texto2_get_local;
}