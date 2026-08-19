const SYHME = Object.freeze({
  HOJAS: Object.freeze({
    CONFIGURACION: 'CONFIGURACION',
    PUESTOS: 'PUESTOS',
    EMPLEADOS: 'EMPLEADOS',
    CLIENTES: 'CLIENTES',
    EVENTOS: 'EVENTOS',
    COTIZACIONES: 'COTIZACIONES',
    ASIGNACIONES: 'ASIGNACIONES',
    TURNOS: 'TURNOS',
    NOMINAS: 'NOMINAS',
    DETALLE_NOMINA: 'DETALLE_NOMINA',
    PRESTAMOS: 'PRESTAMOS',
    MOVIMIENTOS_PRESTAMO: 'MOVIMIENTOS_PRESTAMO',
    PAGOS: 'PAGOS',
    CONTROL_FINANCIERO: 'CONTROL_FINANCIERO',
    VISTA_NOMINA: 'VISTA_NOMINA',
    BITACORA: 'BITACORA',
  }),
  ESTATUS_ACTIVO: 'ACTIVO',
  ID_CARPETA_RAIZ_EVENTOS: '1r2vUndE43ync3fmBl9_cayv-QZ1a4YtP',
  PUESTO_DIRECCION: 'DIRECCIÓN',
  FRACCIONES_TURNO: Object.freeze({ '': 0, '1/2': 0.5, '12': 1, '12 1/2': 1.5, '24': 2 }),
  PALETA_DIAS: Object.freeze([
    { fondo: '#fff2cc', activo: '#bf9000' },
    { fondo: '#fce5cd', activo: '#b45f06' },
    { fondo: '#f4cccc', activo: '#cc0000' },
    { fondo: '#d9ead3', activo: '#38761d' },
    { fondo: '#c9daf8', activo: '#1155cc' },
    { fondo: '#d9d2e9', activo: '#674ea7' },
  ]),
});

function onOpen() {
  try { actualizarEstatusEventos_(); } catch (e) { /* silencioso: no debe romper la apertura del archivo */ }

  const ui = SpreadsheetApp.getUi();

  const menuAbrir = ui.createMenu('Abrir')
    .addItem('Catálogo de empleados', 'abrirCatalogoEmpleados')
    .addItem('Catálogo de puestos', 'abrirCatalogoPuestos')
    .addItem('Catálogo de clientes', 'abrirCatalogoClientes')
    .addItem('Catálogo de eventos', 'abrirCatalogoEventos')
    .addItem('Catálogo de cotizaciones', 'abrirCatalogoCotizaciones')
    .addItem('Catálogo de asignaciones', 'abrirCatalogoAsignaciones')
    .addItem('Catálogo de turnos', 'abrirCatalogoTurnos')
    .addItem('Catálogo de nóminas', 'abrirCatalogoNominas')
    .addItem('Catálogo de detalle de nómina', 'abrirCatalogoDetalleNomina')
    .addItem('Catálogo de préstamos', 'abrirCatalogoPrestamos')
    .addItem('Catálogo de movimientos de préstamo', 'abrirCatalogoMovimientosPrestamo')
    .addItem('Catálogo de pagos', 'abrirCatalogoPagos')
    .addItem('Catálogo de control financiero', 'abrirCatalogoControlFinanciero');

  ui.createMenu('Syhme')
    .addItem('Dar de alta empleado', 'mostrarFormularioAltaEmpleado')
    .addSeparator()
    .addItem('Administrar puestos y tarifas', 'mostrarFormularioPuestos')
    .addSeparator()
    .addItem('Administrar clientes', 'mostrarFormularioClientes')
    .addItem('Administrar eventos', 'mostrarFormularioEventos')
    .addSeparator()
    .addItem('Administrar cotizaciones', 'mostrarFormularioCotizaciones')
    .addSeparator()
    .addItem('Administrar asignaciones', 'mostrarFormularioAsignaciones')
    .addItem('Administrar turnos', 'mostrarFormularioTurnos')
    .addSeparator()
    .addItem('Dar de alta préstamo', 'mostrarFormularioPrestamos')
    .addSeparator()
    .addItem('Generar / regenerar nómina', 'mostrarFormularioNominas')
    .addSeparator()
    .addItem('Administrar pagos de cliente', 'mostrarFormularioPagos')
    .addItem('Generar control financiero', 'mostrarFormularioControlFinanciero')
    .addSeparator()
    .addItem('Ver resumen de evento (Vista Nómina)', 'mostrarFormularioVistaNomina')
    .addSeparator()
    .addSubMenu(menuAbrir)
    .addSeparator()
    .addItem('Actualizar estatus de eventos ahora', 'actualizarEstatusEventos_')
    .addItem('Activar actualización diaria automática', 'instalarActualizacionDiariaEventos')
    .addItem('Verificar estructura', 'verificarEstructura')
    .addToUi();
}

function onInstall() {
  onOpen();
}

function mostrarFormularioAltaEmpleado() {
  const puestos = obtenerPuestosActivos_();
  if (!puestos.length) {
    SpreadsheetApp.getUi().alert(
      'No hay puestos activos. Registra al menos un puesto antes de dar de alta empleados.',
    );
    return;
  }

  const puestosJson = JSON.stringify(puestos).replace(/</g, '\\u003c');
  const html = HtmlService.createHtmlOutput(`
    <!doctype html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: Arial, sans-serif; margin: 0; padding: 22px; color: #263238; }
          h2 { margin: 0 0 18px; color: #2d516a; font-size: 20px; }
          label { display: block; margin: 14px 0 6px; font-weight: 700; }
          input, select { box-sizing: border-box; width: 100%; padding: 10px; border: 1px solid #b0bec5; border-radius: 6px; font-size: 14px; }
          input:focus, select:focus { border-color: #2d516a; outline: none; box-shadow: 0 0 0 2px #dbe7ee; }
          .acciones { display: flex; justify-content: flex-end; gap: 8px; margin-top: 22px; }
          button { border: 0; border-radius: 6px; padding: 10px 16px; cursor: pointer; font-weight: 700; }
          .primario { background: #2d516a; color: white; }
          .secundario { background: #eceff1; color: #37474f; }
          button:disabled { cursor: wait; opacity: .65; }
          #mensaje { min-height: 20px; margin-top: 14px; font-size: 13px; }
          .ok { color: #1b5e20; }
          .error { color: #b71c1c; }
        </style>
      </head>
      <body>
        <h2>Alta de empleado</h2>
        <form id="formulario">
          <label for="nombre">Nombre completo</label>
          <input id="nombre" name="nombre" maxlength="120" autocomplete="off" required>

          <label for="idPuesto">Puesto</label>
          <select id="idPuesto" name="idPuesto" required></select>

          <div id="mensaje" aria-live="polite"></div>
          <div class="acciones">
            <button type="button" class="secundario" onclick="google.script.host.close()">Cancelar</button>
            <button id="guardar" type="submit" class="primario">Guardar empleado</button>
          </div>
        </form>
        <script>
          const puestos = ${puestosJson};
          const selector = document.getElementById('idPuesto');
          puestos.forEach(({ id, nombre }) => {
            const opcion = document.createElement('option');
            opcion.value = id;
            opcion.textContent = nombre;
            selector.appendChild(opcion);
          });

          document.getElementById('formulario').addEventListener('submit', (evento) => {
            evento.preventDefault();
            const boton = document.getElementById('guardar');
            const mensaje = document.getElementById('mensaje');
            boton.disabled = true;
            mensaje.className = '';
            mensaje.textContent = 'Guardando...';

            google.script.run
              .withSuccessHandler((resultado) => {
                boton.disabled = false;
                mensaje.className = resultado.ok ? 'ok' : 'error';
                mensaje.textContent = resultado.mensaje;
                if (resultado.ok) {
                  document.getElementById('nombre').value = '';
                  document.getElementById('nombre').focus();
                }
              })
              .withFailureHandler((error) => {
                boton.disabled = false;
                mensaje.className = 'error';
                mensaje.textContent = error && error.message ? error.message : 'No fue posible guardar el empleado.';
              })
              .registrarEmpleadoDesdeFormulario({
                nombre: document.getElementById('nombre').value,
                idPuesto: selector.value,
              });
          });
        </script>
      </body>
    </html>
  `).setWidth(430).setHeight(380);

  SpreadsheetApp.getUi().showModalDialog(html, 'Syhme');
}

function registrarEmpleadoDesdeFormulario(datos) {
  const nombre = normalizarNombre_(datos && datos.nombre);
  const idPuesto = String((datos && datos.idPuesto) || '').trim();

  if (!nombre) return { ok: false, mensaje: 'Escribe el nombre del empleado.' };
  if (!idPuesto) return { ok: false, mensaje: 'Selecciona un puesto.' };

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) {
    return { ok: false, mensaje: 'El sistema está ocupado. Intenta nuevamente en unos segundos.' };
  }

  try {
    const libro = SpreadsheetApp.getActiveSpreadsheet();
    const hojaEmpleados = obtenerHojaObligatoria_(libro, SYHME.HOJAS.EMPLEADOS);
    const puesto = obtenerPuestoActivoPorId_(idPuesto);
    if (!puesto) return { ok: false, mensaje: 'El puesto seleccionado ya no está activo.' };

    const duplicado = buscarEmpleadoPorNombre_(hojaEmpleados, nombre);
    if (duplicado) {
      return {
        ok: false,
        mensaje: `Ya existe ${duplicado.nombre} con el ID ${duplicado.id}.`,
      };
    }

    const idEmpleado = generarId_('EMP', hojaEmpleados);
    const ahora = new Date();
    hojaEmpleados.appendRow([
      idEmpleado,
      nombre,
      puesto.id,
      puesto.nombre,
      SYHME.ESTATUS_ACTIVO,
      ahora,
      ahora,
    ]);

    registrarBitacora_({
      modulo: 'EMPLEADOS',
      accion: 'ALTA',
      idRegistro: idEmpleado,
      valorNuevo: JSON.stringify({ nombre, puesto: puesto.nombre }),
      detalle: 'Empleado registrado desde el formulario de alta.',
    });

    return { ok: true, mensaje: `Empleado guardado correctamente con ID ${idEmpleado}.`, idEmpleado };
  } finally {
    lock.releaseLock();
  }
}

function abrirCatalogoEmpleados() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  libro.setActiveSheet(obtenerHojaObligatoria_(libro, SYHME.HOJAS.EMPLEADOS));
}

function obtenerEmpleados_() {
  const hoja = obtenerHojaObligatoria_(SpreadsheetApp.getActiveSpreadsheet(), SYHME.HOJAS.EMPLEADOS);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return [];
  return hoja.getRange(2, 1, ultimaFila - 1, 7).getValues()
    .filter((f) => f[0] && f[1])
    .map((f, i) => ({
      fila: i + 2, id: String(f[0]), nombre: String(f[1]), idPuesto: String(f[2]),
      puesto: String(f[3]), estatus: String(f[4]), fechaAlta: f[5],
    }));
}

function obtenerEmpleadosActivos_() {
  return obtenerEmpleados_().filter((e) => e.estatus === SYHME.ESTATUS_ACTIVO);
}

function mostrarFormularioPuestos() {
  const puestosJson = JSON.stringify(obtenerPuestos_()).replace(/</g, '\\u003c');
  const html = HtmlService.createHtmlOutput(`
    <!doctype html><html><head><base target="_top"><style>
      body{font-family:Arial,sans-serif;margin:0;padding:22px;color:#263238}h2{margin:0 0 12px;color:#2d516a}
      p{color:#546e7a;font-size:12px}label{display:block;margin:11px 0 5px;font-weight:700;font-size:13px}
      input,select,textarea{box-sizing:border-box;width:100%;padding:9px;border:1px solid #b0bec5;border-radius:6px;font-size:14px}
      textarea{min-height:58px;resize:vertical}.grid{display:grid;grid-template-columns:1fr 1fr;gap:0 12px}.full{grid-column:1/-1}
      .acciones{display:flex;justify-content:flex-end;gap:8px;margin-top:18px}button{border:0;border-radius:6px;padding:10px 16px;cursor:pointer;font-weight:700}
      .primario{background:#2d516a;color:white}.secundario{background:#eceff1;color:#37474f}button:disabled{opacity:.65}
      #mensaje{min-height:18px;margin-top:11px;font-size:13px}.ok{color:#1b5e20}.error{color:#b71c1c}
    </style></head><body>
      <h2>Puestos y tarifas</h2><p>Crea un puesto o selecciona uno para actualizarlo. Las nóminas generadas conservarán su tarifa histórica.</p>
      <form id="formulario"><label for="idPuesto">Registro</label><select id="idPuesto"><option value="">+ NUEVO PUESTO</option></select>
      <div class="grid"><div class="full"><label for="nombre">Nombre del puesto</label><input id="nombre" maxlength="80" required></div>
      <div><label for="tarifa">Tarifa por turno</label><input id="tarifa" type="number" min="0.01" step="0.01" required></div>
      <div><label for="estatus">Estatus</label><select id="estatus"><option>ACTIVO</option><option>INACTIVO</option></select></div>
      <div><label for="orden">Orden</label><input id="orden" type="number" min="1" step="1" required></div>
      <div class="full"><label for="observaciones">Observaciones</label><textarea id="observaciones" maxlength="300"></textarea></div></div>
      <div id="mensaje"></div><div class="acciones"><button type="button" class="secundario" onclick="google.script.host.close()">Cancelar</button>
      <button id="guardar" type="submit" class="primario">Guardar puesto</button></div></form>
      <script>
        const puestos=${puestosJson},selector=document.getElementById('idPuesto');
        puestos.forEach(p=>{const o=document.createElement('option');o.value=p.id;o.textContent=p.nombre;selector.appendChild(o)});
        function cargar(){const p=puestos.find(x=>x.id===selector.value);nombre.value=p?p.nombre:'';tarifa.value=p?p.tarifa:'';estatus.value=p?p.estatus:'ACTIVO';orden.value=p?p.orden:puestos.length+1;observaciones.value=p?p.observaciones:'';mensaje.textContent=''}
        selector.addEventListener('change',cargar);cargar();
        formulario.addEventListener('submit',e=>{e.preventDefault();guardar.disabled=true;mensaje.className='';mensaje.textContent='Guardando...';
          google.script.run.withSuccessHandler(r=>{guardar.disabled=false;mensaje.className=r.ok?'ok':'error';mensaje.textContent=r.mensaje;if(r.ok)setTimeout(()=>google.script.host.close(),900)})
          .withFailureHandler(e=>{guardar.disabled=false;mensaje.className='error';mensaje.textContent=e&&e.message?e.message:'No fue posible guardar el puesto.'})
          .guardarPuestoDesdeFormulario({idPuesto:selector.value,nombre:nombre.value,tarifa:tarifa.value,estatus:estatus.value,orden:orden.value,observaciones:observaciones.value})});
      </script></body></html>
  `).setWidth(520).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Syhme');
}

function guardarPuestoDesdeFormulario(datos) {
  const idPuesto=String((datos&&datos.idPuesto)||'').trim();
  const nombre=normalizarNombre_(datos&&datos.nombre);
  const tarifa=Number(datos&&datos.tarifa);
  const estatus=String((datos&&datos.estatus)||'').trim().toUpperCase();
  const orden=Number(datos&&datos.orden);
  const observaciones=String((datos&&datos.observaciones)||'').trim().slice(0,300);
  if(!nombre)return{ok:false,mensaje:'Escribe el nombre del puesto.'};
  if(!Number.isFinite(tarifa)||tarifa<=0)return{ok:false,mensaje:'La tarifa debe ser mayor que cero.'};
  if(!['ACTIVO','INACTIVO'].includes(estatus))return{ok:false,mensaje:'Selecciona un estatus válido.'};
  if(!Number.isInteger(orden)||orden<1)return{ok:false,mensaje:'El orden debe ser un entero mayor que cero.'};
  const lock=LockService.getDocumentLock();
  if(!lock.tryLock(15000))return{ok:false,mensaje:'El sistema está ocupado. Intenta nuevamente.'};
  try{
    const libro=SpreadsheetApp.getActiveSpreadsheet(),hoja=obtenerHojaObligatoria_(libro,SYHME.HOJAS.PUESTOS),puestos=obtenerPuestos_();
    const duplicado=puestos.find(p=>claveTexto_(p.nombre)===claveTexto_(nombre)&&p.id!==idPuesto);
    if(duplicado)return{ok:false,mensaje:`Ya existe el puesto ${duplicado.nombre}.`};
    if(idPuesto){
      const actual=puestos.find(p=>p.id===idPuesto);if(!actual)return{ok:false,mensaje:'El puesto seleccionado ya no existe.'};
      if(estatus==='INACTIVO'&&actual.estatus==='ACTIVO'&&contarEmpleadosActivosPorPuesto_(idPuesto)>0)return{ok:false,mensaje:'No se puede desactivar: hay empleados activos asignados.'};
      hoja.getRange(actual.fila,2,1,7).setValues([[nombre,tarifa,estatus,actual.fechaAlta||new Date(),new Date(),observaciones,orden]]);
      registrarBitacora_({modulo:'PUESTOS',accion:'MODIFICACION',idRegistro:idPuesto,valorAnterior:JSON.stringify({nombre:actual.nombre,tarifa:actual.tarifa,estatus:actual.estatus,orden:actual.orden}),valorNuevo:JSON.stringify({nombre,tarifa,estatus,orden}),detalle:'Puesto y tarifa actualizados desde el formulario.'});
      return{ok:true,mensaje:`Puesto ${idPuesto} actualizado correctamente.`};
    }
    const nuevoId=generarId_('PUE',hoja),ahora=new Date();
    hoja.appendRow([nuevoId,nombre,tarifa,estatus,ahora,ahora,observaciones,orden]);
    registrarBitacora_({modulo:'PUESTOS',accion:'ALTA',idRegistro:nuevoId,valorNuevo:JSON.stringify({nombre,tarifa,estatus,orden}),detalle:'Puesto registrado desde el formulario.'});
    return{ok:true,mensaje:`Puesto guardado correctamente con ID ${nuevoId}.`,idPuesto:nuevoId};
  }finally{lock.releaseLock()}
}

function abrirCatalogoPuestos(){const libro=SpreadsheetApp.getActiveSpreadsheet();libro.setActiveSheet(obtenerHojaObligatoria_(libro,SYHME.HOJAS.PUESTOS))}

function obtenerPuestos_(){
  const hoja=obtenerHojaObligatoria_(SpreadsheetApp.getActiveSpreadsheet(),SYHME.HOJAS.PUESTOS),ultimaFila=hoja.getLastRow();if(ultimaFila<2)return[];
  return hoja.getRange(2,1,ultimaFila-1,8).getValues().filter(f=>f[0]&&f[1]).map((f,i)=>({fila:i+2,id:String(f[0]),nombre:String(f[1]),tarifa:Number(f[2])||0,estatus:String(f[3]),fechaAlta:f[4],observaciones:String(f[6]||''),orden:Number(f[7])||i+1})).sort((a,b)=>a.orden-b.orden||a.nombre.localeCompare(b.nombre,'es'));
}

function contarEmpleadosActivosPorPuesto_(idPuesto){
  const hoja=obtenerHojaObligatoria_(SpreadsheetApp.getActiveSpreadsheet(),SYHME.HOJAS.EMPLEADOS),ultimaFila=hoja.getLastRow();if(ultimaFila<2)return 0;
  return hoja.getRange(2,3,ultimaFila-1,3).getDisplayValues().filter(f=>f[0]===idPuesto&&f[2]===SYHME.ESTATUS_ACTIVO).length;
}

function mostrarFormularioClientes() {
  const clientesJson=JSON.stringify(obtenerClientes_()).replace(/</g,'\\u003c');
  const html=HtmlService.createHtmlOutput(`<!doctype html><html><head><base target="_top"><style>${estilosFormulario_()}</style></head><body>
  <h2>Clientes</h2><p>Crea un cliente o selecciona uno para actualizarlo.</p><form id="formulario">
  <label for="idCliente">Registro</label><select id="idCliente"><option value="">+ NUEVO CLIENTE</option></select>
  <label for="nombre">Nombre del cliente</label><input id="nombre" maxlength="120" required>
  <label for="correo">Correo de facturación</label><input id="correo" type="email" maxlength="150">
  <label for="estatus">Estatus</label><select id="estatus"><option>ACTIVO</option><option>INACTIVO</option></select>
  <label for="observaciones">Observaciones</label><textarea id="observaciones" maxlength="300"></textarea>
  <div id="mensaje"></div><div class="acciones"><button type="button" class="secundario" onclick="google.script.host.close()">Cancelar</button><button id="guardar" class="primario">Guardar cliente</button></div></form>
  <script>const datos=${clientesJson},q=id=>document.getElementById(id),sel=q('idCliente');datos.forEach(x=>{const o=document.createElement('option');o.value=x.id;o.textContent=x.nombre;sel.appendChild(o)});
  function cargar(){const x=datos.find(v=>v.id===sel.value);q('nombre').value=x?x.nombre:'';q('correo').value=x?x.correo:'';q('estatus').value=x?x.estatus:'ACTIVO';q('observaciones').value=x?x.observaciones:'';q('mensaje').textContent=''}sel.onchange=cargar;cargar();
  q('formulario').onsubmit=e=>{e.preventDefault();q('guardar').disabled=true;q('mensaje').textContent='Guardando...';google.script.run.withSuccessHandler(r=>{q('guardar').disabled=false;q('mensaje').className=r.ok?'ok':'error';q('mensaje').textContent=r.mensaje;if(r.ok)setTimeout(()=>google.script.host.close(),900)}).withFailureHandler(e=>{q('guardar').disabled=false;q('mensaje').className='error';q('mensaje').textContent=e.message||'No fue posible guardar.'}).guardarClienteDesdeFormulario({idCliente:sel.value,nombre:q('nombre').value,correo:q('correo').value,estatus:q('estatus').value,observaciones:q('observaciones').value})};</script></body></html>`).setWidth(500).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html,'Syhme');
}

function guardarClienteDesdeFormulario(datos) {
  const id=String((datos&&datos.idCliente)||'').trim(),nombre=normalizarNombre_(datos&&datos.nombre),correo=String((datos&&datos.correo)||'').trim().toLowerCase(),estatus=String((datos&&datos.estatus)||'').trim().toUpperCase(),observaciones=String((datos&&datos.observaciones)||'').trim().slice(0,300);
  if(!nombre)return{ok:false,mensaje:'Escribe el nombre del cliente.'};
  if(correo&&!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(correo))return{ok:false,mensaje:'El correo no tiene un formato válido.'};
  if(!['ACTIVO','INACTIVO'].includes(estatus))return{ok:false,mensaje:'Selecciona un estatus válido.'};
  const lock=LockService.getDocumentLock();if(!lock.tryLock(15000))return{ok:false,mensaje:'El sistema está ocupado. Intenta nuevamente.'};
  try{const hoja=obtenerHojaObligatoria_(SpreadsheetApp.getActiveSpreadsheet(),SYHME.HOJAS.CLIENTES),clientes=obtenerClientes_(),duplicado=clientes.find(x=>claveTexto_(x.nombre)===claveTexto_(nombre)&&x.id!==id);if(duplicado)return{ok:false,mensaje:`Ya existe el cliente ${duplicado.nombre}.`};
    if(id){const actual=clientes.find(x=>x.id===id);if(!actual)return{ok:false,mensaje:'El cliente seleccionado ya no existe.'};if(estatus==='INACTIVO'&&actual.estatus==='ACTIVO'&&contarEventosAbiertosPorCliente_(id)>0)return{ok:false,mensaje:'No se puede desactivar: el cliente tiene eventos abiertos.'};
      hoja.getRange(actual.fila,2,1,6).setValues([[nombre,correo,estatus,actual.fechaAlta||new Date(),new Date(),observaciones]]);registrarBitacora_({modulo:'CLIENTES',accion:'MODIFICACION',idRegistro:id,valorAnterior:JSON.stringify({nombre:actual.nombre,correo:actual.correo,estatus:actual.estatus}),valorNuevo:JSON.stringify({nombre,correo,estatus}),detalle:'Cliente actualizado desde el formulario.'});return{ok:true,mensaje:`Cliente ${id} actualizado correctamente.`}}
    const nuevoId=generarId_('CLI',hoja),ahora=new Date();hoja.appendRow([nuevoId,nombre,correo,estatus,ahora,ahora,observaciones]);registrarBitacora_({modulo:'CLIENTES',accion:'ALTA',idRegistro:nuevoId,valorNuevo:JSON.stringify({nombre,correo,estatus}),detalle:'Cliente registrado desde el formulario.'});return{ok:true,mensaje:`Cliente guardado con ID ${nuevoId}.`};
  }finally{lock.releaseLock()}
}

/* ======================= EVENTOS (estatus dinámico + carpeta Drive automática + auto-asignación de dirección) ======================= */

function calcularEstatusEvento_(inicio, fin, hoy) {
  hoy = hoy || new Date();
  const aDia = (f) => new Date(f.getFullYear(), f.getMonth(), f.getDate());
  const inicioDia = aDia(inicio);
  const finDia = aDia(fin);
  const hoyDia = aDia(hoy);

  if (hoyDia < inicioDia) return 'CONFIRMADO';
  if (hoyDia > finDia) return 'FINALIZADO';
  return 'EN CURSO';
}

function obtenerCarpetaRaizEventos_() {
  return DriveApp.getFolderById(SYHME.ID_CARPETA_RAIZ_EVENTOS);
}

function normalizarNombreCarpeta_(valor) {
  return String(valor || '')
    .trim()
    .replace(/[\\/]/g, ' - ')
    .replace(/\s+/g, ' ')
    .substring(0, 200);
}

function crearOEncontrarCarpetaEvento_(idEvento, nombreEvento) {
  const raiz = obtenerCarpetaRaizEventos_();
  const nombreCarpeta = normalizarNombreCarpeta_(`${idEvento} - ${nombreEvento}`);
  const existentes = raiz.getFoldersByName(nombreCarpeta);
  const carpeta = existentes.hasNext() ? existentes.next() : raiz.createFolder(nombreCarpeta);
  return carpeta.getUrl();
}

function formulaCarpeta_(url) {
  return `=HYPERLINK("${url}","Carpeta")`;
}

function mostrarFormularioEventos() {
  const eventosJson=JSON.stringify(obtenerEventos_()).replace(/</g,'\\u003c'),clientesJson=JSON.stringify(obtenerClientes_().filter(x=>x.estatus==='ACTIVO')).replace(/</g,'\\u003c');
  if(clientesJson==='[]'){SpreadsheetApp.getUi().alert('Registra al menos un cliente activo antes de crear eventos.');return}
  const html=HtmlService.createHtmlOutput(`<!doctype html><html><head><base target="_top"><style>${estilosFormulario_()}
    .estatus-actual{background:#eceff1;border-radius:6px;padding:9px;font-weight:700;color:#2d516a}
    .check-cancelar{display:flex;align-items:center;gap:8px;font-weight:400;margin-top:10px}
    .check-cancelar input{width:auto}
  </style></head><body><h2>Eventos</h2><p>El estatus se calcula solo según las fechas. Al crear el evento se genera automáticamente su carpeta en Drive y se asigna a los directores activos.</p><form id="formulario">
  <label for="idEvento">Registro</label><select id="idEvento"><option value="">+ NUEVO EVENTO</option></select><label for="nombre">Nombre del evento</label><input id="nombre" maxlength="120" required>
  <label for="idCliente">Cliente</label><select id="idCliente" required></select><div class="grid"><div><label for="inicio">Fecha de inicio</label><input id="inicio" type="date" required></div><div><label for="fin">Fecha de fin</label><input id="fin" type="date" required></div></div>
  <div id="bloqueEstatus" style="display:none">
    <label>Estatus actual</label>
    <div id="estatusActual" class="estatus-actual">—</div>
    <label class="check-cancelar"><input type="checkbox" id="cancelar"> Cancelar este evento</label>
  </div>
  <label for="observaciones">Observaciones</label><textarea id="observaciones" maxlength="300"></textarea><div id="mensaje"></div><div class="acciones"><button type="button" class="secundario" onclick="google.script.host.close()">Cancelar</button><button id="guardar" class="primario">Guardar evento</button></div></form>
  <script>const eventos=${eventosJson},clientes=${clientesJson},q=id=>document.getElementById(id),sel=q('idEvento'),sc=q('idCliente');clientes.forEach(x=>{const o=document.createElement('option');o.value=x.id;o.textContent=x.nombre;sc.appendChild(o)});eventos.forEach(x=>{const o=document.createElement('option');o.value=x.id;o.textContent=x.nombre+' — '+x.inicio;sel.appendChild(o)});
  function cargar(){const x=eventos.find(v=>v.id===sel.value),nuevo=!x;q('nombre').value=x?x.nombre:'';if(x&&![...sc.options].some(o=>o.value===x.idCliente)){const o=document.createElement('option');o.value=x.idCliente;o.textContent=x.cliente+' (INACTIVO)';sc.appendChild(o)}if(x)sc.value=x.idCliente;q('inicio').value=x?x.inicio:'';q('fin').value=x?x.fin:'';q('observaciones').value=x?x.observaciones:'';q('mensaje').textContent='';q('bloqueEstatus').style.display=nuevo?'none':'block';q('estatusActual').textContent=x?x.estatus:'';q('cancelar').checked=x?x.estatus==='CANCELADO':false}sel.onchange=cargar;cargar();
  q('formulario').onsubmit=e=>{e.preventDefault();q('guardar').disabled=true;q('mensaje').textContent='Guardando...';google.script.run.withSuccessHandler(r=>{q('guardar').disabled=false;q('mensaje').className=r.ok?'ok':'error';q('mensaje').textContent=r.mensaje;if(r.ok)setTimeout(()=>google.script.host.close(),900)}).withFailureHandler(e=>{q('guardar').disabled=false;q('mensaje').className='error';q('mensaje').textContent=e.message||'No fue posible guardar.'}).guardarEventoDesdeFormulario({idEvento:sel.value,nombre:q('nombre').value,idCliente:sc.value,inicio:q('inicio').value,fin:q('fin').value,cancelar:q('cancelar').checked,observaciones:q('observaciones').value})};</script></body></html>`).setWidth(540).setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html,'Syhme');
}

function guardarEventoDesdeFormulario(datos) {
  const id=String((datos&&datos.idEvento)||'').trim(),nombre=normalizarNombre_(datos&&datos.nombre),idCliente=String((datos&&datos.idCliente)||'').trim(),inicio=fechaDesdeIso_(datos&&datos.inicio),fin=fechaDesdeIso_(datos&&datos.fin),cancelar=Boolean(datos&&datos.cancelar),observaciones=String((datos&&datos.observaciones)||'').trim().slice(0,300);
  if(!nombre)return{ok:false,mensaje:'Escribe el nombre del evento.'};if(!idCliente)return{ok:false,mensaje:'Selecciona un cliente.'};if(!inicio||!fin)return{ok:false,mensaje:'Captura fechas válidas.'};if(fin<inicio)return{ok:false,mensaje:'La fecha final no puede ser anterior a la inicial.'};
  const lock=LockService.getDocumentLock();if(!lock.tryLock(15000))return{ok:false,mensaje:'El sistema está ocupado. Intenta nuevamente.'};
  try{
    const libro=SpreadsheetApp.getActiveSpreadsheet(),hoja=obtenerHojaObligatoria_(libro,SYHME.HOJAS.EVENTOS),eventos=obtenerEventos_(),cliente=obtenerClientes_().find(x=>x.id===idCliente);
    if(!cliente)return{ok:false,mensaje:'El cliente seleccionado ya no existe.'};if(!id&&cliente.estatus!=='ACTIVO')return{ok:false,mensaje:'El cliente seleccionado no está activo.'};
    const isoInicio=formatearFechaIso_(inicio),duplicado=eventos.find(x=>claveTexto_(x.nombre)===claveTexto_(nombre)&&x.inicio===isoInicio&&x.id!==id);if(duplicado)return{ok:false,mensaje:`Ya existe ${duplicado.nombre} con esa fecha de inicio.`};
    const dias=Math.floor((fin-inicio)/86400000)+1;
    const estatus=cancelar?'CANCELADO':calcularEstatusEvento_(inicio,fin);

    if(id){
      const actual=eventos.find(x=>x.id===id);if(!actual)return{ok:false,mensaje:'El evento seleccionado ya no existe.'};
      const urlCarpeta=crearOEncontrarCarpetaEvento_(id,nombre);
      hoja.getRange(actual.fila,2,1,9).setValues([[nombre,idCliente,cliente.nombre,inicio,fin,dias,estatus,formulaCarpeta_(urlCarpeta),observaciones]]);
      registrarBitacora_({modulo:'EVENTOS',accion:'MODIFICACION',idRegistro:id,valorAnterior:JSON.stringify({nombre:actual.nombre,cliente:actual.cliente,inicio:actual.inicio,fin:actual.fin,estatus:actual.estatus}),valorNuevo:JSON.stringify({nombre,cliente:cliente.nombre,inicio:isoInicio,fin:formatearFechaIso_(fin),estatus}),detalle:'Evento actualizado desde el formulario.'});
      return{ok:true,mensaje:`Evento ${id} actualizado correctamente.`};
    }

    const nuevoId=generarId_('EVE',hoja);
    const urlCarpeta=crearOEncontrarCarpetaEvento_(nuevoId,nombre);
    hoja.appendRow([nuevoId,nombre,idCliente,cliente.nombre,inicio,fin,dias,estatus,formulaCarpeta_(urlCarpeta),observaciones]);
    registrarBitacora_({modulo:'EVENTOS',accion:'ALTA',idRegistro:nuevoId,valorNuevo:JSON.stringify({nombre,cliente:cliente.nombre,inicio:isoInicio,fin:formatearFechaIso_(fin),estatus,carpeta:urlCarpeta}),detalle:'Evento y carpeta de Drive registrados desde el formulario.'});

    try {
      asegurarDirectoresAsignados_(nuevoId, nombre);
    } catch (errorDirectores) {
      // No bloquea la creación del evento si falla la auto-asignación de directores.
    }

    return{ok:true,mensaje:`Evento guardado con ID ${nuevoId}. Carpeta creada y directores asignados automáticamente.`,idEvento:nuevoId};
  }finally{lock.releaseLock()}
}

function actualizarEstatusEventos_() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = obtenerHojaObligatoria_(libro, SYHME.HOJAS.EVENTOS);
  const eventos = obtenerEventos_();
  const hoy = new Date();

  eventos.forEach((evento) => {
    if (evento.estatus === 'CANCELADO') return;

    const inicio = fechaDesdeIso_(evento.inicio);
    const fin = fechaDesdeIso_(evento.fin);
    if (!inicio || !fin) return;

    const nuevoEstatus = calcularEstatusEvento_(inicio, fin, hoy);
    if (nuevoEstatus !== evento.estatus) {
      hoja.getRange(evento.fila, 8).setValue(nuevoEstatus);
      registrarBitacora_({
        modulo: 'EVENTOS',
        accion: 'ACTUALIZACION_AUTOMATICA',
        idRegistro: evento.id,
        valorAnterior: evento.estatus,
        valorNuevo: nuevoEstatus,
        detalle: 'Estatus recalculado automáticamente por fecha.',
      });
    }
  });
}

function instalarActualizacionDiariaEventos() {
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === 'actualizarEstatusEventos_')
    .forEach((t) => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('actualizarEstatusEventos_').timeBased().everyDays(1).atHour(1).create();
  SpreadsheetApp.getActive().toast('Actualización diaria de estatus activada.', 'Syhme', 5);
}

/* ======================= FIN EVENTOS ======================= */

function obtenerClientes_(){const hoja=obtenerHojaObligatoria_(SpreadsheetApp.getActiveSpreadsheet(),SYHME.HOJAS.CLIENTES),n=hoja.getLastRow();if(n<2)return[];return hoja.getRange(2,1,n-1,7).getValues().filter(f=>f[0]&&f[1]).map((f,i)=>({fila:i+2,id:String(f[0]),nombre:String(f[1]),correo:String(f[2]||''),estatus:String(f[3]),fechaAlta:f[4],observaciones:String(f[6]||'')}))}
function obtenerEventos_(){const hoja=obtenerHojaObligatoria_(SpreadsheetApp.getActiveSpreadsheet(),SYHME.HOJAS.EVENTOS),n=hoja.getLastRow();if(n<2)return[];return hoja.getRange(2,1,n-1,10).getValues().filter(f=>f[0]&&f[1]).map((f,i)=>({fila:i+2,id:String(f[0]),nombre:String(f[1]),idCliente:String(f[2]),cliente:String(f[3]),inicio:formatearFechaIso_(f[4]),fin:formatearFechaIso_(f[5]),estatus:String(f[7]),urlCarpeta:String(f[8]||''),observaciones:String(f[9]||'')}))}
function contarEventosAbiertosPorCliente_(id){return obtenerEventos_().filter(x=>x.idCliente===id&&!['FINALIZADO','CANCELADO'].includes(x.estatus)).length}
function fechaDesdeIso_(v){const m=String(v||'').match(/^(\d{4})-(\d{2})-(\d{2})$/);if(!m)return null;const f=new Date(+m[1],+m[2]-1,+m[3],12);return f.getFullYear()===+m[1]&&f.getMonth()===+m[2]-1&&f.getDate()===+m[3]?f:null}
function formatearFechaIso_(v){if(!(v instanceof Date)||isNaN(v))return'';return Utilities.formatDate(v,Session.getScriptTimeZone()||'America/Mexico_City','yyyy-MM-dd')}
function abrirCatalogoClientes(){const l=SpreadsheetApp.getActiveSpreadsheet();l.setActiveSheet(obtenerHojaObligatoria_(l,SYHME.HOJAS.CLIENTES))}
function abrirCatalogoEventos(){const l=SpreadsheetApp.getActiveSpreadsheet();l.setActiveSheet(obtenerHojaObligatoria_(l,SYHME.HOJAS.EVENTOS))}
function estilosFormulario_(){return`body{font-family:Arial,sans-serif;margin:0;padding:22px;color:#263238}h2{margin:0 0 10px;color:#2d516a}p{font-size:12px;color:#546e7a}label{display:block;margin:11px 0 5px;font-weight:700;font-size:13px}input,select,textarea{box-sizing:border-box;width:100%;padding:9px;border:1px solid #b0bec5;border-radius:6px;font-size:14px}textarea{min-height:58px;resize:vertical}.grid{display:grid;grid-template-columns:1fr 1fr;gap:12px}.acciones{display:flex;justify-content:flex-end;gap:8px;margin-top:18px}button{border:0;border-radius:6px;padding:10px 16px;cursor:pointer;font-weight:700}.primario{background:#2d516a;color:#fff}.secundario{background:#eceff1;color:#37474f}button:disabled{opacity:.65}#mensaje{min-height:18px;margin-top:11px;font-size:13px}.ok{color:#1b5e20}.error{color:#b71c1c}`}

function mostrarFormularioCotizaciones() {
  const eventos = obtenerEventos_().filter((evento) => evento.estatus !== 'CANCELADO');
  if (!eventos.length) {
    SpreadsheetApp.getUi().alert('Registra un evento que no esté cancelado antes de crear una cotización.');
    return;
  }

  const tarifaElemento = obtenerTarifaPredeterminada_('ELEMENTO');
  const tarifaCoordinador = obtenerTarifaPredeterminada_('COORDINADOR');
  if (tarifaElemento === null || tarifaCoordinador === null) {
    SpreadsheetApp.getUi().alert('Se requieren los puestos activos ELEMENTO y COORDINADOR para cotizar.');
    return;
  }

  const datos = JSON.stringify({
    eventos,
    cotizaciones: obtenerCotizaciones_(),
    tarifaElemento,
    tarifaCoordinador,
  }).replace(/</g, '\\u003c');

  const html = HtmlService.createHtmlOutput(`
    <!doctype html><html><head><base target="_top"><style>${estilosFormulario_()}</style></head><body>
      <h2>Cotizaciones</h2>
      <p>Las tarifas se copian del catálogo al crear la cotización y quedan guardadas como historial del documento.</p>
      <form id="formulario">
        <label for="idCotizacion">Registro</label>
        <select id="idCotizacion"><option value="">+ NUEVA COTIZACIÓN</option></select>
        <label for="idEvento">Evento</label>
        <select id="idEvento" required></select>
        <div class="grid">
          <div><label for="cantidadElementos">Cantidad de elementos</label><input id="cantidadElementos" type="number" min="0" step="1" required></div>
          <div><label for="tarifaElemento">Tarifa por elemento</label><input id="tarifaElemento" type="number" min="0" step="0.01" required></div>
          <div><label for="cantidadCoordinadores">Cantidad de coordinadores</label><input id="cantidadCoordinadores" type="number" min="0" step="1" required></div>
          <div><label for="tarifaCoordinador">Tarifa por coordinador</label><input id="tarifaCoordinador" type="number" min="0" step="0.01" required></div>
        </div>
        <label for="ivaPorcentaje">IVA</label>
        <select id="ivaPorcentaje"><option value="0.16">16%</option><option value="0">EXENTO</option></select>
        <label for="estatus">Estatus</label>
        <select id="estatus"><option>BORRADOR</option><option>ENVIADA</option><option>ACEPTADA</option><option>RECHAZADA</option><option>CANCELADA</option></select>
        <label for="observaciones">Observaciones</label><textarea id="observaciones" maxlength="300"></textarea>
        <div class="resumen"><span>Subtotal: <b id="subtotal">$0.00</b></span><span>IVA: <b id="iva">$0.00</b></span><span>Total: <b id="total">$0.00</b></span></div>
        <div id="mensaje"></div>
        <div class="acciones"><button type="button" class="secundario" onclick="google.script.host.close()">Cancelar</button><button id="guardar" class="primario">Guardar cotización</button></div>
      </form>
      <script>
        const datos=${datos},q=id=>document.getElementById(id),sel=q('idCotizacion'),evento=q('idEvento'),moneda=new Intl.NumberFormat('es-MX',{style:'currency',currency:'MXN'});
        datos.eventos.forEach(x=>{const o=document.createElement('option');o.value=x.id;o.textContent=x.nombre+' — '+x.inicio;evento.appendChild(o)});
        datos.cotizaciones.forEach(x=>{const o=document.createElement('option');o.value=x.id;o.textContent=x.id+' — '+x.evento;o.disabled=['ACEPTADA','CANCELADA'].includes(x.estatus);sel.appendChild(o)});
        function numero(id){return Number(q(id).value)||0}
        function actualizarResumen(){const sub=numero('cantidadElementos')*numero('tarifaElemento')+numero('cantidadCoordinadores')*numero('tarifaCoordinador'),iva=sub*numero('ivaPorcentaje');q('subtotal').textContent=moneda.format(sub);q('iva').textContent=moneda.format(iva);q('total').textContent=moneda.format(sub+iva)}
        function cargar(){const x=datos.cotizaciones.find(v=>v.id===sel.value),nuevo=!x;q('idEvento').value=x?x.idEvento:(datos.eventos[0]||{}).id||'';q('cantidadElementos').value=x?x.cantidadElementos:0;q('tarifaElemento').value=x?x.tarifaElemento:datos.tarifaElemento;q('cantidadCoordinadores').value=x?x.cantidadCoordinadores:0;q('tarifaCoordinador').value=x?x.tarifaCoordinador:datos.tarifaCoordinador;q('ivaPorcentaje').value=x?x.ivaPorcentaje:'0.16';q('estatus').value=x?x.estatus:'BORRADOR';q('observaciones').value=x?x.observaciones:'';q('idEvento').disabled=!nuevo;q('mensaje').textContent='';actualizarResumen()}
        sel.onchange=cargar;['cantidadElementos','tarifaElemento','cantidadCoordinadores','tarifaCoordinador','ivaPorcentaje'].forEach(id=>q(id).addEventListener('input',actualizarResumen));cargar();
        q('formulario').onsubmit=e=>{e.preventDefault();q('guardar').disabled=true;q('mensaje').className='';q('mensaje').textContent='Guardando...';google.script.run.withSuccessHandler(r=>{q('guardar').disabled=false;q('mensaje').className=r.ok?'ok':'error';q('mensaje').textContent=r.mensaje;if(r.ok)setTimeout(()=>google.script.host.close(),900)}).withFailureHandler(e=>{q('guardar').disabled=false;q('mensaje').className='error';q('mensaje').textContent=e.message||'No fue posible guardar la cotización.'}).guardarCotizacionDesdeFormulario({idCotizacion:sel.value,idEvento:evento.value,cantidadElementos:q('cantidadElementos').value,tarifaElemento:q('tarifaElemento').value,cantidadCoordinadores:q('cantidadCoordinadores').value,tarifaCoordinador:q('tarifaCoordinador').value,ivaPorcentaje:q('ivaPorcentaje').value,estatus:q('estatus').value,observaciones:q('observaciones').value})};
      </script>
    </body></html>
  `).setWidth(560).setHeight(760);
  SpreadsheetApp.getUi().showModalDialog(html, 'Syhme');
}

function guardarCotizacionDesdeFormulario(datos) {
  const id = String((datos && datos.idCotizacion) || '').trim();
  const idEvento = String((datos && datos.idEvento) || '').trim();
  const cantidadElementos = Number(datos && datos.cantidadElementos);
  const tarifaElemento = Number(datos && datos.tarifaElemento);
  const cantidadCoordinadores = Number(datos && datos.cantidadCoordinadores);
  const tarifaCoordinador = Number(datos && datos.tarifaCoordinador);
  const ivaPorcentaje = Number(datos && datos.ivaPorcentaje);
  const estatus = String((datos && datos.estatus) || '').trim().toUpperCase();
  const observaciones = String((datos && datos.observaciones) || '').trim().slice(0, 300);

  if (!idEvento) return { ok: false, mensaje: 'Selecciona un evento.' };
  for (const cantidad of [cantidadElementos, cantidadCoordinadores]) {
    if (!Number.isInteger(cantidad) || cantidad < 0) return { ok: false, mensaje: 'Las cantidades deben ser números enteros iguales o mayores que cero.' };
  }
  for (const tarifa of [tarifaElemento, tarifaCoordinador]) {
    if (!Number.isFinite(tarifa) || tarifa < 0) return { ok: false, mensaje: 'Las tarifas deben ser números válidos iguales o mayores que cero.' };
  }
  if (!Number.isFinite(ivaPorcentaje) || ![0, 0.16].includes(ivaPorcentaje)) return { ok: false, mensaje: 'Selecciona un IVA válido.' };
  if (!['BORRADOR', 'ENVIADA', 'ACEPTADA', 'RECHAZADA', 'CANCELADA'].includes(estatus)) return { ok: false, mensaje: 'Selecciona un estatus válido.' };
  if (cantidadElementos + cantidadCoordinadores === 0) return { ok: false, mensaje: 'Indica al menos un elemento o un coordinador.' };

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) return { ok: false, mensaje: 'El sistema está ocupado. Intenta nuevamente.' };

  try {
    const libro = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = obtenerHojaObligatoria_(libro, SYHME.HOJAS.COTIZACIONES);
    const cotizaciones = obtenerCotizaciones_();
    const evento = obtenerEventos_().find((item) => item.id === idEvento);
    if (!evento) return { ok: false, mensaje: 'El evento seleccionado ya no existe.' };
    if (!id && evento.estatus === 'CANCELADO') return { ok: false, mensaje: 'No se puede cotizar un evento cancelado.' };

    const subtotal = cantidadElementos * tarifaElemento + cantidadCoordinadores * tarifaCoordinador;
    const iva = subtotal * ivaPorcentaje;
    const total = subtotal + iva;
    const ahora = new Date();

    if (id) {
      const actual = cotizaciones.find((item) => item.id === id);
      if (!actual) return { ok: false, mensaje: 'La cotización seleccionada ya no existe.' };
      if (['ACEPTADA', 'CANCELADA'].includes(actual.estatus)) return { ok: false, mensaje: `La cotización ${actual.id} está ${actual.estatus} y no se puede modificar.` };
      hoja.getRange(actual.fila, 4, 1, 10).setValues([[
        ahora, cantidadElementos, tarifaElemento, cantidadCoordinadores, tarifaCoordinador, subtotal, iva, total, estatus, observaciones,
      ]]);
      registrarBitacora_({
        modulo: 'COTIZACIONES', accion: 'MODIFICACION', idRegistro: id,
        valorAnterior: JSON.stringify({ cantidadElementos: actual.cantidadElementos, tarifaElemento: actual.tarifaElemento, cantidadCoordinadores: actual.cantidadCoordinadores, tarifaCoordinador: actual.tarifaCoordinador, total: actual.total, estatus: actual.estatus }),
        valorNuevo: JSON.stringify({ cantidadElementos, tarifaElemento, cantidadCoordinadores, tarifaCoordinador, total, estatus }),
        detalle: 'Cotización actualizada desde el formulario.',
      });
      return { ok: true, mensaje: `Cotización ${id} actualizada. Total: ${formatearMoneda_(total)}.` };
    }

    const nuevoId = generarId_('COT', hoja);
    hoja.appendRow([nuevoId, evento.id, evento.nombre, ahora, cantidadElementos, tarifaElemento, cantidadCoordinadores, tarifaCoordinador, subtotal, iva, total, estatus, observaciones]);
    registrarBitacora_({
      modulo: 'COTIZACIONES', accion: 'ALTA', idRegistro: nuevoId,
      valorNuevo: JSON.stringify({ evento: evento.nombre, cantidadElementos, tarifaElemento, cantidadCoordinadores, tarifaCoordinador, subtotal, iva, total, estatus }),
      detalle: 'Cotización registrada desde el formulario.',
    });
    return { ok: true, mensaje: `Cotización guardada con ID ${nuevoId}. Total: ${formatearMoneda_(total)}.`, idCotizacion: nuevoId };
  } finally {
    lock.releaseLock();
  }
}

function obtenerCotizaciones_() {
  const hoja = obtenerHojaObligatoria_(SpreadsheetApp.getActiveSpreadsheet(), SYHME.HOJAS.COTIZACIONES);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return [];
  return hoja.getRange(2, 1, ultimaFila - 1, 13).getValues()
    .filter((fila) => fila[0] && fila[1])
    .map((fila, indice) => ({
      fila: indice + 2, id: String(fila[0]), idEvento: String(fila[1]), evento: String(fila[2]), fecha: formatearFechaIso_(fila[3]),
      cantidadElementos: Number(fila[4]) || 0, tarifaElemento: Number(fila[5]) || 0,
      cantidadCoordinadores: Number(fila[6]) || 0, tarifaCoordinador: Number(fila[7]) || 0,
      subtotal: Number(fila[8]) || 0, iva: Number(fila[9]) || 0, total: Number(fila[10]) || 0,
      ivaPorcentaje: Number(fila[8]) ? Number(fila[9]) / Number(fila[8]) : 0,
      estatus: String(fila[11]), observaciones: String(fila[12] || ''),
    }));
}

function obtenerTarifaPredeterminada_(nombrePuesto) {
  const puesto = obtenerPuestosActivos_().find((item) => claveTexto_(item.nombre) === claveTexto_(nombrePuesto));
  return puesto ? Number(puesto.tarifa) : null;
}

function formatearMoneda_(importe) {
  return Utilities.formatString('$%,.2f', Number(importe) || 0);
}

function abrirCatalogoCotizaciones() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  libro.setActiveSheet(obtenerHojaObligatoria_(libro, SYHME.HOJAS.COTIZACIONES));
}

/* ======================= ASIGNACIONES ======================= */

function mostrarFormularioAsignaciones() {
  const eventos = obtenerEventos_().filter((e) => e.estatus !== 'CANCELADO');
  if (!eventos.length) {
    SpreadsheetApp.getUi().alert('Registra un evento que no esté cancelado antes de crear asignaciones.');
    return;
  }

  const empleados = obtenerEmpleadosActivos_();
  if (!empleados.length) {
    SpreadsheetApp.getUi().alert('No hay empleados activos. Da de alta al menos un empleado antes de asignarlo a un evento.');
    return;
  }

  const datos = JSON.stringify({
    eventos,
    empleados,
    asignaciones: obtenerAsignaciones_(),
  }).replace(/</g, '\\u003c');

  const html = HtmlService.createHtmlOutput(`
    <!doctype html><html><head><base target="_top"><style>${estilosFormulario_()}</style></head><body>
      <h2>Asignaciones</h2>
      <p>Selecciona el evento y el empleado. La tarifa se copia del puesto y puedes ajustarla si el acuerdo para este evento es distinto.</p>
      <form id="formulario">
        <label for="idAsignacion">Registro</label>
        <select id="idAsignacion"><option value="">+ NUEVA ASIGNACIÓN</option></select>
        <label for="idEvento">Evento</label>
        <select id="idEvento" required></select>
        <label for="idEmpleado">Empleado</label>
        <select id="idEmpleado" required></select>
        <div class="grid">
          <div><label for="puestoActual">Puesto</label><input id="puestoActual" disabled></div>
          <div><label for="tarifaAcordada">Tarifa por turno acordada</label><input id="tarifaAcordada" type="number" min="0.01" step="0.01" required></div>
        </div>
        <label for="estatus">Estatus</label>
        <select id="estatus"><option>ACTIVO</option><option>INACTIVO</option></select>
        <label for="observaciones">Observaciones</label><textarea id="observaciones" maxlength="300"></textarea>
        <div id="mensaje"></div>
        <div class="acciones"><button type="button" class="secundario" onclick="google.script.host.close()">Cancelar</button><button id="guardar" class="primario">Guardar asignación</button></div>
      </form>
      <script>
        const datos=${datos},q=id=>document.getElementById(id),sel=q('idAsignacion'),se=q('idEvento'),sp=q('idEmpleado');
        datos.eventos.forEach(x=>{const o=document.createElement('option');o.value=x.id;o.textContent=x.nombre+' — '+x.inicio;se.appendChild(o)});
        datos.empleados.forEach(x=>{const o=document.createElement('option');o.value=x.id;o.textContent=x.nombre+' ('+x.puesto+')';sp.appendChild(o)});
        datos.asignaciones.forEach(x=>{const o=document.createElement('option');o.value=x.id;o.textContent=x.id+' — '+x.empleado+' — '+x.evento;sel.appendChild(o)});

        function tarifaPuesto(idEmpleado){const e=datos.empleados.find(x=>x.id===idEmpleado);return e?e.tarifaPuesto:0}
        function puestoDe(idEmpleado){const e=datos.empleados.find(x=>x.id===idEmpleado);return e?e.puesto:''}

        function cargar(){
          const x=datos.asignaciones.find(v=>v.id===sel.value),nuevo=!x;
          if(x){
            se.value=x.idEvento;
            if(![...sp.options].some(o=>o.value===x.idEmpleado)){const o=document.createElement('option');o.value=x.idEmpleado;o.textContent=x.empleado+' ('+x.puesto+') (INACTIVO)';sp.appendChild(o)}
            sp.value=x.idEmpleado;
            q('puestoActual').value=x.puesto;
            q('tarifaAcordada').value=x.tarifaAcordada;
            q('estatus').value=x.estatus;
            q('observaciones').value=x.observaciones;
          } else {
            q('puestoActual').value=puestoDe(sp.value);
            q('tarifaAcordada').value=tarifaPuesto(sp.value);
            q('estatus').value='ACTIVO';
            q('observaciones').value='';
          }
          se.disabled=!nuevo;sp.disabled=!nuevo;q('mensaje').textContent='';
        }
        sel.onchange=cargar;
        sp.onchange=()=>{if(!sel.value){q('puestoActual').value=puestoDe(sp.value);q('tarifaAcordada').value=tarifaPuesto(sp.value)}};
        cargar();

        q('formulario').onsubmit=e=>{e.preventDefault();q('guardar').disabled=true;q('mensaje').className='';q('mensaje').textContent='Guardando...';
          google.script.run.withSuccessHandler(r=>{q('guardar').disabled=false;q('mensaje').className=r.ok?'ok':'error';q('mensaje').textContent=r.mensaje;if(r.ok)setTimeout(()=>google.script.host.close(),900)})
          .withFailureHandler(e=>{q('guardar').disabled=false;q('mensaje').className='error';q('mensaje').textContent=e.message||'No fue posible guardar la asignación.'})
          .guardarAsignacionDesdeFormulario({idAsignacion:sel.value,idEvento:se.value,idEmpleado:sp.value,tarifaAcordada:q('tarifaAcordada').value,estatus:q('estatus').value,observaciones:q('observaciones').value})};
      </script>
    </body></html>
  `).setWidth(540).setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, 'Syhme');
}

function guardarAsignacionDesdeFormulario(datos) {
  const id = String((datos && datos.idAsignacion) || '').trim();
  const idEvento = String((datos && datos.idEvento) || '').trim();
  const idEmpleado = String((datos && datos.idEmpleado) || '').trim();
  const tarifaAcordada = Number(datos && datos.tarifaAcordada);
  const estatus = String((datos && datos.estatus) || '').trim().toUpperCase();
  const observaciones = String((datos && datos.observaciones) || '').trim().slice(0, 300);

  if (!idEvento) return { ok: false, mensaje: 'Selecciona un evento.' };
  if (!idEmpleado) return { ok: false, mensaje: 'Selecciona un empleado.' };
  if (!Number.isFinite(tarifaAcordada) || tarifaAcordada <= 0) return { ok: false, mensaje: 'La tarifa debe ser mayor que cero.' };
  if (!['ACTIVO', 'INACTIVO'].includes(estatus)) return { ok: false, mensaje: 'Selecciona un estatus válido.' };

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) return { ok: false, mensaje: 'El sistema está ocupado. Intenta nuevamente.' };

  try {
    const libro = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = obtenerHojaObligatoria_(libro, SYHME.HOJAS.ASIGNACIONES);
    const asignaciones = obtenerAsignaciones_();
    const evento = obtenerEventos_().find((e) => e.id === idEvento);
    if (!evento) return { ok: false, mensaje: 'El evento seleccionado ya no existe.' };
    if (!id && evento.estatus === 'CANCELADO') return { ok: false, mensaje: 'No se puede asignar personal a un evento cancelado.' };

    const empleado = obtenerEmpleados_().find((e) => e.id === idEmpleado);
    if (!empleado) return { ok: false, mensaje: 'El empleado seleccionado ya no existe.' };
    if (!id && empleado.estatus !== SYHME.ESTATUS_ACTIVO) return { ok: false, mensaje: 'El empleado seleccionado no está activo.' };

    const duplicado = asignaciones.find((a) => a.idEvento === idEvento && a.idEmpleado === idEmpleado && a.estatus === 'ACTIVO' && a.id !== id);
    if (duplicado) return { ok: false, mensaje: `${empleado.nombre} ya está asignado a este evento (${duplicado.id}).` };

    if (id) {
      const actual = asignaciones.find((a) => a.id === id);
      if (!actual) return { ok: false, mensaje: 'La asignación seleccionada ya no existe.' };
      hoja.getRange(actual.fila, 7, 1, 3).setValues([[tarifaAcordada, estatus, observaciones]]);
      registrarBitacora_({
        modulo: 'ASIGNACIONES', accion: 'MODIFICACION', idRegistro: id,
        valorAnterior: JSON.stringify({ tarifaAcordada: actual.tarifaAcordada, estatus: actual.estatus }),
        valorNuevo: JSON.stringify({ tarifaAcordada, estatus }),
        detalle: 'Asignación actualizada desde el formulario.',
      });
      return { ok: true, mensaje: `Asignación ${id} actualizada correctamente.` };
    }

    const nuevoId = generarId_('ASG', hoja);
    hoja.appendRow([nuevoId, idEvento, idEmpleado, empleado.nombre, empleado.idPuesto, empleado.puesto, tarifaAcordada, estatus, observaciones]);
    registrarBitacora_({
      modulo: 'ASIGNACIONES', accion: 'ALTA', idRegistro: nuevoId,
      valorNuevo: JSON.stringify({ evento: evento.nombre, empleado: empleado.nombre, puesto: empleado.puesto, tarifaAcordada, estatus }),
      detalle: 'Asignación registrada desde el formulario.',
    });
    return { ok: true, mensaje: `Asignación guardada con ID ${nuevoId}.`, idAsignacion: nuevoId };
  } finally {
    lock.releaseLock();
  }
}

function obtenerAsignaciones_() {
  const hoja = obtenerHojaObligatoria_(SpreadsheetApp.getActiveSpreadsheet(), SYHME.HOJAS.ASIGNACIONES);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return [];
  return hoja.getRange(2, 1, ultimaFila - 1, 9).getValues()
    .filter((f) => f[0] && f[1])
    .map((f, i) => ({
      fila: i + 2, id: String(f[0]), idEvento: String(f[1]), idEmpleado: String(f[2]), empleado: String(f[3]),
      idPuesto: String(f[4]), puesto: String(f[5]), tarifaAcordada: Number(f[6]) || 0, estatus: String(f[7]), observaciones: String(f[8] || ''),
    }))
    .map((a) => {
      const evento = obtenerEventos_().find((e) => e.id === a.idEvento);
      return Object.assign({}, a, { evento: evento ? evento.nombre : a.idEvento });
    });
}

function obtenerAsignacionesActivasPorEvento_(idEvento) {
  return obtenerAsignaciones_().filter((a) => a.idEvento === idEvento && a.estatus === 'ACTIVO');
}

function asegurarDirectoresAsignados_(idEvento, nombreEvento) {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = obtenerHojaObligatoria_(libro, SYHME.HOJAS.ASIGNACIONES);
  const directores = obtenerEmpleadosActivos_().filter((e) => claveTexto_(e.puesto) === claveTexto_(SYHME.PUESTO_DIRECCION));
  if (!directores.length) return;

  const asignacionesActuales = obtenerAsignaciones_();
  const puestoDireccion = obtenerPuestos_().find((p) => claveTexto_(p.nombre) === claveTexto_(SYHME.PUESTO_DIRECCION));
  const tarifaDireccion = puestoDireccion ? puestoDireccion.tarifa : 0;

  directores.forEach((director) => {
    const yaAsignado = asignacionesActuales.find((a) => a.idEvento === idEvento && a.idEmpleado === director.id && a.estatus === 'ACTIVO');
    if (yaAsignado) return;

    const nuevoId = generarId_('ASG', hoja);
    hoja.appendRow([nuevoId, idEvento, director.id, director.nombre, director.idPuesto, director.puesto, tarifaDireccion, 'ACTIVO', 'Asignación automática de dirección']);
    registrarBitacora_({
      modulo: 'ASIGNACIONES', accion: 'ALTA_AUTOMATICA', idRegistro: nuevoId,
      valorNuevo: JSON.stringify({ evento: nombreEvento, empleado: director.nombre, tarifaAcordada: tarifaDireccion }),
      detalle: 'Dirección asignada automáticamente al crear el evento.',
    });
  });
}

function abrirCatalogoAsignaciones() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  libro.setActiveSheet(obtenerHojaObligatoria_(libro, SYHME.HOJAS.ASIGNACIONES));
}

/* ======================= FIN ASIGNACIONES ======================= */

/* ======================= TURNOS (casillas de un solo clic, coloreadas por día) ======================= */

function mostrarFormularioTurnos() {
  const eventos = obtenerEventos_().filter((e) => e.estatus !== 'CANCELADO');
  if (!eventos.length) {
    SpreadsheetApp.getUi().alert('No hay eventos activos para capturar turnos.');
    return;
  }

  const asignaciones = obtenerAsignaciones_().filter((a) => a.estatus === 'ACTIVO');
  if (!asignaciones.length) {
    SpreadsheetApp.getUi().alert('No hay asignaciones activas. Asigna personal a un evento antes de capturar turnos.');
    return;
  }

  const datos = JSON.stringify({
    eventos,
    asignaciones,
    turnos: obtenerTurnos_(),
    tiposTurno: Object.keys(SYHME.FRACCIONES_TURNO).filter((t) => t !== ''),
    paleta: SYHME.PALETA_DIAS,
  }).replace(/</g, '\\u003c');

  const html = HtmlService.createHtmlOutput(`
    <!doctype html><html><head><base target="_top"><style>${estilosFormulario_()}
      #tablaDias{max-height:340px;overflow-y:auto;margin-top:10px;padding-right:4px}
      .fila-dia{display:flex;align-items:center;gap:10px;padding:6px 0;border-bottom:1px solid #eceff1}
      .etiqueta-dia{width:70px;font-size:12px;font-weight:700;color:#37474f;flex-shrink:0}
      .opciones-turno{display:flex;gap:6px}
      .opciones-turno label{margin:0;display:block}
      .opciones-turno input{position:absolute;opacity:0;width:0;height:0}
      .casilla{width:38px;height:32px;border:2px solid #cfd8dc;border-radius:6px;display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:700;cursor:pointer;color:#78909c;user-select:none}
      .resumen-turnos{margin-top:10px;font-size:13px;color:#2d516a;font-weight:700}
    </style></head><body>
      <h2>Turnos</h2>
      <p>Elige el evento y el empleado asignado. Marca con un clic el turno de cada día y guarda todo de una vez.</p>
      <form id="formulario">
        <label for="idEvento">Evento</label>
        <select id="idEvento" required></select>
        <label for="idAsignacion">Empleado asignado</label>
        <select id="idAsignacion" required></select>
        <div id="tablaDias"></div>
        <div class="resumen-turnos" id="resumenTurnos"></div>
        <div id="mensaje"></div>
        <div class="acciones"><button type="button" class="secundario" onclick="google.script.host.close()">Cerrar</button><button id="guardar" class="primario">Guardar turnos</button></div>
      </form>
      <script>
        const datos=${datos},q=id=>document.getElementById(id),se=q('idEvento'),sa=q('idAsignacion'),contenedor=q('tablaDias');
        const nombresDias=['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado'];
        const fracciones={'':0,'1/2':0.5,'12':1,'12 1/2':1.5,'24':2};

        datos.eventos.forEach(x=>{const o=document.createElement('option');o.value=x.id;o.textContent=x.nombre+' — '+x.inicio+' a '+x.fin;se.appendChild(o)});

        function fechasEntre(inicioIso,finIso){
          const dias=[],inicio=new Date(inicioIso+'T12:00:00'),fin=new Date(finIso+'T12:00:00');
          for(let f=new Date(inicio);f<=fin;f.setDate(f.getDate()+1)){dias.push(f.toISOString().slice(0,10))}
          return dias;
        }

        function cargarAsignaciones(){
          sa.innerHTML='';
          const asigEvento=datos.asignaciones.filter(a=>a.idEvento===se.value);
          asigEvento.forEach(a=>{const o=document.createElement('option');o.value=a.id;o.textContent=a.empleado+' ('+a.puesto+')';sa.appendChild(o)});
          cargarDias();
        }

        function cargarDias(){
          contenedor.innerHTML='';
          const evento=datos.eventos.find(x=>x.id===se.value);
          if(!evento||!sa.value){actualizarResumen();return}
          const dias=fechasEntre(evento.inicio,evento.fin);
          const turnosExistentes=datos.turnos.filter(t=>t.idEvento===se.value&&t.idAsignacion===sa.value);

          dias.forEach((fechaIso,indice)=>{
            const color=datos.paleta[indice%datos.paleta.length];
            const fechaObj=new Date(fechaIso+'T12:00:00');
            const etiqueta=String(fechaObj.getDate()).padStart(2,'0')+' '+nombresDias[fechaObj.getDay()];
            const existente=turnosExistentes.find(t=>t.fecha===fechaIso);

            const fila=document.createElement('div');fila.className='fila-dia';
            const nombreDia=document.createElement('div');nombreDia.className='etiqueta-dia';nombreDia.textContent=etiqueta;
            const opciones=document.createElement('div');opciones.className='opciones-turno';

            datos.tiposTurno.forEach(tipo=>{
              const label=document.createElement('label');
              const input=document.createElement('input');
              input.type='radio';input.name='dia-'+fechaIso;input.value=tipo;
              const caja=document.createElement('div');caja.className='casilla';caja.textContent=tipo;
              caja.style.borderColor=color.fondo;

              if(existente&&existente.tipo===tipo){input.checked=true;caja.style.background=color.activo;caja.style.color='#fff';caja.style.borderColor=color.activo}

              input.addEventListener('change',()=>{
                [...opciones.querySelectorAll('.casilla')].forEach(c=>{c.style.background='';c.style.color='#78909c';c.style.borderColor=color.fondo});
                caja.style.background=color.activo;caja.style.color='#fff';caja.style.borderColor=color.activo;
                actualizarResumen();
              });

              label.appendChild(input);label.appendChild(caja);
              opciones.appendChild(label);
            });

            fila.dataset.fecha=fechaIso;
            fila.appendChild(nombreDia);fila.appendChild(opciones);
            contenedor.appendChild(fila);
          });
          actualizarResumen();
        }

        function actualizarResumen(){
          let total=0,capturados=0;
          [...contenedor.querySelectorAll('.fila-dia')].forEach(fila=>{
            const marcado=fila.querySelector('input:checked');
            if(marcado){total+=fracciones[marcado.value]||0;capturados++}
          });
          q('resumenTurnos').textContent='Días con turno: '+capturados+' — Total de turnos: '+total;
        }

        se.addEventListener('change',cargarAsignaciones);
        sa.addEventListener('change',cargarDias);
        cargarAsignaciones();

        q('formulario').onsubmit=e=>{
          e.preventDefault();
          const turnos=[...contenedor.querySelectorAll('.fila-dia')].map(fila=>{
            const marcado=fila.querySelector('input:checked');
            return{fecha:fila.dataset.fecha,tipo:marcado?marcado.value:''};
          });
          q('guardar').disabled=true;q('mensaje').className='';q('mensaje').textContent='Guardando...';
          google.script.run.withSuccessHandler(r=>{q('guardar').disabled=false;q('mensaje').className=r.ok?'ok':'error';q('mensaje').textContent=r.mensaje})
            .withFailureHandler(err=>{q('guardar').disabled=false;q('mensaje').className='error';q('mensaje').textContent=err.message||'No fue posible guardar los turnos.'})
            .guardarTurnosDesdeFormulario({idEvento:se.value,idAsignacion:sa.value,turnos});
        };
      </script>
    </body></html>
  `).setWidth(420).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Syhme');
}

function guardarTurnosDesdeFormulario(datos) {
  const idEvento = String((datos && datos.idEvento) || '').trim();
  const idAsignacion = String((datos && datos.idAsignacion) || '').trim();
  const turnos = Array.isArray(datos && datos.turnos) ? datos.turnos : [];

  if (!idEvento) return { ok: false, mensaje: 'Selecciona un evento.' };
  if (!idAsignacion) return { ok: false, mensaje: 'Selecciona un empleado asignado.' };
  if (!turnos.length) return { ok: false, mensaje: 'No hay días para guardar.' };

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(20000)) return { ok: false, mensaje: 'El sistema está ocupado. Intenta nuevamente.' };

  try {
    const libro = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = obtenerHojaObligatoria_(libro, SYHME.HOJAS.TURNOS);

    const evento = obtenerEventos_().find((e) => e.id === idEvento);
    if (!evento) return { ok: false, mensaje: 'El evento seleccionado ya no existe.' };

    const asignacion = obtenerAsignaciones_().find((a) => a.id === idAsignacion);
    if (!asignacion) return { ok: false, mensaje: 'La asignación seleccionada ya no existe.' };
    if (asignacion.idEvento !== idEvento) return { ok: false, mensaje: 'La asignación no corresponde a este evento.' };

    const turnosExistentes = obtenerTurnos_().filter((t) => t.idEvento === idEvento && t.idAsignacion === idAsignacion);
    let creados = 0;
    let actualizados = 0;
    let totalFracciones = 0;

    turnos.forEach(({ fecha, tipo }) => {
      const fechaIso = String(fecha || '').trim();
      const tipoTurno = String(tipo || '').trim();
      if (!fechaIso) return;

      const fraccion = SYHME.FRACCIONES_TURNO.hasOwnProperty(tipoTurno) ? SYHME.FRACCIONES_TURNO[tipoTurno] : 0;
      totalFracciones += fraccion;

      const existente = turnosExistentes.find((t) => t.fecha === fechaIso);
      if (existente) {
        hoja.getRange(existente.fila, 8, 1, 2).setValues([[tipoTurno, fraccion]]);
        actualizados++;
      } else if (tipoTurno) {
        const nuevoId = generarId_('TUR', hoja);
        hoja.appendRow([
          nuevoId, idEvento, evento.nombre, idAsignacion, asignacion.idEmpleado, asignacion.empleado,
          fechaDesdeIso_(fechaIso), tipoTurno, fraccion, '',
        ]);
        creados++;
      }
    });

    registrarBitacora_({
      modulo: 'TURNOS', accion: 'CAPTURA', idRegistro: idAsignacion,
      valorNuevo: JSON.stringify({ evento: evento.nombre, empleado: asignacion.empleado, diasCreados: creados, diasActualizados: actualizados, totalFracciones }),
      detalle: 'Turnos capturados desde el formulario.',
    });

    return {
      ok: true,
      mensaje: `Turnos guardados para ${asignacion.empleado}. Días nuevos: ${creados}, actualizados: ${actualizados}. Total de turnos: ${totalFracciones}.`,
    };
  } finally {
    lock.releaseLock();
  }
}

function obtenerTurnos_() {
  const hoja = obtenerHojaObligatoria_(SpreadsheetApp.getActiveSpreadsheet(), SYHME.HOJAS.TURNOS);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return [];
  return hoja.getRange(2, 1, ultimaFila - 1, 10).getValues()
    .filter((f) => f[0] && f[1])
    .map((f, i) => ({
      fila: i + 2, id: String(f[0]), idEvento: String(f[1]), evento: String(f[2]),
      idAsignacion: String(f[3]), idEmpleado: String(f[4]), empleado: String(f[5]),
      fecha: formatearFechaIso_(f[6]), tipo: String(f[7] || ''), fraccion: Number(f[8]) || 0,
      observaciones: String(f[9] || ''),
    }));
}

function abrirCatalogoTurnos() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  libro.setActiveSheet(obtenerHojaObligatoria_(libro, SYHME.HOJAS.TURNOS));
}

/* ======================= FIN TURNOS ======================= */

/* ======================= PRESTAMOS + MOVIMIENTOS_PRESTAMO ======================= */

function mostrarFormularioPrestamos() {
  const empleados = obtenerEmpleadosActivos_();
  if (!empleados.length) {
    SpreadsheetApp.getUi().alert('No hay empleados activos. Da de alta al menos un empleado antes de registrar un préstamo.');
    return;
  }

  const datos = JSON.stringify({ empleados }).replace(/</g, '\\u003c');

  const html = HtmlService.createHtmlOutput(`
    <!doctype html><html><head><base target="_top"><style>${estilosFormulario_()}</style></head><body>
      <h2>Nuevo préstamo</h2>
      <p>El saldo pendiente se descontará automáticamente de la siguiente nómina que se genere para este empleado.</p>
      <form id="formulario">
        <label for="idEmpleado">Empleado</label>
        <select id="idEmpleado" required></select>
        <div class="grid">
          <div><label for="importe">Importe del préstamo</label><input id="importe" type="number" min="0.01" step="0.01" required></div>
          <div><label for="fecha">Fecha del préstamo</label><input id="fecha" type="date" required></div>
        </div>
        <label for="observaciones">Observaciones</label><textarea id="observaciones" maxlength="300"></textarea>
        <div id="mensaje"></div>
        <div class="acciones"><button type="button" class="secundario" onclick="google.script.host.close()">Cancelar</button><button id="guardar" class="primario">Guardar préstamo</button></div>
      </form>
      <script>
        const datos=${datos},q=id=>document.getElementById(id),se=q('idEmpleado');
        datos.empleados.forEach(x=>{const o=document.createElement('option');o.value=x.id;o.textContent=x.nombre+' ('+x.puesto+')';se.appendChild(o)});
        q('fecha').valueAsDate=new Date();

        q('formulario').onsubmit=e=>{
          e.preventDefault();q('guardar').disabled=true;q('mensaje').className='';q('mensaje').textContent='Guardando...';
          google.script.run.withSuccessHandler(r=>{q('guardar').disabled=false;q('mensaje').className=r.ok?'ok':'error';q('mensaje').textContent=r.mensaje;if(r.ok)setTimeout(()=>google.script.host.close(),900)})
            .withFailureHandler(err=>{q('guardar').disabled=false;q('mensaje').className='error';q('mensaje').textContent=err.message||'No fue posible guardar el préstamo.'})
            .guardarPrestamoDesdeFormulario({idEmpleado:se.value,importe:q('importe').value,fecha:q('fecha').value,observaciones:q('observaciones').value});
        };
      </script>
    </body></html>
  `).setWidth(460).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Syhme');
}

function guardarPrestamoDesdeFormulario(datos) {
  const idEmpleado = String((datos && datos.idEmpleado) || '').trim();
  const importe = Number(datos && datos.importe);
  const fecha = fechaDesdeIso_(datos && datos.fecha);
  const observaciones = String((datos && datos.observaciones) || '').trim().slice(0, 300);

  if (!idEmpleado) return { ok: false, mensaje: 'Selecciona un empleado.' };
  if (!Number.isFinite(importe) || importe <= 0) return { ok: false, mensaje: 'El importe debe ser mayor que cero.' };
  if (!fecha) return { ok: false, mensaje: 'Captura una fecha válida.' };

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) return { ok: false, mensaje: 'El sistema está ocupado. Intenta nuevamente.' };

  try {
    const libro = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = obtenerHojaObligatoria_(libro, SYHME.HOJAS.PRESTAMOS);
    const empleado = obtenerEmpleados_().find((e) => e.id === idEmpleado);
    if (!empleado) return { ok: false, mensaje: 'El empleado seleccionado ya no existe.' };

    const nuevoId = generarId_('PRE', hoja);
    hoja.appendRow([nuevoId, idEmpleado, empleado.nombre, fecha, importe, importe, 'ACTIVO', '', observaciones, new Date()]);

    registrarBitacora_({
      modulo: 'PRESTAMOS', accion: 'ALTA', idRegistro: nuevoId,
      valorNuevo: JSON.stringify({ empleado: empleado.nombre, importe }),
      detalle: 'Préstamo registrado desde el formulario.',
    });

    return { ok: true, mensaje: `Préstamo guardado con ID ${nuevoId} por ${formatearMoneda_(importe)}.`, idPrestamo: nuevoId };
  } finally {
    lock.releaseLock();
  }
}

function obtenerPrestamos_() {
  const hoja = obtenerHojaObligatoria_(SpreadsheetApp.getActiveSpreadsheet(), SYHME.HOJAS.PRESTAMOS);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return [];
  return hoja.getRange(2, 1, ultimaFila - 1, 10).getValues()
    .filter((f) => f[0] && f[1])
    .map((f, i) => ({
      fila: i + 2, id: String(f[0]), idEmpleado: String(f[1]), empleado: String(f[2]),
      fechaPrestamo: formatearFechaIso_(f[3]), importeOriginal: Number(f[4]) || 0, saldo: Number(f[5]) || 0,
      estatus: String(f[6] || ''), fechaLiquidacion: formatearFechaIso_(f[7]), observaciones: String(f[8] || ''),
    }));
}

function obtenerPrestamosActivosPorEmpleado_(idEmpleado) {
  return obtenerPrestamos_().filter((p) => p.idEmpleado === idEmpleado && p.estatus === 'ACTIVO' && p.saldo > 0);
}

function registrarAbonoPrestamo_(hojaPrestamos, hojaMovimientos, prestamo, importeAbono, idNomina) {
  const nuevoSaldo = Math.max(0, prestamo.saldo - importeAbono);
  const seLiquida = nuevoSaldo <= 0;

  hojaPrestamos.getRange(prestamo.fila, 6).setValue(nuevoSaldo);
  if (seLiquida) {
    hojaPrestamos.getRange(prestamo.fila, 7, 1, 2).setValues([['LIQUIDADO', new Date()]]);
  }

  const nuevoIdMovimiento = generarId_('MOV', hojaMovimientos);
  hojaMovimientos.appendRow([
    nuevoIdMovimiento, prestamo.id, prestamo.idEmpleado, prestamo.empleado,
    new Date(), 'ABONO', importeAbono, nuevoSaldo, idNomina, 'Abono automático generado por nómina',
  ]);

  registrarBitacora_({
    modulo: 'PRESTAMOS', accion: 'ABONO_AUTOMATICO', idRegistro: prestamo.id,
    valorAnterior: JSON.stringify({ saldo: prestamo.saldo }),
    valorNuevo: JSON.stringify({ saldo: nuevoSaldo, abono: importeAbono, idNomina }),
    detalle: seLiquida ? 'Préstamo liquidado desde nómina.' : 'Abono parcial registrado desde nómina.',
  });

  return nuevoSaldo;
}

function obtenerMovimientosPrestamo_() {
  const hoja = obtenerHojaObligatoria_(SpreadsheetApp.getActiveSpreadsheet(), SYHME.HOJAS.MOVIMIENTOS_PRESTAMO);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return [];
  return hoja.getRange(2, 1, ultimaFila - 1, 10).getValues()
    .filter((f) => f[0] && f[1])
    .map((f, i) => ({
      fila: i + 2, id: String(f[0]), idPrestamo: String(f[1]), idEmpleado: String(f[2]), empleado: String(f[3]),
      fecha: f[4], tipo: String(f[5] || ''), importe: Number(f[6]) || 0, saldoResultante: Number(f[7]) || 0,
      idNomina: String(f[8] || ''), observaciones: String(f[9] || ''),
    }));
}

function abrirCatalogoPrestamos() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  libro.setActiveSheet(obtenerHojaObligatoria_(libro, SYHME.HOJAS.PRESTAMOS));
}

function abrirCatalogoMovimientosPrestamo() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  libro.setActiveSheet(obtenerHojaObligatoria_(libro, SYHME.HOJAS.MOVIMIENTOS_PRESTAMO));
}

/* ======================= FIN PRESTAMOS ======================= */

/* ======================= NOMINAS + DETALLE_NOMINA (con descuento automático de préstamos) ======================= */

function mostrarFormularioNominas() {
  const eventos = obtenerEventos_().filter((e) => e.estatus !== 'CANCELADO');
  if (!eventos.length) {
    SpreadsheetApp.getUi().alert('No hay eventos activos para generar nómina.');
    return;
  }

  const nominas = obtenerNominas_();
  const datos = JSON.stringify({ eventos, nominas }).replace(/</g, '\\u003c');

  const html = HtmlService.createHtmlOutput(`
    <!doctype html><html><head><base target="_top"><style>${estilosFormulario_()}
      .info-nomina{background:#eceff1;border-radius:6px;padding:9px;font-size:12px;color:#37474f;margin-top:6px}
    </style></head><body>
      <h2>Generar nómina</h2>
      <p>Selecciona el evento. Si ya existe una nómina generada, se regenerará usando las asignaciones y turnos actuales. Los préstamos activos se descuentan automáticamente.</p>
      <form id="formulario">
        <label for="idEvento">Evento</label>
        <select id="idEvento" required></select>
        <div class="info-nomina" id="infoNomina"></div>
        <div id="mensaje"></div>
        <div class="acciones"><button type="button" class="secundario" onclick="google.script.host.close()">Cerrar</button><button id="guardar" class="primario">Generar / regenerar nómina</button></div>
      </form>
      <script>
        const datos=${datos},q=id=>document.getElementById(id),se=q('idEvento');
        datos.eventos.forEach(x=>{const o=document.createElement('option');o.value=x.id;o.textContent=x.nombre+' — '+x.inicio;se.appendChild(o)});

        function actualizarInfo(){
          const existente=datos.nominas.find(n=>n.idEvento===se.value);
          q('infoNomina').textContent=existente?('Ya existe la nómina '+existente.id+' generada el '+existente.fechaGeneracion+'. Se regenerará.'):'Aún no se ha generado una nómina para este evento.';
        }
        se.addEventListener('change',actualizarInfo);actualizarInfo();

        q('formulario').onsubmit=e=>{
          e.preventDefault();q('guardar').disabled=true;q('mensaje').className='';q('mensaje').textContent='Generando...';
          google.script.run.withSuccessHandler(r=>{q('guardar').disabled=false;q('mensaje').className=r.ok?'ok':'error';q('mensaje').textContent=r.mensaje;if(r.ok)actualizarInfo()})
            .withFailureHandler(err=>{q('guardar').disabled=false;q('mensaje').className='error';q('mensaje').textContent=err.message||'No fue posible generar la nómina.'})
            .generarNominaDesdeFormulario({idEvento:se.value});
        };
      </script>
    </body></html>
  `).setWidth(460).setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(html, 'Syhme');
}

function generarNominaDesdeFormulario(datos) {
  const idEvento = String((datos && datos.idEvento) || '').trim();
  if (!idEvento) return { ok: false, mensaje: 'Selecciona un evento.' };

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(25000)) return { ok: false, mensaje: 'El sistema está ocupado. Intenta nuevamente.' };

  try {
    return generarNominaEvento_(idEvento);
  } finally {
    lock.releaseLock();
  }
}

function generarNominaEvento_(idEvento) {
  const evento = obtenerEventos_().find((e) => e.id === idEvento);
  if (!evento) return { ok: false, mensaje: 'El evento seleccionado ya no existe.' };

  const asignaciones = obtenerAsignacionesActivasPorEvento_(idEvento);
  if (!asignaciones.length) return { ok: false, mensaje: 'Este evento no tiene asignaciones activas.' };

  const turnos = obtenerTurnos_().filter((t) => t.idEvento === idEvento);
  const puestos = obtenerPuestos_();

  const filasDetalle = asignaciones.map((a) => {
    const totalTurnos = turnos
      .filter((t) => t.idAsignacion === a.id)
      .reduce((suma, t) => suma + (t.fraccion || 0), 0);
    const puesto = puestos.find((p) => p.id === a.idPuesto);
    const subtotal = a.tarifaAcordada * totalTurnos;
    return {
      idAsignacion: a.id, idEmpleado: a.idEmpleado, empleado: a.empleado,
      idPuesto: a.idPuesto, puesto: a.puesto, tarifa: a.tarifaAcordada,
      totalTurnos, subtotal, orden: puesto ? puesto.orden : 999,
    };
  }).sort((x, y) => x.orden - y.orden || x.empleado.localeCompare(y.empleado, 'es'));

  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hojaNominas = obtenerHojaObligatoria_(libro, SYHME.HOJAS.NOMINAS);
  const hojaDetalle = obtenerHojaObligatoria_(libro, SYHME.HOJAS.DETALLE_NOMINA);
  const hojaPrestamos = obtenerHojaObligatoria_(libro, SYHME.HOJAS.PRESTAMOS);
  const hojaMovimientos = obtenerHojaObligatoria_(libro, SYHME.HOJAS.MOVIMIENTOS_PRESTAMO);

  const nominas = obtenerNominas_();
  const nominaExistente = nominas.find((n) => n.idEvento === idEvento);

  let idNomina;
  if (nominaExistente) {
    idNomina = nominaExistente.id;
    eliminarDetalleNominaPorId_(hojaDetalle, idNomina);
    hojaNominas.getRange(nominaExistente.fila, 4).setValue(new Date());
  } else {
    idNomina = generarId_('NOM', hojaNominas);
    const filaNueva = hojaNominas.getLastRow() + 1;
    hojaNominas.appendRow([
      idNomina, idEvento, evento.nombre, new Date(),
      `=COUNTIF(${SYHME.HOJAS.DETALLE_NOMINA}!B:B,A${filaNueva})`,
      `=SUMIF(${SYHME.HOJAS.DETALLE_NOMINA}!B:B,A${filaNueva},${SYHME.HOJAS.DETALLE_NOMINA}!I:I)`,
      `=SUMIF(${SYHME.HOJAS.DETALLE_NOMINA}!B:B,A${filaNueva},${SYHME.HOJAS.DETALLE_NOMINA}!L:L)`,
      'GENERADA', '',
    ]);
  }

  let empleadosConAbono = 0;

  filasDetalle.forEach((d) => {
    const prestamosActivos = obtenerPrestamosActivosPorEmpleado_(d.idEmpleado);
    const saldoPendienteTotal = prestamosActivos.reduce((suma, p) => suma + p.saldo, 0);
    const abonoDisponible = Math.min(saldoPendienteTotal, Math.max(0, d.subtotal));

    let restantePorAbonar = abonoDisponible;
    if (abonoDisponible > 0) {
      empleadosConAbono++;
      for (const prestamo of prestamosActivos) {
        if (restantePorAbonar <= 0) break;
        const abonoAEstePrestamo = Math.min(prestamo.saldo, restantePorAbonar);
        registrarAbonoPrestamo_(hojaPrestamos, hojaMovimientos, prestamo, abonoAEstePrestamo, idNomina);
        restantePorAbonar -= abonoAEstePrestamo;
      }
    }

    const nuevoIdDetalle = generarId_('DET', hojaDetalle);
    const filaDetalle = hojaDetalle.getLastRow() + 1;
    hojaDetalle.appendRow([
      nuevoIdDetalle, idNomina, idEvento, d.idEmpleado, d.empleado, d.idPuesto, d.puesto,
      d.tarifa, d.totalTurnos,
      `=H${filaDetalle}*I${filaDetalle}`,
      abonoDisponible,
      `=J${filaDetalle}-K${filaDetalle}`,
      'PENDIENTE',
    ]);
  });

  registrarBitacora_({
    modulo: 'NOMINAS', accion: nominaExistente ? 'REGENERACION' : 'ALTA', idRegistro: idNomina,
    valorNuevo: JSON.stringify({ evento: evento.nombre, empleados: filasDetalle.length, conAbonoPrestamo: empleadosConAbono }),
    detalle: nominaExistente ? 'Nómina regenerada desde el formulario.' : 'Nómina generada desde el formulario.',
  });

  const mensajePrestamos = empleadosConAbono > 0 ? ` Se descontaron préstamos automáticamente a ${empleadosConAbono} empleado(s).` : '';

  return {
    ok: true,
    mensaje: `Nómina ${idNomina} generada para ${evento.nombre} con ${filasDetalle.length} empleado(s).${mensajePrestamos}`,
  };
}

function eliminarDetalleNominaPorId_(hojaDetalle, idNomina) {
  const ultimaFila = hojaDetalle.getLastRow();
  if (ultimaFila < 2) return;

  const idsNomina = hojaDetalle.getRange(2, 2, ultimaFila - 1, 1).getDisplayValues().flat();
  for (let i = idsNomina.length - 1; i >= 0; i--) {
    if (idsNomina[i] === idNomina) {
      hojaDetalle.deleteRow(i + 2);
    }
  }
}

function obtenerNominas_() {
  const hoja = obtenerHojaObligatoria_(SpreadsheetApp.getActiveSpreadsheet(), SYHME.HOJAS.NOMINAS);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return [];
  return hoja.getRange(2, 1, ultimaFila - 1, 9).getValues()
    .filter((f) => f[0] && f[1])
    .map((f, i) => ({
      fila: i + 2, id: String(f[0]), idEvento: String(f[1]), evento: String(f[2]),
      fechaGeneracion: formatearFechaIso_(f[3]), totalEmpleados: Number(f[4]) || 0,
      totalTurnos: Number(f[5]) || 0, totalNomina: Number(f[6]) || 0,
      estatus: String(f[7] || ''), observaciones: String(f[8] || ''),
    }));
}

function obtenerDetalleNomina_() {
  const hoja = obtenerHojaObligatoria_(SpreadsheetApp.getActiveSpreadsheet(), SYHME.HOJAS.DETALLE_NOMINA);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return [];
  return hoja.getRange(2, 1, ultimaFila - 1, 13).getValues()
    .filter((f) => f[0] && f[1])
    .map((f, i) => ({
      fila: i + 2, id: String(f[0]), idNomina: String(f[1]), idEvento: String(f[2]),
      idEmpleado: String(f[3]), empleado: String(f[4]), idPuesto: String(f[5]), puesto: String(f[6]),
      tarifa: Number(f[7]) || 0, totalTurnos: Number(f[8]) || 0, subtotal: Number(f[9]) || 0,
      prestamo: Number(f[10]) || 0, totalPagar: Number(f[11]) || 0, estatusPago: String(f[12] || ''),
    }));
}

function abrirCatalogoNominas() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  libro.setActiveSheet(obtenerHojaObligatoria_(libro, SYHME.HOJAS.NOMINAS));
}

function abrirCatalogoDetalleNomina() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  libro.setActiveSheet(obtenerHojaObligatoria_(libro, SYHME.HOJAS.DETALLE_NOMINA));
}

/* ======================= FIN NOMINAS ======================= */

/* ======================= PAGOS (anticipo / finiquito del cliente) ======================= */

function mostrarFormularioPagos() {
  const eventos = obtenerEventos_().filter((e) => e.estatus !== 'CANCELADO');
  if (!eventos.length) {
    SpreadsheetApp.getUi().alert('No hay eventos activos para registrar pagos.');
    return;
  }

  const datos = JSON.stringify({
    eventos,
    pagos: obtenerPagos_(),
    cotizaciones: obtenerCotizaciones_().filter((c) => c.estatus === 'ACEPTADA'),
  }).replace(/</g, '\\u003c');

  const html = HtmlService.createHtmlOutput(`
    <!doctype html><html><head><base target="_top"><style>${estilosFormulario_()}
      .resumen{background:#eceff1;border-radius:6px;padding:9px;font-size:12px;color:#37474f;margin-top:8px;display:flex;justify-content:space-between}
      .check-fila{display:flex;align-items:center;gap:8px;font-weight:400;margin-top:8px}
      .check-fila input{width:auto}
    </style></head><body>
      <h2>Pagos del cliente</h2>
      <p>El importe total se toma de la cotización aceptada del evento; puedes ajustarlo si es necesario.</p>
      <form id="formulario">
        <label for="idPago">Registro</label>
        <select id="idPago"><option value="">+ NUEVO PAGO</option></select>
        <label for="idEvento">Evento</label>
        <select id="idEvento" required></select>
        <label for="importeTotal">Importe total</label>
        <input id="importeTotal" type="number" min="0" step="0.01" required>
        <div class="grid">
          <div><label for="porcentajeAnticipo">% Anticipo</label><input id="porcentajeAnticipo" type="number" min="0" max="100" step="1" required></div>
          <div><label for="porcentajeFiniquito">% Finiquito</label><input id="porcentajeFiniquito" disabled></div>
        </div>
        <div class="resumen"><span>Anticipo: <b id="importeAnticipo">$0.00</b></span><span>Finiquito: <b id="importeFiniquito">$0.00</b></span></div>
        <label class="check-fila"><input type="checkbox" id="anticipoPagado"> Anticipo pagado</label>
        <label class="check-fila"><input type="checkbox" id="finiquitoPagado"> Finiquito pagado</label>
        <label for="observaciones">Observaciones</label><textarea id="observaciones" maxlength="300"></textarea>
        <div id="mensaje"></div>
        <div class="acciones"><button type="button" class="secundario" onclick="google.script.host.close()">Cerrar</button><button id="guardar" class="primario">Guardar pago</button></div>
      </form>
      <script>
        const datos=${datos},q=id=>document.getElementById(id),sel=q('idPago'),se=q('idEvento'),moneda=new Intl.NumberFormat('es-MX',{style:'currency',currency:'MXN'});
        datos.eventos.forEach(x=>{const o=document.createElement('option');o.value=x.id;o.textContent=x.nombre+' — '+x.inicio;se.appendChild(o)});
        datos.pagos.forEach(x=>{const o=document.createElement('option');o.value=x.id;o.textContent=x.id+' — '+x.evento;sel.appendChild(o)});

        function importeSugerido(idEvento){const c=datos.cotizaciones.find(v=>v.idEvento===idEvento);return c?c.total:0}

        function actualizarResumen(){
          const total=Number(q('importeTotal').value)||0,pctAnticipo=Number(q('porcentajeAnticipo').value)||0,pctFiniquito=100-pctAnticipo;
          q('porcentajeFiniquito').value=pctFiniquito+'%';
          q('importeAnticipo').textContent=moneda.format(total*pctAnticipo/100);
          q('importeFiniquito').textContent=moneda.format(total*pctFiniquito/100);
        }

        function cargar(){
          const x=datos.pagos.find(v=>v.id===sel.value),nuevo=!x;
          se.value=x?x.idEvento:se.value;
          q('importeTotal').value=x?x.importeTotal:importeSugerido(se.value);
          q('porcentajeAnticipo').value=x?x.porcentajeAnticipo:0;
          q('anticipoPagado').checked=x?x.anticipoPagado==='SI':false;
          q('finiquitoPagado').checked=x?x.finiquitoPagado==='SI':false;
          q('observaciones').value=x?x.observaciones:'';
          se.disabled=!nuevo;q('mensaje').textContent='';
          actualizarResumen();
        }
        sel.onchange=cargar;
        se.onchange=()=>{if(!sel.value){q('importeTotal').value=importeSugerido(se.value);actualizarResumen()}};
        ['importeTotal','porcentajeAnticipo'].forEach(id=>q(id).addEventListener('input',actualizarResumen));
        cargar();

        q('formulario').onsubmit=e=>{
          e.preventDefault();q('guardar').disabled=true;q('mensaje').className='';q('mensaje').textContent='Guardando...';
          google.script.run.withSuccessHandler(r=>{q('guardar').disabled=false;q('mensaje').className=r.ok?'ok':'error';q('mensaje').textContent=r.mensaje;if(r.ok)setTimeout(()=>google.script.host.close(),900)})
            .withFailureHandler(err=>{q('guardar').disabled=false;q('mensaje').className='error';q('mensaje').textContent=err.message||'No fue posible guardar el pago.'})
            .guardarPagoDesdeFormulario({idPago:sel.value,idEvento:se.value,importeTotal:q('importeTotal').value,porcentajeAnticipo:q('porcentajeAnticipo').value,anticipoPagado:q('anticipoPagado').checked,finiquitoPagado:q('finiquitoPagado').checked,observaciones:q('observaciones').value});
        };
      </script>
    </body></html>
  `).setWidth(480).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, 'Syhme');
}

function guardarPagoDesdeFormulario(datos) {
  const id = String((datos && datos.idPago) || '').trim();
  const idEvento = String((datos && datos.idEvento) || '').trim();
  const importeTotal = Number(datos && datos.importeTotal);
  const porcentajeAnticipo = Number(datos && datos.porcentajeAnticipo);
  const anticipoPagado = Boolean(datos && datos.anticipoPagado);
  const finiquitoPagado = Boolean(datos && datos.finiquitoPagado);
  const observaciones = String((datos && datos.observaciones) || '').trim().slice(0, 300);

  if (!idEvento) return { ok: false, mensaje: 'Selecciona un evento.' };
  if (!Number.isFinite(importeTotal) || importeTotal < 0) return { ok: false, mensaje: 'El importe total debe ser un número válido.' };
  if (!Number.isFinite(porcentajeAnticipo) || porcentajeAnticipo < 0 || porcentajeAnticipo > 100) return { ok: false, mensaje: 'El porcentaje de anticipo debe estar entre 0 y 100.' };

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) return { ok: false, mensaje: 'El sistema está ocupado. Intenta nuevamente.' };

  try {
    const libro = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = obtenerHojaObligatoria_(libro, SYHME.HOJAS.PAGOS);
    const evento = obtenerEventos_().find((e) => e.id === idEvento);
    if (!evento) return { ok: false, mensaje: 'El evento seleccionado ya no existe.' };
    const cliente = obtenerClientes_().find((c) => c.id === evento.idCliente);

    const porcentajeFiniquito = 100 - porcentajeAnticipo;
    const importeAnticipo = importeTotal * porcentajeAnticipo / 100;
    const importeFiniquito = importeTotal * porcentajeFiniquito / 100;

    const pagos = obtenerPagos_();

    if (id) {
      const actual = pagos.find((p) => p.id === id);
      if (!actual) return { ok: false, mensaje: 'El pago seleccionado ya no existe.' };
      hoja.getRange(actual.fila, 6, 1, 8).setValues([[
        importeTotal, porcentajeAnticipo, importeAnticipo, anticipoPagado ? 'SI' : 'NO',
        porcentajeFiniquito, importeFiniquito, finiquitoPagado ? 'SI' : 'NO', observaciones,
      ]]);
      registrarBitacora_({
        modulo: 'PAGOS', accion: 'MODIFICACION', idRegistro: id,
        valorAnterior: JSON.stringify({ importeTotal: actual.importeTotal, anticipoPagado: actual.anticipoPagado, finiquitoPagado: actual.finiquitoPagado }),
        valorNuevo: JSON.stringify({ importeTotal, anticipoPagado, finiquitoPagado }),
        detalle: 'Pago actualizado desde el formulario.',
      });
      return { ok: true, mensaje: `Pago ${id} actualizado correctamente.` };
    }

    const duplicado = pagos.find((p) => p.idEvento === idEvento);
    if (duplicado) return { ok: false, mensaje: `Ya existe un registro de pago (${duplicado.id}) para este evento. Selecciónalo para editarlo.` };

    const nuevoId = generarId_('PAG', hoja);
    hoja.appendRow([
      nuevoId, idEvento, evento.nombre, evento.idCliente, cliente ? cliente.nombre : '',
      importeTotal, porcentajeAnticipo, importeAnticipo, anticipoPagado ? 'SI' : 'NO',
      porcentajeFiniquito, importeFiniquito, finiquitoPagado ? 'SI' : 'NO', observaciones,
    ]);

    registrarBitacora_({
      modulo: 'PAGOS', accion: 'ALTA', idRegistro: nuevoId,
      valorNuevo: JSON.stringify({ evento: evento.nombre, importeTotal }),
      detalle: 'Pago registrado desde el formulario.',
    });

    return { ok: true, mensaje: `Pago guardado con ID ${nuevoId}.`, idPago: nuevoId };
  } finally {
    lock.releaseLock();
  }
}

function obtenerPagos_() {
  const hoja = obtenerHojaObligatoria_(SpreadsheetApp.getActiveSpreadsheet(), SYHME.HOJAS.PAGOS);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return [];
  return hoja.getRange(2, 1, ultimaFila - 1, 13).getValues()
    .filter((f) => f[0] && f[1])
    .map((f, i) => ({
      fila: i + 2, id: String(f[0]), idEvento: String(f[1]), evento: String(f[2]),
      idCliente: String(f[3]), cliente: String(f[4]), importeTotal: Number(f[5]) || 0,
      porcentajeAnticipo: Number(f[6]) || 0, importeAnticipo: Number(f[7]) || 0, anticipoPagado: String(f[8] || ''),
      porcentajeFiniquito: Number(f[9]) || 0, importeFiniquito: Number(f[10]) || 0, finiquitoPagado: String(f[11] || ''),
      observaciones: String(f[12] || ''),
    }));
}

function abrirCatalogoPagos() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  libro.setActiveSheet(obtenerHojaObligatoria_(libro, SYHME.HOJAS.PAGOS));
}

/* ======================= FIN PAGOS ======================= */

/* ======================= CONTROL_FINANCIERO (reparto de utilidades) ======================= */

function obtenerSociosAutomaticos_() {
  return obtenerEmpleadosActivos_().filter((e) => claveTexto_(e.puesto) === claveTexto_(SYHME.PUESTO_DIRECCION));
}

function mostrarFormularioControlFinanciero() {
  const eventos = obtenerEventos_().filter((e) => e.estatus !== 'CANCELADO');
  if (!eventos.length) {
    SpreadsheetApp.getUi().alert('No hay eventos activos para generar el control financiero.');
    return;
  }

  const socios = obtenerSociosAutomaticos_();
  if (socios.length < 2) {
    SpreadsheetApp.getUi().alert('Se requieren al menos dos empleados activos con puesto DIRECCIÓN para calcular el reparto.');
    return;
  }

  const datos = JSON.stringify({
    eventos,
    controles: obtenerControlFinanciero_(),
    pagos: obtenerPagos_(),
    nominas: obtenerNominas_(),
    socios: socios.slice(0, 2),
  }).replace(/</g, '\\u003c');

  const html = HtmlService.createHtmlOutput(`
    <!doctype html><html><head><base target="_top"><style>${estilosFormulario_()}
      .resumen{background:#eceff1;border-radius:6px;padding:9px;font-size:12px;color:#37474f;margin-top:8px}
      .resumen div{display:flex;justify-content:space-between;padding:2px 0}
    </style></head><body>
      <h2>Control financiero</h2>
      <p>El importe se toma del pago registrado del evento y los egresos de la nómina generada. Los socios (${socios[0].nombre} y ${socios[1].nombre}) se reparten el neto al 50/50.</p>
      <form id="formulario">
        <label for="idEvento">Evento</label>
        <select id="idEvento" required></select>
        <label for="terceroNombre">Nombre del pago a tercero (opcional)</label>
        <input id="terceroNombre" maxlength="80">
        <label for="terceroImporte">Importe del pago a tercero</label>
        <input id="terceroImporte" type="number" min="0" step="0.01">
        <div class="resumen" id="resumen"></div>
        <label for="observaciones">Observaciones</label><textarea id="observaciones" maxlength="300"></textarea>
        <div id="mensaje"></div>
        <div class="acciones"><button type="button" class="secundario" onclick="google.script.host.close()">Cerrar</button><button id="guardar" class="primario">Generar / regenerar control</button></div>
      </form>
      <script>
        const datos=${datos},q=id=>document.getElementById(id),se=q('idEvento'),moneda=new Intl.NumberFormat('es-MX',{style:'currency',currency:'MXN'});
        datos.eventos.forEach(x=>{const o=document.createElement('option');o.value=x.id;o.textContent=x.nombre+' — '+x.inicio;se.appendChild(o)});

        function calcular(){
          const idEvento=se.value;
          const pago=datos.pagos.find(p=>p.idEvento===idEvento);
          const nomina=datos.nominas.find(n=>n.idEvento===idEvento);
          const importe=pago?pago.importeTotal:0;
          const iva=importe*0.16;
          const total=importe+iva;
          const egresos=nomina?nomina.totalNomina:0;
          const ganancia=total-egresos;
          const fondo=ganancia*0.15;
          const neto=ganancia-fondo;
          const reparticion=neto/2;

          q('resumen').innerHTML=
            '<div><span>Importe</span><b>'+moneda.format(importe)+'</b></div>'+
            '<div><span>IVA (16%)</span><b>'+moneda.format(iva)+'</b></div>'+
            '<div><span>Total</span><b>'+moneda.format(total)+'</b></div>'+
            '<div><span>Egresos nómina</span><b>'+moneda.format(egresos)+'</b></div>'+
            '<div><span>Ganancia</span><b>'+moneda.format(ganancia)+'</b></div>'+
            '<div><span>Fondo utilidad (15%)</span><b>'+moneda.format(fondo)+'</b></div>'+
            '<div><span>Neto a repartir</span><b>'+moneda.format(neto)+'</b></div>'+
            '<div><span>'+datos.socios[0].nombre+'</span><b>'+moneda.format(reparticion)+'</b></div>'+
            '<div><span>'+datos.socios[1].nombre+'</span><b>'+moneda.format(reparticion)+'</b></div>'+
            (pago?'':'<div><span style="color:#b71c1c">Sin pago registrado para este evento</span></div>');
        }
        se.addEventListener('change',calcular);calcular();

        q('formulario').onsubmit=e=>{
          e.preventDefault();q('guardar').disabled=true;q('mensaje').className='';q('mensaje').textContent='Generando...';
          google.script.run.withSuccessHandler(r=>{q('guardar').disabled=false;q('mensaje').className=r.ok?'ok':'error';q('mensaje').textContent=r.mensaje})
            .withFailureHandler(err=>{q('guardar').disabled=false;q('mensaje').className='error';q('mensaje').textContent=err.message||'No fue posible generar el control financiero.'})
            .generarControlFinancieroDesdeFormulario({idEvento:se.value,terceroNombre:q('terceroNombre').value,terceroImporte:q('terceroImporte').value,observaciones:q('observaciones').value});
        };
      </script>
    </body></html>
  `).setWidth(460).setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, 'Syhme');
}

function generarControlFinancieroDesdeFormulario(datos) {
  const idEvento = String((datos && datos.idEvento) || '').trim();
  const terceroNombre = String((datos && datos.terceroNombre) || '').trim().slice(0, 80);
  const terceroImporte = Number(datos && datos.terceroImporte) || 0;
  const observaciones = String((datos && datos.observaciones) || '').trim().slice(0, 300);

  if (!idEvento) return { ok: false, mensaje: 'Selecciona un evento.' };

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(20000)) return { ok: false, mensaje: 'El sistema está ocupado. Intenta nuevamente.' };

  try {
    const evento = obtenerEventos_().find((e) => e.id === idEvento);
    if (!evento) return { ok: false, mensaje: 'El evento seleccionado ya no existe.' };

    const pago = obtenerPagos_().find((p) => p.idEvento === idEvento);
    if (!pago) return { ok: false, mensaje: 'Registra primero el pago del cliente para este evento.' };

    const nomina = obtenerNominas_().find((n) => n.idEvento === idEvento);
    if (!nomina) return { ok: false, mensaje: 'Genera primero la nómina de este evento.' };

    const socios = obtenerSociosAutomaticos_();
    if (socios.length < 2) return { ok: false, mensaje: 'Se requieren al menos dos empleados activos con puesto DIRECCIÓN.' };

    const porcentajeIva = Number(obtenerParametroConfiguracion_('IVA')) || 0.16;
    const porcentajeFondo = Number(obtenerParametroConfiguracion_('FONDO_UTILIDAD')) || 0.15;

    const importe = pago.importeTotal;
    const iva = importe * porcentajeIva;
    const total = importe + iva;
    const egresos = nomina.totalNomina;
    const ganancia = total - egresos;
    const fondo = ganancia * porcentajeFondo;
    const neto = ganancia - fondo;
    const reparticion = neto / 2;

    const libro = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = obtenerHojaObligatoria_(libro, SYHME.HOJAS.CONTROL_FINANCIERO);
    const controles = obtenerControlFinanciero_();
    const existente = controles.find((c) => c.idEvento === idEvento);

    const fila = [
      evento.nombre, importe, iva, total, egresos, ganancia, fondo, reparticion,
      socios[0].nombre, existente ? existente.socio1Estatus : 'PENDIENTE',
      socios[1].nombre, existente ? existente.socio2Estatus : 'PENDIENTE',
      terceroNombre, terceroImporte, existente ? existente.terceroEstatus : (terceroNombre ? 'PENDIENTE' : ''),
      observaciones,
    ];

    let idControl;
    if (existente) {
      idControl = existente.id;
      hoja.getRange(existente.fila, 3, 1, fila.length).setValues([fila]);
      registrarBitacora_({
        modulo: 'CONTROL_FINANCIERO', accion: 'REGENERACION', idRegistro: idControl,
        valorNuevo: JSON.stringify({ evento: evento.nombre, neto, reparticion }),
        detalle: 'Control financiero regenerado desde el formulario.',
      });
    } else {
      idControl = generarId_('CTF', hoja);
      hoja.appendRow([idControl, idEvento, ...fila]);
      registrarBitacora_({
        modulo: 'CONTROL_FINANCIERO', accion: 'ALTA', idRegistro: idControl,
        valorNuevo: JSON.stringify({ evento: evento.nombre, neto, reparticion }),
        detalle: 'Control financiero generado desde el formulario.',
      });
    }

    return {
      ok: true,
      mensaje: `Control financiero ${idControl} generado. Reparto: ${formatearMoneda_(reparticion)} por socio.`,
    };
  } finally {
    lock.releaseLock();
  }
}

function obtenerParametroConfiguracion_(nombreParametro) {
  const hoja = obtenerHojaObligatoria_(SpreadsheetApp.getActiveSpreadsheet(), SYHME.HOJAS.CONFIGURACION);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return '';
  const filas = hoja.getRange(2, 1, ultimaFila - 1, 2).getValues();
  const encontrada = filas.find((f) => String(f[0]).trim() === nombreParametro);
  return encontrada ? String(encontrada[1]).trim() : '';
}

function obtenerControlFinanciero_() {
  const hoja = obtenerHojaObligatoria_(SpreadsheetApp.getActiveSpreadsheet(), SYHME.HOJAS.CONTROL_FINANCIERO);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return [];
  return hoja.getRange(2, 1, ultimaFila - 1, 19).getValues()
    .filter((f) => f[0] && f[1])
    .map((f, i) => ({
      fila: i + 2, id: String(f[0]), idEvento: String(f[1]), evento: String(f[2]),
      importe: Number(f[3]) || 0, iva: Number(f[4]) || 0, total: Number(f[5]) || 0,
      egresosNomina: Number(f[6]) || 0, ganancia: Number(f[7]) || 0, fondoUtilidad: Number(f[8]) || 0,
      netoRepartir: Number(f[9]) || 0, reparticionPorSocio: Number(f[10]) || 0,
      socio1: String(f[11] || ''), socio1Estatus: String(f[12] || ''),
      socio2: String(f[13] || ''), socio2Estatus: String(f[14] || ''),
      terceroNombre: String(f[15] || ''), terceroImporte: Number(f[16]) || 0, terceroEstatus: String(f[17] || ''),
      observaciones: String(f[18] || ''),
    }));
}

function abrirCatalogoControlFinanciero() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  libro.setActiveSheet(obtenerHojaObligatoria_(libro, SYHME.HOJAS.CONTROL_FINANCIERO));
}

/* ======================= FIN CONTROL_FINANCIERO ======================= */

/* ======================= VISTA_NOMINA (resumen consolidado por evento) ======================= */

function estiloRangoVista_(rango, fondo, fuente, negrita, size) {
  rango.setBackground(fondo).setFontColor(fuente).setFontWeight(negrita ? 'bold' : 'normal')
    .setFontFamily('Montserrat').setFontSize(size).setVerticalAlignment('middle').setWrap(true);
}

function mostrarFormularioVistaNomina() {
  const eventos = obtenerEventos_();
  if (!eventos.length) {
    SpreadsheetApp.getUi().alert('No hay eventos registrados.');
    return;
  }

  const datos = JSON.stringify({ eventos }).replace(/</g, '\\u003c');

  const html = HtmlService.createHtmlOutput(`
    <!doctype html><html><head><base target="_top"><style>${estilosFormulario_()}</style></head><body>
      <h2>Vista Nómina</h2>
      <p>Genera el resumen consolidado de un evento: asignaciones, turnos, nómina, préstamos, pagos y reparto financiero.</p>
      <form id="formulario">
        <label for="idEvento">Evento</label>
        <select id="idEvento" required></select>
        <div id="mensaje"></div>
        <div class="acciones"><button type="button" class="secundario" onclick="google.script.host.close()">Cerrar</button><button id="guardar" class="primario">Generar resumen</button></div>
      </form>
      <script>
        const datos=${datos},q=id=>document.getElementById(id),se=q('idEvento');
        datos.eventos.forEach(x=>{const o=document.createElement('option');o.value=x.id;o.textContent=x.nombre+' — '+x.inicio;se.appendChild(o)});

        q('formulario').onsubmit=e=>{
          e.preventDefault();q('guardar').disabled=true;q('mensaje').className='';q('mensaje').textContent='Generando resumen...';
          google.script.run.withSuccessHandler(r=>{q('guardar').disabled=false;q('mensaje').className=r.ok?'ok':'error';q('mensaje').textContent=r.mensaje;if(r.ok)setTimeout(()=>google.script.host.close(),1200)})
            .withFailureHandler(err=>{q('guardar').disabled=false;q('mensaje').className='error';q('mensaje').textContent=err.message||'No fue posible generar el resumen.'})
            .generarVistaNominaDesdeFormulario({idEvento:se.value});
        };
      </script>
    </body></html>
  `).setWidth(440).setHeight(260);
  SpreadsheetApp.getUi().showModalDialog(html, 'Syhme');
}

function generarVistaNominaDesdeFormulario(datos) {
  const idEvento = String((datos && datos.idEvento) || '').trim();
  if (!idEvento) return { ok: false, mensaje: 'Selecciona un evento.' };

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(20000)) return { ok: false, mensaje: 'El sistema está ocupado. Intenta nuevamente.' };

  try {
    return generarVistaNominaEvento_(idEvento);
  } finally {
    lock.releaseLock();
  }
}

function generarVistaNominaEvento_(idEvento) {
  const evento = obtenerEventos_().find((e) => e.id === idEvento);
  if (!evento) return { ok: false, mensaje: 'El evento seleccionado ya no existe.' };

  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = obtenerHojaObligatoria_(libro, SYHME.HOJAS.VISTA_NOMINA);
  hoja.clear();
  hoja.clearFormats();

  const cliente = obtenerClientes_().find((c) => c.id === evento.idCliente);
  const nomina = obtenerNominas_().find((n) => n.idEvento === idEvento);
  const detalle = nomina ? obtenerDetalleNomina_().filter((d) => d.idNomina === nomina.id) : [];
  const pago = obtenerPagos_().find((p) => p.idEvento === idEvento);
  const control = obtenerControlFinanciero_().find((c) => c.idEvento === idEvento);
  const movimientosNomina = nomina ? obtenerMovimientosPrestamo_().filter((m) => m.idNomina === nomina.id) : [];

  let fila = 1;
  const anchoTotal = 8;

  // Título
  hoja.getRange(fila, 1, 1, anchoTotal).merge().setValue(`RESUMEN DE EVENTO — ${evento.nombre}`);
  estiloRangoVista_(hoja.getRange(fila, 1, 1, anchoTotal), '#2d516a', '#ffffff', true, 14);
  fila += 2;

  // Datos generales
  const filaInicio = fila;
  hoja.getRange(fila, 1).setValue('Cliente:'); hoja.getRange(fila, 2, 1, 3).merge().setValue(cliente ? cliente.nombre : '—');
  hoja.getRange(fila, 5).setValue('Estatus:'); hoja.getRange(fila, 6, 1, 3).merge().setValue(evento.estatus);
  fila++;
  hoja.getRange(fila, 1).setValue('Fecha inicio:'); hoja.getRange(fila, 2).setValue(evento.inicio);
  hoja.getRange(fila, 3).setValue('Fecha fin:'); hoja.getRange(fila, 4).setValue(evento.fin);
  fila++;
  estiloRangoVista_(hoja.getRange(filaInicio, 1, fila - filaInicio, anchoTotal), '#f0f3f4', '#000000', false, 10);
  hoja.getRange(filaInicio, 1, fila - filaInicio, 1).setFontWeight('bold');
  hoja.getRange(filaInicio, 3, fila - filaInicio, 1).setFontWeight('bold');
  hoja.getRange(filaInicio, 5, fila - filaInicio, 1).setFontWeight('bold');
  fila += 1;

  // Sección Nómina
  hoja.getRange(fila, 1, 1, anchoTotal).merge().setValue('DETALLE DE NÓMINA');
  estiloRangoVista_(hoja.getRange(fila, 1, 1, anchoTotal), '#424949', '#ffffff', true, 12);
  fila++;

  if (!nomina || !detalle.length) {
    hoja.getRange(fila, 1, 1, anchoTotal).merge().setValue('Aún no se ha generado la nómina de este evento.');
    fila += 2;
  } else {
    const encabezados = [['Empleado', 'Puesto', 'Tarifa', 'Turnos', 'Subtotal', 'Préstamo', 'Total a pagar', 'Estatus']];
    const rangoEncabezados = hoja.getRange(fila, 1, 1, anchoTotal);
    rangoEncabezados.setValues(encabezados);
    estiloRangoVista_(rangoEncabezados, '#666666', '#ffffff', true, 9);
    fila++;

    const filaDatosInicio = fila;
    detalle.forEach((d) => {
      hoja.getRange(fila, 1, 1, anchoTotal).setValues([[d.empleado, d.puesto, d.tarifa, d.totalTurnos, d.subtotal, d.prestamo, d.totalPagar, d.estatusPago]]);
      fila++;
    });
    estiloRangoVista_(hoja.getRange(filaDatosInicio, 1, detalle.length, anchoTotal), '#ffffff', '#000000', false, 9);
    hoja.getRange(filaDatosInicio, 3, detalle.length, 4).setNumberFormat('_-"$"* #,##0.00_-;_-"$"* \\-#,##0.00_-;_-"$"* "-"??_-;_-@');

    hoja.getRange(fila, 1, 1, 6).merge().setValue('Total nómina');
    hoja.getRange(fila, 7).setValue(`=SUM(G${filaDatosInicio}:G${fila - 1})`);
    hoja.getRange(fila, 7).setNumberFormat('_-"$"* #,##0.00_-;_-"$"* \\-#,##0.00_-;_-"$"* "-"??_-;_-@');
    estiloRangoVista_(hoja.getRange(fila, 1, 1, anchoTotal), '#424949', '#ffffff', true, 9);
    fila += 2;
  }

  // Sección Préstamos abonados en esta nómina
  if (movimientosNomina.length) {
    hoja.getRange(fila, 1, 1, anchoTotal).merge().setValue('PRÉSTAMOS ABONADOS EN ESTA NÓMINA');
    estiloRangoVista_(hoja.getRange(fila, 1, 1, anchoTotal), '#424949', '#ffffff', true, 12);
    fila++;

    hoja.getRange(fila, 1, 1, 4).setValues([['Empleado', 'Importe abonado', 'Saldo resultante', '']]);
    estiloRangoVista_(hoja.getRange(fila, 1, 1, 4), '#666666', '#ffffff', true, 9);
    fila++;

    movimientosNomina.forEach((m) => {
      hoja.getRange(fila, 1, 1, 3).setValues([[m.empleado, m.importe, m.saldoResultante]]);
      fila++;
    });
    fila += 1;
  }

  // Sección Pago del cliente
  hoja.getRange(fila, 1, 1, anchoTotal).merge().setValue('PAGO DEL CLIENTE');
  estiloRangoVista_(hoja.getRange(fila, 1, 1, anchoTotal), '#424949', '#ffffff', true, 12);
  fila++;

  if (!pago) {
    hoja.getRange(fila, 1, 1, anchoTotal).merge().setValue('Sin pago registrado para este evento.');
    fila += 2;
  } else {
    const filaPagoInicio = fila;
    hoja.getRange(fila, 1).setValue('Importe total:'); hoja.getRange(fila, 2).setValue(pago.importeTotal);
    hoja.getRange(fila, 3).setValue('Anticipo (' + pago.porcentajeAnticipo + '%):'); hoja.getRange(fila, 4).setValue(pago.importeAnticipo);
    hoja.getRange(fila, 5).setValue('Pagado:'); hoja.getRange(fila, 6).setValue(pago.anticipoPagado);
    fila++;
    hoja.getRange(fila, 3).setValue('Finiquito (' + pago.porcentajeFiniquito + '%):'); hoja.getRange(fila, 4).setValue(pago.importeFiniquito);
    hoja.getRange(fila, 5).setValue('Pagado:'); hoja.getRange(fila, 6).setValue(pago.finiquitoPagado);
    fila++;
    estiloRangoVista_(hoja.getRange(filaPagoInicio, 1, fila - filaPagoInicio, anchoTotal), '#f0f3f4', '#000000', false, 10);
    fila += 1;
  }

  // Sección Control financiero
  hoja.getRange(fila, 1, 1, anchoTotal).merge().setValue('REPARTO FINANCIERO');
  estiloRangoVista_(hoja.getRange(fila, 1, 1, anchoTotal), '#424949', '#ffffff', true, 12);
  fila++;

  if (!control) {
    hoja.getRange(fila, 1, 1, anchoTotal).merge().setValue('Aún no se ha generado el control financiero de este evento.');
    fila += 2;
  } else {
    const filaControlInicio = fila;
    hoja.getRange(fila, 1).setValue('Importe:'); hoja.getRange(fila, 2).setValue(control.importe);
    hoja.getRange(fila, 3).setValue('IVA:'); hoja.getRange(fila, 4).setValue(control.iva);
    hoja.getRange(fila, 5).setValue('Total:'); hoja.getRange(fila, 6).setValue(control.total);
    fila++;
    hoja.getRange(fila, 1).setValue('Egresos nómina:'); hoja.getRange(fila, 2).setValue(control.egresosNomina);
    hoja.getRange(fila, 3).setValue('Ganancia:'); hoja.getRange(fila, 4).setValue(control.ganancia);
    hoja.getRange(fila, 5).setValue('Fondo utilidad:'); hoja.getRange(fila, 6).setValue(control.fondoUtilidad);
    fila++;
    hoja.getRange(fila, 1).setValue('Neto a repartir:'); hoja.getRange(fila, 2).setValue(control.netoRepartir);
    fila++;
    hoja.getRange(fila, 1).setValue(control.socio1 + ':'); hoja.getRange(fila, 2).setValue(control.reparticionPorSocio);
    hoja.getRange(fila, 3).setValue('Estatus:'); hoja.getRange(fila, 4).setValue(control.socio1Estatus);
    fila++;
    hoja.getRange(fila, 1).setValue(control.socio2 + ':'); hoja.getRange(fila, 2).setValue(control.reparticionPorSocio);
    hoja.getRange(fila, 3).setValue('Estatus:'); hoja.getRange(fila, 4).setValue(control.socio2Estatus);
    fila++;
    if (control.terceroNombre) {
      hoja.getRange(fila, 1).setValue(control.terceroNombre + ':'); hoja.getRange(fila, 2).setValue(control.terceroImporte);
      hoja.getRange(fila, 3).setValue('Estatus:'); hoja.getRange(fila, 4).setValue(control.terceroEstatus);
      fila++;
    }
    estiloRangoVista_(hoja.getRange(filaControlInicio, 1, fila - filaControlInicio, anchoTotal), '#f0f3f4', '#000000', false, 10);
    hoja.getRange(filaControlInicio, 1, fila - filaControlInicio, anchoTotal).getCell(1, 1);
  }

  hoja.setColumnWidths(1, 8, 110);
  hoja.setFrozenRows(1);
  libro.setActiveSheet(hoja);

  registrarBitacora_({
    modulo: 'VISTA_NOMINA', accion: 'GENERACION', idRegistro: idEvento,
    valorNuevo: JSON.stringify({ evento: evento.nombre }),
    detalle: 'Vista consolidada generada desde el formulario.',
  });

  return { ok: true, mensaje: `Resumen de ${evento.nombre} generado correctamente.` };
}

/* ======================= FIN VISTA_NOMINA ======================= */

function verificarEstructura() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const requeridas = Object.values(SYHME.HOJAS);
  const faltantes = requeridas.filter((nombre) => !libro.getSheetByName(nombre));

  if (faltantes.length) {
    SpreadsheetApp.getUi().alert(`Faltan las hojas: ${faltantes.join(', ')}`);
    return false;
  }

  SpreadsheetApp.getActive().toast('La estructura base está completa.', 'Syhme', 5);
  return true;
}

function obtenerPuestosActivos_() {
  return obtenerPuestos_()
    .filter((puesto) => puesto.estatus === SYHME.ESTATUS_ACTIVO)
    .map((puesto) => ({ id: puesto.id, nombre: puesto.nombre, tarifa: puesto.tarifa }));
}

function obtenerPuestoActivoPorId_(idPuesto) {
  return obtenerPuestosActivos_().find((puesto) => puesto.id === idPuesto) || null;
}

function buscarEmpleadoPorNombre_(hoja, nombre) {
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return null;

  const objetivo = claveTexto_(nombre);
  const valores = hoja.getRange(2, 1, ultimaFila - 1, 2).getDisplayValues();
  const encontrado = valores.find((fila) => claveTexto_(fila[1]) === objetivo);
  return encontrado ? { id: encontrado[0], nombre: encontrado[1] } : null;
}

function generarId_(prefijo, hoja) {
  const ultimaFila = hoja.getLastRow();
  let maximo = 0;

  if (ultimaFila >= 2) {
    const expresion = new RegExp(`^${prefijo}-(\\d+)$`);
    hoja.getRange(2, 1, ultimaFila - 1, 1).getDisplayValues().flat().forEach((id) => {
      const coincidencia = String(id).match(expresion);
      if (coincidencia) maximo = Math.max(maximo, Number(coincidencia[1]));
    });
  }

  return `${prefijo}-${String(maximo + 1).padStart(5, '0')}`;
}

function registrarBitacora_({ modulo, accion, idRegistro, valorAnterior = '', valorNuevo = '', detalle = '' }) {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = obtenerHojaObligatoria_(libro, SYHME.HOJAS.BITACORA);
  const idBitacora = generarId_('BIT', hoja);
  const usuario = Session.getActiveUser().getEmail() || 'USUARIO';

  hoja.appendRow([
    idBitacora,
    new Date(),
    usuario,
    modulo,
    accion,
    idRegistro,
    valorAnterior,
    valorNuevo,
    detalle,
  ]);
}

function obtenerHojaObligatoria_(libro, nombre) {
  const hoja = libro.getSheetByName(nombre);
  if (!hoja) throw new Error(`No existe la hoja obligatoria ${nombre}.`);
  return hoja;
}

function normalizarNombre_(valor) {
  return String(valor || '')
    .trim()
    .replace(/\s+/g, ' ')
    .toLocaleUpperCase('es-MX');
}

function claveTexto_(valor) {
  return normalizarNombre_(valor)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}