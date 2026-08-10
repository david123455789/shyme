const SYHME = Object.freeze({
  HOJAS: Object.freeze({
    CONFIGURACION: 'CONFIGURACION',
    PUESTOS: 'PUESTOS',
    EMPLEADOS: 'EMPLEADOS',
    CLIENTES: 'CLIENTES',
    EVENTOS: 'EVENTOS',
    COTIZACIONES: 'COTIZACIONES',
    BITACORA: 'BITACORA',
  }),
  ESTATUS_ACTIVO: 'ACTIVO',
});

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Syhme')
    .addItem('Dar de alta empleado', 'mostrarFormularioAltaEmpleado')
    .addItem('Abrir catálogo de empleados', 'abrirCatalogoEmpleados')
    .addSeparator()
    .addItem('Administrar puestos y tarifas', 'mostrarFormularioPuestos')
    .addItem('Abrir catálogo de puestos', 'abrirCatalogoPuestos')
    .addSeparator()
    .addItem('Administrar clientes', 'mostrarFormularioClientes')
    .addItem('Administrar eventos', 'mostrarFormularioEventos')
    .addItem('Abrir catálogo de clientes', 'abrirCatalogoClientes')
    .addItem('Abrir catálogo de eventos', 'abrirCatalogoEventos')
    .addSeparator()
    .addItem('Administrar cotizaciones', 'mostrarFormularioCotizaciones')
    .addItem('Abrir catálogo de cotizaciones', 'abrirCatalogoCotizaciones')
    .addSeparator()
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

function mostrarFormularioEventos() {
  const eventosJson=JSON.stringify(obtenerEventos_()).replace(/</g,'\\u003c'),clientesJson=JSON.stringify(obtenerClientes_().filter(x=>x.estatus==='ACTIVO')).replace(/</g,'\\u003c');
  if(clientesJson==='[]'){SpreadsheetApp.getUi().alert('Registra al menos un cliente activo antes de crear eventos.');return}
  const html=HtmlService.createHtmlOutput(`<!doctype html><html><head><base target="_top"><style>${estilosFormulario_()}</style></head><body><h2>Eventos</h2><p>Crea un evento o selecciona uno para actualizarlo.</p><form id="formulario">
  <label for="idEvento">Registro</label><select id="idEvento"><option value="">+ NUEVO EVENTO</option></select><label for="nombre">Nombre del evento</label><input id="nombre" maxlength="120" required>
  <label for="idCliente">Cliente</label><select id="idCliente" required></select><div class="grid"><div><label for="inicio">Fecha de inicio</label><input id="inicio" type="date" required></div><div><label for="fin">Fecha de fin</label><input id="fin" type="date" required></div></div>
  <label for="estatus">Estatus</label><select id="estatus"><option>BORRADOR</option><option>CONFIRMADO</option><option>EN CURSO</option><option>FINALIZADO</option><option>CANCELADO</option></select>
  <label for="observaciones">Observaciones</label><textarea id="observaciones" maxlength="300"></textarea><div id="mensaje"></div><div class="acciones"><button type="button" class="secundario" onclick="google.script.host.close()">Cancelar</button><button id="guardar" class="primario">Guardar evento</button></div></form>
  <script>const eventos=${eventosJson},clientes=${clientesJson},q=id=>document.getElementById(id),sel=q('idEvento'),sc=q('idCliente');clientes.forEach(x=>{const o=document.createElement('option');o.value=x.id;o.textContent=x.nombre;sc.appendChild(o)});eventos.forEach(x=>{const o=document.createElement('option');o.value=x.id;o.textContent=x.nombre+' — '+x.inicio;sel.appendChild(o)});
  function cargar(){const x=eventos.find(v=>v.id===sel.value);q('nombre').value=x?x.nombre:'';if(x&&![...sc.options].some(o=>o.value===x.idCliente)){const o=document.createElement('option');o.value=x.idCliente;o.textContent=x.cliente+' (INACTIVO)';sc.appendChild(o)}if(x)sc.value=x.idCliente;q('inicio').value=x?x.inicio:'';q('fin').value=x?x.fin:'';q('estatus').value=x?x.estatus:'BORRADOR';q('observaciones').value=x?x.observaciones:'';q('mensaje').textContent=''}sel.onchange=cargar;cargar();
  q('formulario').onsubmit=e=>{e.preventDefault();q('guardar').disabled=true;q('mensaje').textContent='Guardando...';google.script.run.withSuccessHandler(r=>{q('guardar').disabled=false;q('mensaje').className=r.ok?'ok':'error';q('mensaje').textContent=r.mensaje;if(r.ok)setTimeout(()=>google.script.host.close(),900)}).withFailureHandler(e=>{q('guardar').disabled=false;q('mensaje').className='error';q('mensaje').textContent=e.message||'No fue posible guardar.'}).guardarEventoDesdeFormulario({idEvento:sel.value,nombre:q('nombre').value,idCliente:sc.value,inicio:q('inicio').value,fin:q('fin').value,estatus:q('estatus').value,observaciones:q('observaciones').value})};</script></body></html>`).setWidth(540).setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html,'Syhme');
}

function guardarEventoDesdeFormulario(datos) {
  const id=String((datos&&datos.idEvento)||'').trim(),nombre=normalizarNombre_(datos&&datos.nombre),idCliente=String((datos&&datos.idCliente)||'').trim(),inicio=fechaDesdeIso_(datos&&datos.inicio),fin=fechaDesdeIso_(datos&&datos.fin),estatus=String((datos&&datos.estatus)||'').trim().toUpperCase(),observaciones=String((datos&&datos.observaciones)||'').trim().slice(0,300);
  if(!nombre)return{ok:false,mensaje:'Escribe el nombre del evento.'};if(!idCliente)return{ok:false,mensaje:'Selecciona un cliente.'};if(!inicio||!fin)return{ok:false,mensaje:'Captura fechas válidas.'};if(fin<inicio)return{ok:false,mensaje:'La fecha final no puede ser anterior a la inicial.'};
  const estados=['BORRADOR','CONFIRMADO','EN CURSO','FINALIZADO','CANCELADO'];if(!estados.includes(estatus))return{ok:false,mensaje:'Selecciona un estatus válido.'};
  const lock=LockService.getDocumentLock();if(!lock.tryLock(15000))return{ok:false,mensaje:'El sistema está ocupado. Intenta nuevamente.'};
  try{const libro=SpreadsheetApp.getActiveSpreadsheet(),hoja=obtenerHojaObligatoria_(libro,SYHME.HOJAS.EVENTOS),eventos=obtenerEventos_(),cliente=obtenerClientes_().find(x=>x.id===idCliente);if(!cliente)return{ok:false,mensaje:'El cliente seleccionado ya no existe.'};if(!id&&cliente.estatus!=='ACTIVO')return{ok:false,mensaje:'El cliente seleccionado no está activo.'};
    const isoInicio=formatearFechaIso_(inicio),duplicado=eventos.find(x=>claveTexto_(x.nombre)===claveTexto_(nombre)&&x.inicio===isoInicio&&x.id!==id);if(duplicado)return{ok:false,mensaje:`Ya existe ${duplicado.nombre} con esa fecha de inicio.`};const dias=Math.floor((fin-inicio)/86400000)+1;
    if(id){const actual=eventos.find(x=>x.id===id);if(!actual)return{ok:false,mensaje:'El evento seleccionado ya no existe.'};hoja.getRange(actual.fila,2,1,10).setValues([[nombre,idCliente,cliente.nombre,inicio,fin,dias,estatus,actual.idCarpeta,actual.urlCarpeta,observaciones]]);registrarBitacora_({modulo:'EVENTOS',accion:'MODIFICACION',idRegistro:id,valorAnterior:JSON.stringify({nombre:actual.nombre,cliente:actual.cliente,inicio:actual.inicio,fin:actual.fin,estatus:actual.estatus}),valorNuevo:JSON.stringify({nombre,cliente:cliente.nombre,inicio:isoInicio,fin:formatearFechaIso_(fin),estatus}),detalle:'Evento actualizado desde el formulario.'});return{ok:true,mensaje:`Evento ${id} actualizado correctamente.`}}
    const nuevoId=generarId_('EVE',hoja);hoja.appendRow([nuevoId,nombre,idCliente,cliente.nombre,inicio,fin,dias,estatus,'','',observaciones]);registrarBitacora_({modulo:'EVENTOS',accion:'ALTA',idRegistro:nuevoId,valorNuevo:JSON.stringify({nombre,cliente:cliente.nombre,inicio:isoInicio,fin:formatearFechaIso_(fin),estatus}),detalle:'Evento registrado desde el formulario.'});return{ok:true,mensaje:`Evento guardado con ID ${nuevoId}.`};
  }finally{lock.releaseLock()}
}

function obtenerClientes_(){const hoja=obtenerHojaObligatoria_(SpreadsheetApp.getActiveSpreadsheet(),SYHME.HOJAS.CLIENTES),n=hoja.getLastRow();if(n<2)return[];return hoja.getRange(2,1,n-1,7).getValues().filter(f=>f[0]&&f[1]).map((f,i)=>({fila:i+2,id:String(f[0]),nombre:String(f[1]),correo:String(f[2]||''),estatus:String(f[3]),fechaAlta:f[4],observaciones:String(f[6]||'')}))}
function obtenerEventos_(){const hoja=obtenerHojaObligatoria_(SpreadsheetApp.getActiveSpreadsheet(),SYHME.HOJAS.EVENTOS),n=hoja.getLastRow();if(n<2)return[];return hoja.getRange(2,1,n-1,11).getValues().filter(f=>f[0]&&f[1]).map((f,i)=>({fila:i+2,id:String(f[0]),nombre:String(f[1]),idCliente:String(f[2]),cliente:String(f[3]),inicio:formatearFechaIso_(f[4]),fin:formatearFechaIso_(f[5]),estatus:String(f[7]),idCarpeta:String(f[8]||''),urlCarpeta:String(f[9]||''),observaciones:String(f[10]||'')}))}
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
