const express = require('express');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();

// Cambiar el puerto a uno dinámico proporcionado por el entorno (usando `process.env.PORT`)
// Si no está definido, usa el puerto 3000 (para el entorno local)
const port = process.env.PORT || 3000;

const excelPath = path.join(__dirname, 'datos-asistencia.xlsx');

// Crear archivo Excel si no existe
if (!fs.existsSync(excelPath)) {
    const nuevoLibro = xlsx.utils.book_new();
    const hojaVacia = xlsx.utils.json_to_sheet([]);
    xlsx.utils.book_append_sheet(nuevoLibro, hojaVacia, 'Asistentes');
    xlsx.writeFile(nuevoLibro, excelPath);
}

// Middleware para leer formularios
app.use(express.urlencoded({ extended: true }));

// **MODIFICACIÓN**: Servir archivos estáticos desde la carpeta 'public' completa
app.use(express.static(path.join(__dirname, 'public')));

// Ruta raíz que sirve el formulario HTML
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'sitio', 'formulario.html'));
});

// Función para guardar los datos en el Excel
function guardarRespuesta(nombre, asistencia, deseo) {
    const libro = xlsx.readFile(excelPath);
    const hoja = libro.Sheets['Asistentes'];
    const datos = xlsx.utils.sheet_to_json(hoja);

    datos.push({
        nombre,
        asistencia,
        deseo,
        fecha: new Date().toLocaleString()
    });

    const nuevaHoja = xlsx.utils.json_to_sheet(datos);
    libro.Sheets['Asistentes'] = nuevaHoja;

    try {
        xlsx.writeFile(libro, excelPath);
    } catch (error) {
        console.error('❌ Error al guardar en Excel:', error.message);
    }
}

// Ruta que recibe los datos del formulario
app.post('/enviar', (req, res) => {
    const { nombre, asistencia, deseo } = req.body;

    // Guardamos la respuesta en el Excel
    guardarRespuesta(nombre, asistencia, deseo);

    res.send(`
        <h2>Gracias, ${nombre}.</h2>
        <p>Tu respuesta fue: <strong>${asistencia}</strong></p>
        <p>Tu deseo: <em>"${deseo}"</em></p>
        <a href="/">Volver al formulario</a>
    `);
});

// Iniciar servidor en el puerto asignado
app.listen(port, () => {
    console.log(`Servidor activo en http://localhost:${port}`);
});
