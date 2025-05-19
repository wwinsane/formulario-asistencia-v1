const express = require('express');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();

// Cambiar el puerto a uno dinámico proporcionado por el entorno (usando `process.env.PORT`)
// Si no está definido, usa el puerto 3000 (para el entorno local)
const port = process.env.PORT || 3000;

// Usar path.resolve() para garantizar la ruta correcta
const excelPath = path.resolve(__dirname, 'datos-asistencia.xlsx');

// Crear archivo Excel si no existe
if (!fs.existsSync(excelPath)) {
    const nuevoLibro = xlsx.utils.book_new();
    const hojaVacia = xlsx.utils.json_to_sheet([]);
    xlsx.utils.book_append_sheet(nuevoLibro, hojaVacia, 'Asistentes');
    xlsx.writeFile(nuevoLibro, excelPath);
}

// Middleware para leer formularios
app.use(express.urlencoded({ extended: true }));

// Servir archivos estáticos desde la carpeta 'public' completa
app.use(express.static(path.join(__dirname, 'public')));

// Ruta raíz que sirve el formulario HTML
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'sitio', 'formulario.html'));
});

// Función para guardar los datos en el Excel
function guardarRespuesta(nombre, asistencia, deseo) {
    console.log("Leyendo archivo Excel:", excelPath);
    const libro = xlsx.readFile(excelPath);
    const hoja = libro.Sheets['Asistentes'] || {}; // Asegúrate de que la hoja existe
    const datos = xlsx.utils.sheet_to_json(hoja);

    console.log("Datos antes de agregar:", datos);

    // Insertar nueva respuesta
    datos.push({
        nombre,
        asistencia,
        deseo,
        fecha: new Date().toLocaleString()
    });

    // Convertir los datos a la hoja de Excel
    const nuevaHoja = xlsx.utils.json_to_sheet(datos);
    libro.Sheets['Asistentes'] = nuevaHoja;

    try {
        console.log("Guardando datos en Excel...");
        xlsx.writeFile(libro, excelPath);
        console.log("Datos guardados correctamente.");
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
