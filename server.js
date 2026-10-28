const http = require('http');
const fs = require('fs');
const path = require('path');
const { exec } = require('child_process');

let PORT = 3000;
const MIME_TYPES = {
  '.html': 'text/html',
  '.css': 'text/css',
  '.js': 'text/javascript',
  '.json': 'application/json',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.svg': 'image/svg+xml',
  '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
};

const server = http.createServer((req, res) => {
  let filePath = path.join(__dirname, req.url === '/' ? 'index.html' : req.url.split('?')[0]);
  const ext = path.extname(filePath).toLowerCase();
  const contentType = MIME_TYPES[ext] || 'application/octet-stream';

  fs.readFile(filePath, (err, content) => {
    if (err) {
      if (err.code === 'ENOENT') {
        res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
        res.end('404 Archivo no encontrado');
      } else {
        res.writeHead(500);
        res.end(`Error del servidor: ${err.code}`);
      }
    } else {
      res.writeHead(200, { 'Content-Type': contentType });
      res.end(content, 'utf-8');
    }
  });
});

function startServer(portToTry) {
  server.listen(portToTry, () => {
    const url = `http://localhost:${portToTry}/`;
    console.log(`\n======================================================`);
    console.log(`🚀 ¡Presentación iniciada con éxito!`);
    console.log(`👉 Abre en tu navegador: ${url}`);
    console.log(`Presiona Ctrl + C en la terminal para detener el servidor.`);
    console.log(`======================================================\n`);

    // Abre automáticamente en el navegador predeterminado de Windows
    const startCmd = process.platform === 'win32' ? `start ${url}` :
                     process.platform === 'darwin' ? `open ${url}` : `xdg-open ${url}`;
    exec(startCmd, (err) => {
      if (err) console.log(`(Nota: Abre manualmente ${url} en tu navegador)`);
    });
  });

  server.on('error', (e) => {
    if (e.code === 'EADDRINUSE') {
      console.log(`Puerto ${portToTry} ocupado, intentando en el puerto ${portToTry + 1}...`);
      startServer(portToTry + 1);
    } else {
      console.error("Error al iniciar el servidor:", e);
    }
  });
}

startServer(PORT);
