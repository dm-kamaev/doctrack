'use strict';

const http = require('http');
const server = http.createServer();

server.on('request', (req, res) => {
    console.log('Request => ', req.url);
    // Set the content type to PNG
    res.writeHead(200, { 'Content-Type': 'image/png' });
    res.end('OK\n');
}).listen(5001);