const https = require("https");
const fs = require("fs");
const path = require("path");
const devCerts = require("office-addin-dev-certs");
const url = require("url");

const port = 3000;
const distPath = path.resolve(__dirname, "dist");

async function startServer() {
    try {
        const options = await devCerts.getHttpsServerOptions();
        
        https.createServer(options, (req, res) => {
            const parsedUrl = url.parse(req.url);
            let pathname = parsedUrl.pathname;

            // Default to taskpane.html
            if (pathname === "/" || pathname === "" || pathname === "/taskpane.html") {
                pathname = "taskpane.html";
            }

            // Remove leading slash for path.join
            const cleanPath = pathname.startsWith("/") ? pathname.slice(1) : pathname;
            const filePath = path.join(distPath, cleanPath);

            const extname = path.extname(filePath).toLowerCase();
            const mimeTypes = {
                ".html": "text/html",
                ".js": "text/javascript",
                ".css": "text/css",
                ".json": "application/json",
                ".png": "image/png",
                ".jpg": "image/jpg",
                ".ico": "image/x-icon",
            };

            const contentType = mimeTypes[extname] || "application/octet-stream";

            fs.readFile(filePath, (error, content) => {
                if (error) {
                    console.error(`[404] ${req.url} -> Not found at: ${filePath}`);
                    res.writeHead(404);
                    res.end("File not found");
                } else {
                    res.writeHead(200, { 
                        "Content-Type": contentType,
                        "Access-Control-Allow-Origin": "*" 
                    });
                    res.end(content, "utf-8");
                }
            });
        }).listen(port);

        console.log(`\x1b[32mSUCCESS:\x1b[0m Server running at https://localhost:${port}`);
        console.log(`Serving files from: ${distPath}`);
    } catch (err) {
        console.error("Failed to start server:", err);
    }
}

startServer();
