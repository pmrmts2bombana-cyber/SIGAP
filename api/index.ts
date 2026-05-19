import express from "express";
import path from "path";
import fs from "fs";
import { fileURLToPath } from "url";
import { GoogleGenAI } from "@google/genai";

// Universal __dirname and __filename implementation
let _dirname = process.cwd();
try {
  if (typeof __dirname !== 'undefined') {
    _dirname = __dirname;
  } else if (import.meta && import.meta.url) {
    const __filename = fileURLToPath(import.meta.url);
    _dirname = path.dirname(__filename);
  }
} catch (e) {
  // Safe fallback
}

async function startServer() {
  const app = express();
  const PORT = 3000;
  const isProd = process.env.NODE_ENV === "production" || fs.existsSync(path.join(process.cwd(), "dist", "server.cjs"));

  app.use(express.json());

  // Initialize Gemini (Safe initialization)
  let ai: any = null;
  const getAI = () => {
    if (!ai && process.env.GEMINI_API_KEY) {
      try {
        ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY as string });
      } catch (e) {
        console.error("Failed to initialize Gemini AI:", e);
      }
    }
    return ai;
  };

  const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbyBPczFq0evqKAWc9G-NN3YJeEsbllnkeHfOzrlRPoMfI8aucROYj3gYrYGEZS3QSjrrA/exec";

  // API Routes
  app.post("/api/action", async (req, res) => {
    try {
      const response = await fetch(WEB_APP_URL, {
        method: "POST",
        body: JSON.stringify(req.body),
      });
      const data = await response.json();
      res.json(data);
    } catch (err) {
      console.error("Proxy Error:", err);
      res.status(500).json({ error: "Failed to connect to Google Apps Script" });
    }
  });

  app.get("/api/health", (req, res) => {
    res.json({ 
      status: "ok", 
      mode: process.env.NODE_ENV, 
      isProd,
      cwd: process.cwd(),
      _dirname
    });
  });

  if (!isProd) {
    // Vite middleware for development
    console.log("[Server] Running in DEVELOPMENT mode with Vite middleware");
    const { createServer } = await import("vite");
    const vite = await createServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    // Production static file serving
    console.log("[Server] Running in PRODUCTION mode");
    
    // In production (Cloud Run), files are typically in /workspace/dist because of the build script
    let distPath = path.resolve(process.cwd(), "dist");
    
    // Check possible locations
    const candidates = [
      distPath,
      _dirname,
      path.resolve(_dirname, "..", "dist")
    ];
    
    for (const c of candidates) {
      if (fs.existsSync(path.join(c, "index.html"))) {
        distPath = c;
        break;
      }
    }
    
    console.log(`[Server] Serving static files from: ${distPath}`);
    app.use(express.static(distPath, { index: false }));

    // Catch-all route for SPA fallback
    app.get("*", (req, res, next) => {
      // Skip API
      if (req.path.startsWith("/api/")) return next();
      
      // If request looks like an asset that wasn't caught by express.static
      if (path.extname(req.path)) {
        return res.status(404).send("Asset not found");
      }

      const indexPath = path.join(distPath, "index.html");
      if (fs.existsSync(indexPath)) {
        res.sendFile(indexPath);
      } else {
        res.status(404).send("Application files not found. Please ensure the build completed successfully.");
      }
    });
  }

  // Only listen if not running on Vercel as a function
  if (!process.env.VERCEL) {
    app.listen(PORT, "0.0.0.0", () => {
      console.log(`Server running on http://localhost:${PORT}`);
    });
  }

  return app;
}

const appPromise = startServer();

export default appPromise;
