import { fileURLToPath, URL } from "node:url";
import fs from "fs";
import path from "path";
import { homedir } from "os";

import { defineConfig } from "vite";
import vue from "@vitejs/plugin-vue";
import { viteStaticCopy } from "vite-plugin-static-copy";

// Configuration
const isDev = true;
const devPort = 8989;
const prodUrl = "https://your.addin.production.url/";

// Calculated
const _homeDir = homedir();
const _urlDefault = `https://localhost:3000/`;
const _urlDev = `https://localhost:${devPort}/`;
const _urlProd = `${prodUrl}`;

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [
    vue(),
    viteStaticCopy({
      targets: [
        {
          src: "manifest.xml",
          dest: "",
          transform(content) {
            return content
              .toString()
              .replace(
                new RegExp(_urlDefault, "g"),
                isDev ? _urlDev : _urlProd
              );
          },
        },
      ],
    }),
  ],
  resolve: {
    alias: {
      "@": fileURLToPath(new URL("./src", import.meta.url)),
    },
  },
  server: {
    host: true,
    port: devPort,
    https: {
      key: fs.readFileSync(
        path.resolve(`${_homeDir}/.office-addin-dev-certs/localhost.key`)
      ),
      cert: fs.readFileSync(
        path.resolve(`${_homeDir}/.office-addin-dev-certs/localhost.crt`)
      ),
      ca: fs.readFileSync(
        path.resolve(`${_homeDir}/.office-addin-dev-certs/ca.crt`)
      ),
    },
  },
});
