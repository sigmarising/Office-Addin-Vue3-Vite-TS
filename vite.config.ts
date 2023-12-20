import { fileURLToPath, URL } from "node:url";
import fs from "fs";
import path from "path";
import { homedir } from "os";

import { defineConfig } from "vite";
import vue from "@vitejs/plugin-vue";
import { viteStaticCopy } from "vite-plugin-static-copy";

const _homeDir = homedir();

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [
    vue(),
    viteStaticCopy({
      targets: [
        {
          src: "manifest.xml",
          dest: "",
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
    port: 8989,
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
