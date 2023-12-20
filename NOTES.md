# Steps to Construct this Project

## 1. Using create-vue to create the project

Using [create-vue](https://vuejs.org/guide/quick-start.html#creating-a-vue-application) to create the basic project structure.

The options used for create-vue:
```
✔ Add TypeScript? … Yes
✔ Add JSX Support? … No
✔ Add Vue Router for Single Page Application development? … No
✔ Add Pinia for state management? … No
✔ Add Vitest for Unit testing? … No
✔ Add an End-to-End Testing Solution? … No
✔ Add ESLint for code quality? … No
✔ Add Prettier for code formatting? … No
```

## 2. Install Office-Addin Packages

Install the packages of:
```json
{
  "devDependencies": {
    "@types/office-js": "^1.0.363",
    "@types/office-runtime": "^1.0.35",
    "office-addin-debugging": "^5.0.14",
    "office-addin-dev-certs": "^1.12.0",
    "office-addin-manifest": "^1.12.8"
  }
}
```

Also install the packages of:
```json
{
  "devDependencies": {
    "vite-plugin-static-copy": "^1.0.0"
  }
}
```

## 3. Make the project to fit Office-JS

Add the Add-in icons in the `public/` dirs.

Next add the type declare in `env.d.ts`:
```diff
+ /// <reference types="@types/office-js" />
+ /// <reference types="@types/office-runtime" />
```

Then modify the `src/main.ts`:
```diff
- createApp(App).mount("#app");
+ Office.onReady(() => {
+   createApp(App).mount("#app");
+ });
```

Add the Office JS runtime in the `index.html`'s `<head>`:
```html
<!-- Office JavaScript API -->
<script
  type="text/javascript"
  src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"
></script>
```

Then use `yo office` to generate a Excel Manifest Only.

Modify `package.json` to adapt Office JS Scripts:

```json
{
  "scripts":{
    "start": "office-addin-debugging start manifest.xml --dev-server vite",
    "start:desktop": "office-addin-debugging start manifest.xml desktop --dev-server vite",
    "start:web": "office-addin-debugging start manifest.xml web --dev-server vite",
    "stop": "office-addin-debugging stop manifest.xml",
    "manifest:validate": "office-addin-manifest validate manifest.xml",
    "manifest:update-guid": "office-addin-manifest modify manifest.xml --guid",
    "certs:install": "office-addin-dev-certs install ",
    "certs:verify": "office-addin-dev-certs verify",
    "certs:uninstall": "office-addin-dev-certs uninstall"
  }
}
```

Modify `vite.config.ts` to adapt Https add `vite-plugin-static-copy` to copy `manifest.xml` for development:
```diff
  import { fileURLToPath, URL } from "node:url";
+ import fs from "fs";
+ import path from "path";
+ import { homedir } from "os";
  
  import { defineConfig } from "vite";
  import vue from "@vitejs/plugin-vue";
+ import { viteStaticCopy } from "vite-plugin-static-copy";
  
+ // Configuration
+ const isDev = true;
+ const devPort = 8989;
+ const prodUrl = "https://your.addin.production.url/";
  
+ // Calculated
+ const _homeDir = homedir();
+ const _urlDefault = `https://localhost:3000/`;
+ const _urlDev = `https://localhost:${devPort}/`;
+ const _urlProd = `${prodUrl}`;
  
  // https://vitejs.dev/config/
  export default defineConfig({
    plugins: [
      vue(),
+     viteStaticCopy({
+       targets: [
+         {
+           src: "manifest.xml",
+           dest: "",
+           transform(content) {
+             return content
+               .toString()
+               .replace(
+                 new RegExp(_urlDefault, "g"),
+                 isDev ? _urlDev : _urlProd
+               );
+           },
+         },
+       ],
+     }),
    ],
    resolve: {
      alias: {
        "@": fileURLToPath(new URL("./src", import.meta.url)),
      },
    },
+   server: {
+     host: true,
+     port: devPort,
+     https: {
+       key: fs.readFileSync(
+         path.resolve(`${_homeDir}/.office-addin-dev-certs/localhost.key`)
+       ),
+       cert: fs.readFileSync(
+         path.resolve(`${_homeDir}/.office-addin-dev-certs/localhost.crt`)
+       ),
+       ca: fs.readFileSync(
+         path.resolve(`${_homeDir}/.office-addin-dev-certs/ca.crt`)
+       ),
+     },
+   },
  });
```

## 4. Modify the UI and add Excel Functionality

Mod the code in `src/**/*`.

## References

* [Use Vue to build an Excel task pane add-in](https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/excel-quickstart-vue)