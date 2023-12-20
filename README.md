# Office-Addin-Vue3-Vite-TS

This repository contains template to get start writing a TaskPane Office Add-in with the Vue3 framework using Vite and TypeScript.

**Feature:**
* Vue3
* Vite
* TypeScript
* Excel API

If you have interested about how this project is constructed, please see detail in this [NOTES](./NOTES.md).

## Recommended IDE

We recommend [VSCode](https://code.visualstudio.com/) + [Volar](https://marketplace.visualstudio.com/items?itemName=Vue.volar) (and disable Vetur) + [TypeScript Vue Plugin (Volar)](https://marketplace.visualstudio.com/items?itemName=Vue.vscode-typescript-vue-plugin).

We also recommend enable the [Volar Takeover Mode](https://vuejs.org/guide/typescript/overview.html#volar-takeover-mode).

## Project CMDs

This project use pnpm, see this [instruction](https://pnpm.io/installation) to install pnpm.

Before you go, please modify the configuration part in [`vite.config.ts`](./vite.config.ts):
```ts
// Configuration

// `true` for Dev env
// `false` for Prod env
const isDev = true;

// port that dev server will run on
// the manifest will be modified with this port when run in dev env
const devPort = 8989;

// only used in Prod env
// the manifest will be modified with this url when run in prod env
const prodUrl = "https://your.addin.production.url/";
```

Then init the project by:

```shell
pnpm install
```

Before we start the add-in, please run the following commands to install the https certs:

```shell
pnpm run certs:install
```

The certs verify and uninstall commands also are provided:
```shell
# Verify the Certs
pnpm run certs:verify

# Uninstall the Certs
pnpm run certs:uninstall
```

Next, you can start the add-in debug by:
```shell
# Start Add-in in Desktop
pnpm run start

# Start Add-in in Desktop
pnpm run start:desktop

# Start Add-in in Office Online
pnpm run start:web --document "https://url-to-online-doc"
```

Stop the add-in debug by:
```shell
pnpm run stop
```

> *If you only want to start the dev server, please run `pnpm run dev`*

Before you publish the add-in, we suggest you update the guid and validate the manifest by:
```shell
# Update the Add-in GUID
pnpm run manifest:update-guid

# Validate the Manifest
pnpm run manifest:validate
```

Finally, run the following command to build and bundle all to folder `dist\` (notice that the **final manifest.xml** will be also output here):
```shell
pnpm run build
```

## Compatibility

Project based on this template might use ES6 JavaScript, which is not compatible with [older versions of Office that use the Trident (IE 11) browser engine](https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/browsers-used-by-office-web-add-ins). For information on how to support those platforms, see [Support older Microsoft webviews and Office versions](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/support-ie-11).

## Feedback

[Create an issue](https://github.com/sigmarising/Office-Addin-Vue3-Vite-TS/issues) if you meet bugs or have feature request.

## Disclaimer

***This repo is community contributed and DO NOT have official support by Microsoft.***