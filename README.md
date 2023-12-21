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

### pnpm

This project use pnpm, see this [instruction](https://pnpm.io/installation) to install pnpm.

### Manifest configuration

Before you go, notice that Office Add-in use [`manifest.xml`](./manifest.xml) for the Add-in description. please modify this file to fit your purpose.

The most things worth notice is the localhost URL used in the manifest. This specific the server url of your web app.

The default dev server port is `8989`. *If you want to change the dev server port*, please find and replace all url `https://localhost:8989/` to the url you want in the manifest and change the configuration part in [`vite.config.ts`](./vite.config.ts):
```ts
// Configuration
const devPort = 8989;
```

**When before publish the add-in, don't forget to replace these urls to the real deploy server url.**

### Init the project

Install all dependency packages:

```shell
pnpm install
```

### Https certs

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

### Start add-in debug

You can start the add-in debug by:

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

*If you only want to start the dev server, please run:*
```shell
pnpm run dev
```

### Manifest Utils

Before you publish the add-in, we suggest you update the guid and validate the manifest by:

```shell
# Update the Add-in GUID
pnpm run manifest:update-guid

# Validate the Manifest
pnpm run manifest:validate
```

### Build and Bundle

Before publish the add-in, run the following command to build and bundle all to folder `dist\`:

```shell
pnpm run build
```

Notice that the **`manifest.xml`** will also be copied to `dist\`.

## Compatibility

Project based on this template might use ES6 JavaScript, which is not compatible with [older versions of Office that use the Trident (IE 11) browser engine](https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/browsers-used-by-office-web-add-ins). For information on how to support those platforms, see [Support older Microsoft webviews and Office versions](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/support-ie-11).

## Feedback

[Create an issue](https://github.com/sigmarising/Office-Addin-Vue3-Vite-TS/issues) if you meet bugs or have feature request.

## Disclaimer

***This repo is community contributed and DO NOT have official support by Microsoft.***