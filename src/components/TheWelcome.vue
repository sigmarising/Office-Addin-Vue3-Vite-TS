<script setup lang="ts">
import WelcomeItem from './WelcomeItem.vue'
import DocumentationIcon from './icons/IconDocumentation.vue'
import ToolingIcon from './icons/IconTooling.vue'
import EcosystemIcon from './icons/IconEcosystem.vue'
import CommunityIcon from './icons/IconCommunity.vue'
import SupportIcon from './icons/IconSupport.vue'
</script>

<template>
  <WelcomeItem>
    <template #icon>
      <DocumentationIcon />
    </template>
    <template #heading>Documentation</template>

    Vue’s
    <a href="https://vuejs.org/" target="_blank" rel="noopener">official documentation</a>
    provides you with all information you need to get started.

    <br />

    <a href="https://learn.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins" target="_blank"
      rel="noopener">Office Add-in's documentation</a>
    provides you with all the knowledge you need to interact with Office applications.
  </WelcomeItem>

  <WelcomeItem>
    <template #icon>
      <ToolingIcon />
    </template>
    <template #heading>Tooling</template>

    This project is served and bundled with
    <a href="https://vitejs.dev/guide/features.html" target="_blank" rel="noopener">Vite</a>. The
    recommended IDE setup is
    <a href="https://code.visualstudio.com/" target="_blank" rel="noopener">VSCode</a> +
    <a href="https://github.com/johnsoncodehk/volar" target="_blank" rel="noopener">Volar</a>. If
    you need to test your components and web pages, check out
    <a href="https://www.cypress.io/" target="_blank" rel="noopener">Cypress</a> and
    <a href="https://on.cypress.io/component" target="_blank" rel="noopener">Cypress Component Testing</a>.

    <br />

    This project using <a href="https://github.com/OfficeDev/Office-Addin-Scripts" target="_blank"
      rel="noopener">Office-Addin-Scripts</a>
    to enhance the ability to debug and test Office Add-in.

    <br />

    More instructions are available in <code>README.md</code>.
  </WelcomeItem>

  <WelcomeItem>
    <template #icon>
      <EcosystemIcon />
    </template>
    <template #heading>Ecosystem</template>

    Get official tools and libraries for your project:
    <a href="https://pinia.vuejs.org/" target="_blank" rel="noopener">Pinia</a>,
    <a href="https://router.vuejs.org/" target="_blank" rel="noopener">Vue Router</a>,
    <a href="https://test-utils.vuejs.org/" target="_blank" rel="noopener">Vue Test Utils</a>,
    <a href="https://github.com/vuejs/devtools" target="_blank" rel="noopener">Vue Dev Tools</a>, and
    <a href="https://github.com/vuejs/awesome-vue" target="_blank" rel="noopener">Awesome Vue</a>.

    <br />

    If you need more resources, we suggest paying
    <a href="https://github.com/OfficeDev" target="_blank" rel="noopener">Office Developer</a>
    a visit.
  </WelcomeItem>

  <WelcomeItem>
    <template #icon>
      <CommunityIcon />
    </template>
    <template #heading>Community</template>

    Got stuck about Vue? Ask your question on
    <a href="https://chat.vuejs.org" target="_blank" rel="noopener">Vue Land</a>, official
    Discord server, or
    <a href="https://stackoverflow.com/questions/tagged/vue.js" target="_blank" rel="noopener">StackOverflow</a>. 

    <br />

    Got stuck about Office JS? Ask your question on
    <a href="https://github.com/OfficeDev/office-js/issues" target="_blank" rel="noopener">issues thread</a>, or
    <a href="https://stackoverflow.com/questions/tagged/office-js" target="_blank" rel="noopener">StackOverflow</a>. 
  </WelcomeItem>

  <WelcomeItem>
    <template #icon>
      <SupportIcon />
    </template>
    <template #heading>Support this project</template>

    Welcome to contribute to this project by
    <a href="https://github.com/sigmarising/Office-Addin-Vue3-Vite-TS" target="_blank" rel="noopener">becoming a contributor</a>.
  </WelcomeItem>
</template>
