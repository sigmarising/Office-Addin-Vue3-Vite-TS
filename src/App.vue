<script setup lang="ts">
import HelloWorld from './components/HelloWorld.vue'
import TheWelcome from './components/TheWelcome.vue'

const onClick = () => {
  Excel.run(async context => {
    const range = context.workbook.getSelectedRange();
    range.format.fill.color = 'green';
    await context.sync();
  });
}
</script>

<template>
  <header>
    <img alt="Vue logo" class="logo" src="./assets/logo.svg" width="125" height="125" />

    <div class="wrapper">
      <HelloWorld msg="You did it!" />
    </div>

    <button @click="onClick">Interact with Excel</button>
  </header>

  <main>
    <TheWelcome />
  </main>
</template>

<style scoped>
header {
  line-height: 1.5;
}

.logo {
  display: block;
  margin: 0 auto 2rem;
}

@media (min-width: 1024px) {
  header {
    display: flex;
    place-items: center;
    padding-right: calc(var(--section-gap) / 2);
  }

  .logo {
    margin: 0 2rem 0 0;
  }

  header .wrapper {
    display: flex;
    place-items: flex-start;
    flex-wrap: wrap;
  }
}

button {
  border: #34495e 1.5px solid;
  width: 100%;
  border-radius: 12px;
  margin-top: 30px;
  font-size: 1rem;
  padding: 3px;
  transition: all 0.1s ease-in-out;
}

button:hover {
  border: #41b883 2px solid;
  border-radius: 20px;
  margin-top: 30px;
  font-size: 1.1rem;
  color: #41b883;
  padding: 5px;
}

button:active {
  border: #34495e 2px solid;
  border-radius: 20px;
  margin-top: 30px;
  font-size: 1rem;
  color: #34495e;
  padding: 4px;
}
</style>
