<script lang="ts">
  import "../app.css";
  interface Props {
    children?: import('svelte').Snippet;
  }

  let { children }: Props = $props();
</script>

<!-- link to homepage -->
<div
  class="container mx-auto min-h-screen bg-gray-100 py-2 px-4 sm:px-6 lg:px-8"
>
  <div class="text-center">
    <a href="/" class="text-4xl p-3">🧰</a>
  </div>
  {@render children?.()}
</div>
