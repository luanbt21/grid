<script lang="ts">
    import "../app.css";
    interface Props {
        children?: import("svelte").Snippet;
    }

    let { children }: Props = $props();
</script>

<!-- <svelte:head>
  <meta http-equiv="content-security-policy" content="script-src 'none'" />
</svelte:head> -->
<div class="container mx-auto min-h-screen py-2 px-4 sm:px-6 lg:px-8">
    <div class="text-center">
        <a href="/" class="text-4xl p-3">🧰</a>
    </div>
    {@render children?.()}
</div>
