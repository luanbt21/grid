<script>
  const tools = [
    { name: "Excel megerator", path: "/excel-merge" },
    { name: "Excel Corresponding", path: "/excel-corresponding" },
  ];
</script>

<div class="max-w-3xl mx-auto">
  <h1 class="text-4xl font-extrabold text-center text-gray-900 mb-8">
    Excel Tools
  </h1>

  <div class="bg-white shadow overflow-hidden sm:rounded-lg">
    <div class="px-4 py-5 sm:p-6">
      <h2 class="text-lg leading-6 font-medium text-gray-900 mb-4">
        Choose a tool to get started:
      </h2>
      <div class="grid grid-cols-1 gap-4 sm:grid-cols-2">
        {#each tools as tool}
          <a
            href={tool.path}
            class="inline-flex items-center justify-center px-4 py-2 border border-transparent text-sm font-medium rounded-md text-white bg-indigo-600 hover:bg-indigo-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500"
          >
            {tool.name}
          </a>
        {/each}
      </div>
    </div>
  </div>

  <div class="mt-8 text-center">
    <p class="text-base text-gray-500">
      Our Excel tools help you work more efficiently with your spreadsheets.
      Choose a tool above to get started!
    </p>
  </div>
</div>

