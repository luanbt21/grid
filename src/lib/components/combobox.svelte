<script lang="ts">
  import Check from "lucide-svelte/icons/check";
  import ChevronsUpDown from "lucide-svelte/icons/chevrons-up-down";
  import { tick } from "svelte";
  import * as Command from "$lib/components/ui/command/index.js";
  import * as Popover from "$lib/components/ui/popover/index.js";
  import { Button } from "$lib/components/ui/button/index.js";
  import { cn } from "$lib/utils.js";

  type Data = {
    value: string;
    label: string;
  };

  let {
    data,
    value = $bindable(""),
    placeholder = "item",
  }: { data: Data[]; value?: string; placeholder?: string } = $props();

  let open = $state(false);
  let triggerRef = $state<HTMLButtonElement>(null!);

  const selectedValue = $derived(data.find((s) => s.value === value)?.label);

  function closeAndFocusTrigger() {
    open = false;
    tick().then(() => {
      triggerRef.focus();
    });
  }
</script>

<Popover.Root bind:open>
  <Popover.Trigger bind:ref={triggerRef}>
    {#snippet child({ props })}
      <Button
        variant="outline"
        class="justify-between min-w-40"
        {...props}
        role="combobox"
        aria-expanded={open}
      >
        {selectedValue || `Select a ${placeholder}...`}
        <ChevronsUpDown class="opacity-50" />
      </Button>
    {/snippet}
  </Popover.Trigger>
  <Popover.Content class="max-w-full p-0">
    <Command.Root>
      <Command.Input placeholder={`Search ${placeholder}...`} />
      <Command.List>
        <Command.Empty>{`No ${placeholder} found.`}</Command.Empty>
        <Command.Group>
          {#each data as item}
            <Command.Item
              value={item.value}
              onSelect={() => {
                value = item.value;
                closeAndFocusTrigger();
              }}
            >
              <Check class={cn(value !== item.value && "text-transparent")} />
              {item.label}
            </Command.Item>
          {/each}
        </Command.Group>
      </Command.List>
    </Command.Root>
  </Popover.Content>
</Popover.Root>
