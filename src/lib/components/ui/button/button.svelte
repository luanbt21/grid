<script lang="ts">
	import { Button as ButtonPrimitive } from "bits-ui";
	import { type Events, type ButtonProps, buttonVariants } from "./index.js";
	import { cn } from "$lib/utils.js";

	type $$Props = ButtonProps;
	type $$Events = Events;

	interface Props {
		class?: $$Props["class"];
		variant?: $$Props["variant"];
		size?: $$Props["size"];
		builders?: $$Props["builders"];
		children?: import('svelte').Snippet;
		[key: string]: any
	}

	let {
		class: className = undefined,
		variant = "default",
		size = "default",
		builders = [],
		children,
		...rest
	}: Props = $props();
	
</script>

<ButtonPrimitive.Root
	{builders}
	class={cn(buttonVariants({ variant, size, className }))}
	type="button"
	{...rest}
	on:click
	on:keydown
>
	{@render children?.()}
</ButtonPrimitive.Root>
