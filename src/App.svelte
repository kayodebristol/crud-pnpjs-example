<script>
	
	import {getData, addNewRecord} from './crud';
	import {onMount} from 'svelte'; 
	import { createForm } from "svelte-forms-lib";
	
	let data = []; 

	

	//Uses the sveltejs lifecycle event to get data on application mount
	onMount(async()=>{
		data = getData(); 
	})

	const { form, handleChange, handleSubmit} = createForm({
		initialValues:{
			title: "", 
			areaCode: "", 
			prefix: "",
			number: ""
		},
		onSubmit: values =>{
			alert(JSON.stringify(values));
			addNewRecord(values); 
		}, 

	})

	
</script>

<main>
	<h1>SharePoint svelte Template</h1>
	{#await data}
	<p>...Getting Data</p>
	{:then data}
	<h2> Data Source via PnPjs</h2>
	<p>SharePoint Site Title: {data.Title}</p>
	<p>SharePoint Site Description: {data.Description}</p>
	{/await}
	<p>Visit the <a href="https://svelte.dev/tutorial">Svelte tutorial</a> to learn how to build Svelte apps.</p>
</main>

<form on:submit={handleSubmit}>
<label for='title'>title:</label>
<input
	id="title"
	type='text'
	name='title'
	bind:value={$form.title}
	on:change={handleChange}
	
/>
<label for='areaCode'>areaCode:</label>
<input
	id="areaCode"
	type='text'
	name='areaCode'
	bind:value={$form.areaCode}
	on:change={handleChange}
/>
<label for='prefix'>prefix:</label>
<input
	id="prefix"
	type='text'
	name='prefix'
	bind:value={$form.prefix}
	on:change={handleChange}
/>
<label for='number'>number:</label>
<input
	id="number"
	type='text'
	name='number'
	bind:value={$form.number}
	on:change={handleChange}
/>
<br>
{#if $form.areaCode}       
<p>{`(${$form.areaCode}) ${$form.prefix}-${$form.number}`}</p>
{/if}
<button type="submit" >submit</button>

</form>
<style>
	main {
		text-align: center;
		padding: 1em;
		max-width: 240px;
		margin: 0 auto;
	}

	h1 {
		color: #ff3e00;
		text-transform: uppercase;
		font-size: 4em;
		font-weight: 100;
	}

	@media (min-width: 640px) {
		main {
			max-width: none;
		}
	}
</style>

