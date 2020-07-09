<script>
	
	import {getData, addNewRecord} from './crud';
	import {onMount} from 'svelte'; 
	import {form, bindClass} from 'svelte-forms'; 
	
	
	let data = []; 

	

	//Uses the sveltejs lifecycle event to get data on application mount
	onMount(async()=>{
		data = getData(); 
	})
	let item ={

	 title: '', category: '', affectedResource: '', priority: '', status: '', webApp: '', navLocation: '', siteName: '', siteUrl: '', POC: '', engineer: '', section: '', unit: '', location: '', engineersLog: '', POCPhone: '', description: '', newSiteName: '', contentManagers: ''}; 
	const ticket = form(()=>({
		description: {
			value: item.description, 
			validators: ['required']
		}
	}))

	const addRecord = ()=> addNewRecord(item); 
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

<form>
<label for='title'>title:</label>
<input
	type='text'
	name='title'
	dind:value={item.title}
	use:bindClass={{ form: ticket}}
/>
<button on:click|preventDefault={addRecord}>submit</button>

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

