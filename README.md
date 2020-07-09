# SharePoint CRUD (Create, Read, Update, Delete)  with Pnpjs
Table of contents
- [SharePoint CRUD (Create, Read, Update, Delete)  with Pnpjs](#sharepoint-crud-create-read-update-delete-with-pnpjs)
- [Introduction](#introduction)
- [PnPjs](#pnpjs)
- [Starter Kit](#starter-kit)
- [Example Repo crud-pnpjs-example](#example-repo-crud-pnpjs-example)
  - [Prerequisites](#prerequisites)
  - [Getting started](#getting-started)
  - [Building and running in production mode](#building-and-running-in-production-mode)
  - [Acknowledgments](#acknowledgments)
- [Create](#create)
- [Read](#read)
- [Update](#update)
- [Delete](#delete)
- [Extras](#extras)

# Introduction


# [PnPjs](https://pnp.github.io/pnpjs/)


# Starter Kit

- [sp-sveltejs/template](https://github.com/sp-sveltejs)

# Example Repo [crud-pnpjs-example](https://github.com/kayodebristol/crud-pnpjs-example)

```bash
npx degit kayodebristol/crud-pnpjs-example [your-app-name]
cd sp-svelte-app
```

*Note that you will need to have [Node.js](https://nodejs.org) installed.*

*Looking for a shareable component template the works with SharePoint? Coming Soon...*

---

## Prerequisites

Requires [Node.js](https://nodejs.org/)
It's very helpful if you have access to SharePoint, since this is a SharePoint development starter kit template.
The generated project will work with SharePoint 2013, SharePoint 2016, SharePoint 2019, and SharePoint Online. 

## Getting started

Install the dependencies...

```bash
cd [your-app-name]
npm install
```

Configure sp-rest-proxy
````
npm run proxy
```` 
then answer the interactive questions to configure the proxy connection to your SharePoint site. Recommend selecting On-Demand Credentials for the authentication strategy.
Ctrl-c to end task.

Start development
````
npm run dev
````
Uses concurrently, to start the proxy and dev server simultaneously.
* Develop interactively, with real SharePoint data. Enjoy!

Navigate to [localhost:5000](http://localhost:5000). You should see your app running. Edit a component file in `src`, save it, and reload the page to see your changes.

By default, the server will only respond to requests from localhost. To allow connections from other computers, edit the `sirv` commands in package.json to include the option `--host 0.0.0.0`.

## Building and running in production mode

To create an optimized version of the app:

```bash
npm run build
```
## Acknowledgments
Special thanks to
* [Rich Harris](https://github.com/Rich-Harris)
* [Andrew Koltyakov](https://github.com/koltyakov)


# Create 
 
# Read

# Update

# Delete

# Extras 



