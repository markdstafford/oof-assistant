# Microsoft Graph + Azure Functions = Better Together

This repository contains a sample project demonstrating the use of Azure Functions with Microsoft Graph.

## Prerequisites

* Install git client: https://desktop.github.com/
* Install Visual Studio Code: https://code.visualstudio.com/
* Install NodeJS: https://nodejs.org/

## Clone repository

* `git clone https://github.com/markdstafford/oof-assistant.git`

## Install dependencies

* `npm install` to install project dependencies after moving into the project directory
* `npm run build` to start TypeScript compiler and watch for changes (leave this open in another console)

## Register app

* Register your application at https://apps.dev.microsoft.com/
* Use credentials from your Azure Pass subscription
* Add `http://localhost:3000` as a redirect URL under `Platforms > Web`
* Add Application permissions
  * Directory.ReadWrite.All
  * Calendars.Read
  * User.ReadWrite.All
  * Mail.Send
* Update `config.ts` with your application id, secret, and a tenant domain

## Grant consent

* Visit `https://login.microsoftonline.com/common/adminconsent?client_id=YOUR_APP_ID&state=12345&redirect_uri=http://localhost:3000` and grant the app access. Replace `YOUR_APP_ID`. After granting access you will be redirected to `localhost:3000` and nothing is running there, which is expected.

## Test locally

* `func host start`
* Attach debugger by pressing `F5` in Visual Studio Code
* Test with Postman or browser
  * http://localhost:7071/api/count-meetings
  * http://localhost:7071/api/count-meetings?clear=true to remove extensions
  * http://localhost:7071/api/send-email-summary
* Verify extensions in Graph Explorer
  * https://developer.microsoft.com/en-us/graph/graph-explorer
  * Authenticate as user from correct tenant
  * Enter request URL https://graph.microsoft.com/beta/users?$select=displayName&$expand=extensions

## Deploy to Azure Functions

* Create an Azure Function app at https://portal.azure.com
* You may need to setup credentials under `Platform Features > Code Deployment > Deployment Credentials`
* Configure your app to allow git deployments
  * `Platform features > Deployment Options > Local Git Repository` as the source.
* Add the git remote
  * Under `General Settings > Properties`, select your `Git Url`
  * `git remote add azure YOUR_GIT_URL`
* Push a deployment
  * `git push azure master`

## Explore code

* `git tag -l` to list steps
* `git checkout {tag}` to adjust code to a certain step

## Fiddle!
