---
uid: dotnet-aad-query-sample
description: Learn how to use .NET Graph SDK to query Directory Objects
page_type: sample
createdDate: 09/22/2020 00:00:00 AM
languages:
- csharp
technologies:
  - Microsoft Graph
  - Microsoft identity platform
authors:
- id: Licantrop0
  displayName: Luca Spolidoro
products:
- ms-graph
- dotnet-core
- windows-wpf
extensions:
  contentType: samples
  technologies: 
    - Microsoft Graph
    - Microsoft identity platform
  createdDate: 09/22/2020 00:00:00 AM
codeUrl: https://github.com/microsoftgraph/dotnet-aad-query-sample
zipUrl: https://github.com/microsoftgraph/dotnet-aad-query-sample/archive/master.zip
description: "This sample demonstrates a .NET Desktop (WPF) application showcasing advanced Microsoft Graph Query Capabilities for Directory Objects with .NET"
---
# Explore advanced Microsoft Graph Query Capabilities for Directory Objects with .NET SDK

- [Explore advanced Microsoft Graph Query Capabilities for Directory Objects with .NET SDK](#explore-advanced-microsoft-graph-query-capabilities-for-directory-objects-with-net-sdk)
  - [Overview](#overview)
  - [Prerequisites](#prerequisites)
  - [Setup](#setup)
    - [Step 1: Register your application](#step-1-register-your-application)
    - [Step 2: Set the permissions](#step-2-set-the-permissions)
    - [Step 3: Configure the ClientId using the Secret Manager](#step-3-configure-the-clientid-using-the-secret-manager)
  - [Build & Run](#build--run)
    - [Using the app](#using-the-app)
  - [Code Architecture](#code-architecture)

## Overview

This sample helps you explore the Microsoft Graph's [new query capabilities](https://aka.ms/BlogPostMezzoGA) of the identity APIs using the [Microsoft Graph SDK](https://github.com/microsoftgraph/msgraph-sdk-dotnet) to query Azure AD.
The main code is in [GraphDataService.cs](MSGraphSamples.WPF/Services/GraphDataService.cs) file, where every request pass trough `AddAdvancedOptions` function adding the required `$count=true` QueryString parameter and `ConsistencyLevel=eventual` header.

## Prerequisites

- Visual Studio or just the .NET Core SDK
- An Internet connection
- A Windows machine (necessary if you want to run the app on Windows)
- An OS X machine (necessary if you want to run the app on Mac)
- A Linux machine (necessary if you want to run the app on Linux)
- An Azure Active Directory (Azure AD) tenant. For more information on how to get an Azure AD tenant, see How to get an Azure AD tenant
- A user account in your Azure AD tenant. This sample will not work with a Microsoft account (formerly Windows Live account). Therefore, if you signed in to the Azure portal with a Microsoft account and have never created a user account in your directory before, you need to do that now.

## Setup

### Step 1: Register your application

Use the [Microsoft Application Registration Portal](https://aka.ms/appregistrations) to register your application with the Microsoft Graph APIs.
![Application Registration](docs/register_app.png)
**Note:** Make sure to set the right Redirect URI for .NET Core apps: `http://localhost`.

### Step 2: Set the permissions

Add the delegated permissions for `Directory.Read.All`. We advise you register and use this sample on a Dev/Test tenant and not on your production tenant.

![Api Permissions](docs/api_permissions.png)

### Step 3: Configure the ClientId using the Secret Manager

This application use the [.NET Core Secret Manager](https://docs.microsoft.com/aspnet/core/security/app-secrets) to store the ClientId.
To add the ClientId created on step 1:

1. Open the Developer Command Prompt under `dotnet-aad-query-sample\MSGraphSamples.WPF\` directory
1. type `dotnet user-secrets set "clientId" "<YOUR CLIENT ID>"`

## Build & Run

If everything was configured correctly, you should be able to run the application, and see the first login prompt.
The auth token will be cached for the subsequent runs.

### Using the app

You can query your tenant using the standard OData `$filter`, `$search`, `$orderBy`, `$select` clauses in the relative textboxes.
In the screenshot below you can see the $search operator in action:
![Screenshot of the App](docs/app1.png)

If you double click on a row, a default drill-down will happen (for example by showing the list of transitive groups a user is part of).
If you click on a header, the results will be sorted by that column. **Note: not all columns are supported and you may receive an error**.
If any query error happen, it will displayed with a MessageBox.
The generated URL will appear in the readonly Url textbox. You can click the Graph Explorer button to open the current query in Graph Explorer.

## Code Architecture

This app provides a good starting point for enterprise desktop applications that connects to Microsoft Graph.  
The implementation is a classic WPF MVVM app with Views, ViewModels and Services. ICommand and INotifyPropertyChanged are manually implemented.
Dependency Injection is implemented using [`Microsoft.Extensions.DependencyInjection`](https://docs.microsoft.com/en-us/aspnet/core/fundamentals/dependency-injection?view=aspnetcore-3.1), supporting design-time data.  
Nullable and Code Analysis are enabled to enforce code quality.
