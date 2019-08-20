### **`sp-devops`** is a command line utility to manage devops for **Microsoft Sharepoint Online** for modern Front-End projects.

# Introduction
**sp-devops** is a command line utility to manage devops for **Microsoft Sharepoint Online** for modern Front-End projects.

There are very few custom sites built on modern Front-End frameworks on Sharepoint Online 365. Its a big hassle to manage the deployment which is completely manual. To address this problem I have created this library and is being used in a project of mine



Checkout [**ng-sharepoint**](https://github.com/joyblanks/ng-sharepoint) (https://github.com/joyblanks/ng-sharepoint), a sample project written in Angular and is managed by `sp-devops` to deploy to sharepoint and its databases/lists are being created by it.

# Installation
As a dev dependency in your project
```sh
$ npm i -D sp-devops # npm install sp-devops --save-dev
```
As a global tool
```sh
$ npm i -g sp-devops
```

# Highlights

Deployment: 
```sh
$ sp-devops --deploy
```

Setup (create DBs, Sites, Populate, etc.): 
```
$ sp-devops --setup
```

Get a Authorization Bearer Access token to use in REST API's
```
$ sp-devops --accesstoken
```

# Usage

In your package.json add a script
```json
scripts: {
  "sp-ng-build-deploy": "npm run build --prod && sp-devops --deploy",
  "sp-ng-create-lists": "sp-devops --setup --sp-spec-list=./<path-to>/<spec.json>",
  "sp-get-access-token": "sp-devops --accesstoken"
}
```

Then from terminal run (This will deploy the site in a sharepoint folder)

```sh
$ npm run sp-ng-build-deploy
```

You can also setup all your lists/databases from here. Please note the script tries to drop the databases and then recreates them in order to setup. You can pre-populate the lists as well by providing data in a specification file.


```sh
$ npm run sp-ng-create-lists
```


Get a Authorization Bearer Access token to use in REST API's in POSTMAN or in fetch API's useful for standalone testing
```
$ npm run sp-get-access-token
```



# Command-Line-Args and Environment-Variables
[NOTE]*: Do not put property `accesstoken`, `deploy` and `setup` in .env files

| ENV \| CLI property | Description | Usage |
|--|--|--|
| deploy* | Deploys the distribution to a Sharepoint folder | --deploy |
| setup* | Create Lists/Sites, Preload data etc.|--setup |
| accesstoken* | Get an Authorization Bearer accessToken.|--accesstoken |
| sp-site-url | Your Sharepoint Site Domain | --sp-site-url=https://[$tenant].sharepoint.com|
| sp-subsite | If you are working on a SubSite |--subsite=Sub1/Sub2 |
| sp-app-client-id | Sharepoint App Client ID | --sp-app-client-id=aaaaaaa-zzzz-1234-wxyz|
| sp-app-client-secret | Sharepoint App Client Secret| --sp-app-client-secret=0000000000 |
| sp-access-token | Sharepoint Authorization Bearer Token | --sp-access-token=Bearer [accessToken]|
| sp-dist-folder | Path Local Distribution folder (Source)| --sp-dist-folder=../bin/src|
| sp-remote-folder | Path to Remote Folder in Sharepoint (Destination) | --sp-remote-folder=SiteAssets/App|
| sp-log-level | Logging level of the Application | --sp-log-level=debug|
| sp-spec-list | Path to a JSON specification file <br>which contains configuration to create lists <br> and if data needs to be pre-polulated after creation  | --sp-spec-list=./configs/create-list.json|

# Sample Run [--deloy]
![Sample Run --deploy](https://github.com/joyblanks/sp-devops/raw/master/assets/sample-run-deploy.png)


# Sample Run [--setup]
![Sample Run --setup](https://github.com/joyblanks/sp-devops/raw/master/assets/sample-run-setup.png)


# Sample Run [--accesstoken]
![Sample Run --setup](https://github.com/joyblanks/sp-devops/raw/master/assets/sample-run-accesstoken.png)


# Setup and Configuration
- Setup a Sharepoint App to run sp-devops
  - Visit **/_layouts/15/appregnew.aspx** and generate your App Client Id and Client Secret. Provide details as shown in the diagram below and click create.

    
    | Figure: Create Sharepoint App |
    |:-:|
    | ![Create Sharepoint App](https://github.com/joyblanks/sp-devops/raw/master/assets/create-sharepoint-app.jpg) |
    
  
  
  - Copy the Client ID and Secret to later use this library.
  - Visit /_layouts/15/appinv.aspx and enter your App Client Id that you have just created and click **Lookup**. 
  
  - Enter the Permission Request XML on the last field and click **Create**.
    ```xml
    <AppPermissionRequests AllowAppOnlyPolicy="true">
      <AppPermissionRequest Scope="http://sharepoint/content/sitecollection" Right="Manage" />
      <AppPermissionRequest Scope="http://sharepoint/content/sitecollection/web" Right="Manage" />
      <AppPermissionRequest Scope="http://sharepoint/content/sitecollection/web/list" Right="FullControl" />
    </AppPermissionRequests>
    ```
    
    
    | Figure: Request Permission for Sharepoint App |
    |:-:|
    | ![Request Permission for Sharepoint App](https://github.com/joyblanks/sp-devops/raw/master/assets/app-request-permission.jpg) |
    

  - Click on **Trust It** button to grant the Sharepoint App the required permissions
    
    | Figure: Grant Permission to Sharepoint App |
    |:-:|
    | ![Grant Permission to Sharepoint App](https://github.com/joyblanks/sp-devops/raw/master/assets/app-grant-permission.jpg) |
    
    
    

- If you do not have  authorization to create an App, you can ask your Sharepoint admin to generate an **Authorization Bearer 'AccessToken'** token and share with you which you can also use to fire this library.
- Create a **.env** file in your project and specify these values
  ```sh
  # -------------------------------------------- Common --------------------------------------------#

  # Your Sharepoint Site Domain
  sp-site-url=https://${tenant}.sharepoint.com

  # If you are deploying to a subsite (Uncomment SP_SUBSITE)
  # sp-subsite=mysubsite

  # Create a Sharepoint App 
  # then get details from /_layouts/15/AppRegNew.aspx (generate App ClientId and ClientSecret here)
  # and trust your app from /_layouts/15/appinv.aspx (Give permissions and Trust app here)

  # Sharepoint APP's CLIENT_ID & CLEINT_SECRET
  sp-app-client-id=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
  sp-app-client-secret=*------------------------------*

  # Access Token from other sources (If you dont have Sharepoint APP's CLIENT_ID & CLEINT_SECRET)
  # sp-access-token=Bearer <ACCESS_TOKEN>

  # Logging level Allowed = [info, debug, warn, error]
  sp-log-level=info

  # ------------------------------------------ Deployment ------------------------------------------#
  # Sharepoint-Remote-Folder: where the code-base is to be uploaded
  # Make sure it is a dedicated folder all files/folders are deleted before uploading
  sp-remote-folder=SiteAssets

  # Path to your local distribution folder location
  sp-dist-folder=dist

  # --------------------------------------------- Setups ---------------------------------------------#
  # Path to spec file for creating Lists, Sites, Populate data, etc.
  #sp-spec-list=./sharepoint-list-spec.json
  #sp-spec-site=./sharepoint-site-spec.json
  ```

- In your package.json add a script and run (Refer #Usage)
  ```json
  scripts: {
    "sp-ng-build-deploy": "npm run build --prod && sp-devops --deploy"
  }
  ```
- For Setup (like creating DBs and Sites/Subsites) [Look under ./assets/samples in this repository]
  
  - you need to pass a JSON file to create your lists or to pre-populate data
    ```json
    {
      "config": [
        {
          "name": "SampleList",
          "addToView": true,
          "addToQuickLaunch": true,
          "dropIfExists": false,
          "columns": [
            { "__metadata": { "type": "SP.Field" }, "FieldTypeKind": 3, "Title": "Message"  },
            { "__metadata": { "type": "SP.Field" }, "FieldTypeKind": 20, "Title": "AssignedTo" },
            {
              "__metadata": { "type": "SP.FieldChoice" }, "FieldTypeKind": 6, "Title": "Status",
              "Choices": {
                "__metadata": { "type": "Collection(Edm.String)" },
                "results": [
                  "Created", "Assigned", "In-Progress", "Completed", "Archived", "On-Hold", "Cancelled"
                ]
              }
            }
          ]
        },
        {
          "name": "ConfigList",
          "addToView": false,
          "addToQuickLaunch": true,
          "dropIfExists": true,
          "columns": [
            { "__metadata": { "type": "SP.Field" }, "FieldTypeKind": 3, "Title": "JSON" },
            { "__metadata": { "type": "SP.Field" }, "FieldTypeKind": 3, "Title": "Description" }
          ],
          "items": [
            {
              "Title": "App",
              "JSON": "{\"pageSize\": 10, \"theme\": \"light\"}",
              "Description": "Config for App (light theme)"
            },
            {
              "Title": "Dark-App",
              "JSON": "{\"pageSize\": 10, \"theme\": \"dark\"}",
              "Description": "Config for App (dark theme)"
            }
          ]
        }
      ]
    }
    ```

  - In you package.json add a script and run (Refer #Usage)
    ```json 
    scripts: {
      "sp-ng-create-lists": "sp-devops --setup --sp-spec-list=./<path-to>/<spec.json>"
    }
    ```
- All items in CLI/ENV options can be triggered. [NOTE]: CLI options overrides the ENV vars.
  ```sh
  npm run sp-ng-create-lists -- --sp-subsite=inside/more_inside
  ```

---

Please feel free to log Pull requests, issues, bugs and improvement ideas


Author: [**Joy Biswas** @joyblanks](https://github.com/joyblanks/sp-devops/commits?author=joyblanks)
