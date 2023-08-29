# my1stOfficeAdd-in
Learn to develop my 1st office add-in.  
The best part of learning it is that I have to write both front and server side ( by node.js though).
# How to learn
As I have a little coding skills, my route starts from [MS Office Add-ins Documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/overview/set-up-your-dev-environment?tabs=yeomangenerator). By this guidance, a Node.js environment is established on my win10 laptop.  *Come on, I'm not a full-time coder, and there is no more easier way for me to code on my laptop directly.*
## Steps.
### 1.  Familiar with coding environment.

### 2. Have a look of [Office JavaScript API](https://learn.microsoft.com/en-us/training/modules/understand-office-javascript-apis/index).  

>Since I'm familiar with VBA, all I need to do is build a function map between VBA and JS API.
### 3. Familiar with [Office Add-ins manifest](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/add-in-manifests).  


***
This part is still editing. 
***

## Details of building environment
The main tools to be installed are:
 - Node.js: *[Download and install Node.JS from Offical Site](https://nodejs.org/)*
 - npm: start a cli and ```npm install npm -g```
 - VS Code: *[Download and install VS Code from Offical Site](https://code.visualstudio.com/)*
 - Yo Office:   ```npm install -g yo generator-office```
 
    > - About using Yeoman generator, there is a more [detailed description and guidance of **Yeoman generator** here](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/yeoman-generator-overview).  
    > - After installed Yo Office, if you change dir to install dir and run `npm start`, it will open a new excel workspace with a warning ```To debug in Web vew, please attach VS code to your web view instance by **Microsoft Debugger for Edge** extension ``` blabla. ~~It seems can be solved by the guidance: [Debug a task pane add-in using Microsoft Edge DevTools Preview](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/debug-add-ins-using-devtools-edge-legacy)~~. Maybe it's not solved yet, I'm going to skip this issue recently.  
    >- After installed Yo Office, if command `yo offic`  is successfully run once, a project folder could be created.

 - Install the Office JavaScript linter: ```    npm install office-addin-lint --save-dev```  and     ``` npm install eslint-plugin-office-addins --save-dev```
   

- Install Script Lab: [install here from AppSource](https://appsource.microsoft.com/zh-CN/marketplace/apps?page=1&search=script%20lab).
    > About using Script Lab, there is a more [detailed description and guidance of **Script Lab** here](https://learn.microsoft.com/en-us/office/dev/add-ins/overview/explore-with-script-lab).


# Learn to code
 [Learn here](https://learn.microsoft.com/en-us/office/dev/add-ins/)

