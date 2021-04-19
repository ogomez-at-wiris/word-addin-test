# Introduction

This document outlines the configuration of the environment for developing MS Office Add-ins.

## First Steps

1. npm install -g yo generator-office office-addin-debugging
2. npm ci (to install all deps).
3. Install the Microsoft Office Debugger Visual studio code extension.
4. The launch.json file configuration that the extension suggests has already been configured in this repo, including the setup for the HOST application which may need to be changed if we decide to make an Excel or Powerpoint Add-in.

## Next Steps

1. Open the command prompt and ensure you are at the root folder of your project. Run the command npm start to start the dev server. When your add-in loads in the Office host application, open the task pane.
2. Return to Visual Studio Code and choose View > Debug or enter CTRL + SHIFT + D to switch to debug view.
3. From the Debug options, choose Attach to Office Add-ins. Select F5 or choose Debug -> Start Debugging from the menu to begin debugging.