# Introduction

This document outlines the configuration of the environment for developing MS Office Add-ins.

## First Steps

1. npm install -g yo generator-office office-addin-debugging
2. npm ci (to install all deps).
3. Install the Microsoft Office Debugger Visual studio code extension.
4. Microsoft's instructions for debugging Office Add-ins seems to be outdated. Go to the VSCode debug view (Ctrl + Shift + D) and click debug options. Then, choose Microsoft Word (Edge Chromium) if using Word.
5. Press F5 to begin debugging. It will open Word and begin debugging your extension as soon as you click on the command that appears in the ribbon for the Add-in.

## Next Steps

1. Open the command prompt and ensure you are at the root folder of your project. Run the command npm start to start the dev server. When your add-in loads in the Office host application, open the task pane.
2. Return to Visual Studio Code and choose View > Debug or enter CTRL + SHIFT + D to switch to debug view.
3. From the Debug options, choose Attach to Office Add-ins. Select F5 or choose Debug -> Start Debugging from the menu to begin debugging.