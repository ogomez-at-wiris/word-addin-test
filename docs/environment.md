# Introduction

This document outlines the configuration of the environment for developing MS Office Add-ins.

## Setup environment

1. npm install -g yo generator-office office-addin-debugging
2. npm ci (to install all deps).
3. Search for the "Debugger for Microsoft Edge" extension and install it.
4. The launch.json for this project has been modified with Microsoft's instructions for debugging in Chromium based edge, but for other Add-ins it will need to be done using this template.
5. Microsoft's instructions for debugging Office Add-ins seems to be outdated. Go to the VSCode debug view (Ctrl + Shift + D) and click debug options. Then, choose Microsoft Word (Edge Chromium) if using Word.
6. Press F5 to begin debugging. It will open Word and begin debugging your extension as soon as you click on the command that appears in the ribbon for the Add-in. Click on cancel on the dialog box that pops up saying that "you need Microsoft Edge debugger extension" (Microsoft says it will appear, so it is not an error).

