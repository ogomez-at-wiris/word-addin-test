# Introduction

This document outlines the configuration of the environment for developing MS Office Add-ins.

## Setup environment

1. npm install -g yo generator-office office-addin-debugging
2. npm ci (to install all deps).
3. Install the Microsoft Office Debugger Visual studio code extension.
4. Microsoft's instructions for debugging Office Add-ins seems to be outdated. Go to the VSCode debug view (Ctrl + Shift + D) and click debug options. Then, choose Microsoft Word (Edge Chromium) if using Word.
5. Press F5 to begin debugging. It will open Word and begin debugging your extension as soon as you click on the command that appears in the ribbon for the Add-in.

