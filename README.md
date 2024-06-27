# LibreOffice Basic Libraries
This repository contains a collection of LibreOffice Basic functions that are designed to make programming LibreOffice Basic easier.

## Possible limitations
The functions provided in this repository were created and tested exclusively on Linux (Ubuntu 22.04). There are no plans to adapt any of these functions to work on Windows or Mac.

## Code readability and comments
The code has been written, indented and commented with maximum ease of readability and comprehensibility for VBA programmers who want to switch to LibreOffice Basic in mind. Shorthand techniques such as nested statements have been avoided to the greatest extent possible.

## How to sync the libraries in this repository with LibreOffice
To sync the libraries in this repository with the System and User Basic Macro Storage ("My Macros & Dialogs") of your LibreOffice installation, first download (or git clone) the libraries and the modules contained in them to a folder on your local drive. Then you can use [obasync](https://github.com/imacat/obasync) to sync the downloaded libraries with your System and User Basic Macro Storage ("My Macros & Dialogs"). 

|   |   |
|:---|:---|
| ![image](https://github.com/RayCulp/LibreOffice-Basic-Libraries/assets/7621330/bdbaf1bb-9277-48d4-9df5-784e26f443d4) | __NOTICE__ <br/>Currently, obasync does **not** appear to be able to sync with the Basic Macro Storage of a specific document.|

See [this video](https://www.youtube.com/watch?v=qB1rAAgkYGY) by [imacat](https://github.com/imacat) for more information.

## Libraries included in this repository
This repository includes the following libraries:

| Library name | Purpose |
|:------------- |:------------- |
| FIle IO | Provides functions to read from and write to files in the file system |
