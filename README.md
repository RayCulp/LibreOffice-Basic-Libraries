# LibreOffice Basic Libraries
This repository contains a collection of LibreOffice Basic functions that are designed to make programming LibreOffice Basic easier.

## Possible limitations
The functions provided in this repository were created and tested exclusively on Linux (Ubuntu 22.04). There are no plans to adapt any of these functions to work on Windows or Mac.

## Code readability and comments
The code has been written, indented and commented with maximum ease of readability and comprehensibility for VBA programmers who want to switch to LibreOffice Basic in mind. Shorthand techniques such as nested statements have been avoided to the greatest extent possible.

## How to sync the libraries in this repository with LibreOffice
You can use [obasync](https://github.com/imacat/obasync) to sync the libraries in this repository with the System and User Basic Macro Storage ("My Macros & Dialogs"). obasync does **not** appear to be able to sync with the Basic Macro Storage of a specific document. See [this video](https://www.youtube.com/watch?v=qB1rAAgkYGY) by [imacat](https://github.com/imacat) for more information.

## Libraries included in this repository
This repository includes the following libraries:

| Library name | Purpose |
|:------------- |:------------- |
| FIle IO | Provides functions to read from and write to files in the file system |