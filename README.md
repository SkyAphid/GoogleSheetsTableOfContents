# Google Sheets - Table Of Contents Generator
This handy script will let you generate Table of Contents sheets for your Google Sheets document.

![capture](https://user-images.githubusercontent.com/6147299/51238036-806e7c00-193b-11e9-9b00-f35fb7f5933e.JPG)

## About ##
Google Sheets, despite being really great free software, unfortunately doesn't have a built-in table of contents/outline tool like Google Docs has. I originally checked the web store and only found one add-on that handled this functionality, but it was unreliable and easy to break. So I decided to take one for the team and put this together in a few hours. I'm posting the software here because I know there has to be others like me who need it for their various large-scale projects.

## Features ##
- Table of Contents Generator: a sheet with links to various defined sections can be automatically created
- Outline functionality that allows you to jump to sections quickly from the sidebar
- Sort button that will automatically re-organize your sheets alphabetically by section
- Built-in instruction manual that will tell you how to use the tool

## Installation ##
- Open Tools dropdown from the Google Sheet Toolbar
- Select Script Editor
- Copy Code.gs and Sidebar.html into the Script Editor from this repo
- Close the Script Editor

## Using It (After Installation) ##
- Hover over Add-Ons
- Find 'NOKORI•WARE Table of Contents Generator'
- Hover over that entry to open the submenu containing all of this plugin's tools available
- For more in-depth instruction on usage, select the 'Open Instruction Manual' option

A table that's ready to have a Table of Contents generated from it will follow this format:

![tableexample](https://user-images.githubusercontent.com/6147299/49590311-330cf100-f931-11e8-817c-e83173ba6a6f.JPG)

Data used by table of contents generator are located in the A column, so that column must be reserved for section names. Cell A1 is the name of the Table of Contents. Everything below that cell denotes section names (e.g. equivalent to chapters in books). Empty cells are ignored. Non-empty cells are added to the Table of Contents and automatically linked to by the generator.
