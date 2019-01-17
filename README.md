# Google Sheets - Table Of Contents Generator
This handy script will let you generate Table of Contents sheets for your Google Sheets document.

![capture](https://user-images.githubusercontent.com/6147299/51238036-806e7c00-193b-11e9-9b00-f35fb7f5933e.JPG)

## About ##
Google Sheets doesn't have a built in Table of Contents generator like Google Docs has, so I built one myself and also added a basic Outline system to it. I originally checked the web store and only found one addon that handled this functionality, but it was unreliable and easy to break. So I decided to take one for the team and put this together in a few hours. My hope is that my sacrifice means that no one else will ever have to bang their heads against the awful Google API like I did to get basic functionality out of their spreadsheets.

## Features ##
- Table of Contents Generator: a sheet with links to various defined sections can be automatically created
- Outline functionality that allows you to jump to sections quickly from the sidebar
- Sort button that will automatically re-organize your sheets alphabetically by section
- Built-in instruction manual that will tell you how to use the tool

## Chrome Web-Store Installation ##

I've made this tool available on the Chrome web-store. It's currently pending, but I'll post a link here when it's ready!

## Manual Installation ##
- Open Tools dropdown from the Google Sheet Toolbar
- Select Script Editor
- Copy Code.gs and Sidebar.html into the Script Editor from this repo
- Close the Script Editor

## Using It ##
After installing:
- Hover over Add-Ons
- Find 'NOKORI•WARE Table of Contents Generator'
- Hover over that entry to open the submenu containing all of this plugin's tools available
- For more in-depth instruction on usage, select the 'Open Instruction Manual' option

A table that's ready to have a Table of Contents generated from it will follow this format:

![tableexample](https://user-images.githubusercontent.com/6147299/49590311-330cf100-f931-11e8-817c-e83173ba6a6f.JPG)

Data used by table of contents generator are located in the A column, so that column must be reserved for section names. Cell A1 is the name of the Table of Contents. Everything below that cell denotes section names (e.g. equivalent to chapters in books). Empty cells are ignored. Non-empty cells are added to the Table of Contents and automatically linked to by the generator.
