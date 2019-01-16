# Google Sheets - Table Of Contents Generator
This handy script will let you generate Table of Contents sheets for your Google Sheets document.

![capture](https://user-images.githubusercontent.com/6147299/51238036-806e7c00-193b-11e9-9b00-f35fb7f5933e.JPG)

## About ##
Google Sheets doesn't have a built in Table of Contents generator like Google Docs has, so I built one myself and also added a basic Outline system to it. I originally checked the web store and only found one addon that handled this functionality, but it was unreliable and easy to break. So I decided to take one for the team and put this together in a few hours. My hope is that my sacrifice means that no one else will ever have to bang their heads against the awful Google API like I did to get basic functionality out of their spreadsheets.

## Features ##
- Table of Contents Generator
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
- Open Tools dropdown from the Google Sheet Toolbar
- Hover over "Macros"
- Select "Import"
- Find "NOKORIWARETableOfContentsGenerator"
- Click "ADD FUNCTION" to its lower right

You can rename the function to something prettier like "Table of Contents Generator" by using the Manage Macros button.

## Using It ##
After installing:
- Open Tools dropdown from the Google Sheet toobar
- Hover Macros
- Click NOKORIWARETableOfContentsGenerator (or whatever you changed the name to)

This will start the program and a sidebar will appear on the right. From there, you can click the Instructions/About button to see more information on how to use it.

A table that's ready to have a Table of Contents generated from it will follow this format:

![tableexample](https://user-images.githubusercontent.com/6147299/49590311-330cf100-f931-11e8-817c-e83173ba6a6f.JPG)

The table of contents sections are in the A column. Empty cells are ignored. Non-empty cells are added to the Table of Contents and automatically linked to by the generator.
