# Google Sheets - Table Of Contents Generator
This handy script will let you generate Table of Contents sheets for your Google Sheets document.

![tableofcontentsscreenshot](https://user-images.githubusercontent.com/6147299/49584828-3482ed00-f922-11e8-93dd-7a26b6798fff.JPG)

## About ##
Google Sheets doesn't have a built in Table of Contents generator like Google Docs has, so I built one myself and also added a basic Outline system to it. I originally checked the web store and only found one addon that handled this functionality, but it was unreliable and easy to break. I decided to take one for the team and put this together in a few hours so that no one else would have to bang their heads against the awful Google API like I did to make one themselves.

## Installation ##
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

## Soapbox ##
I apologize that this is a bit of an inconvienence to install. Normally I'd try and make the process easier and faster, but this was a bit of a doozy. 

By the time I finished this, I was so tired of Google's awful APIs that I wanted to get as far away from it as possible. I had intended to make it an add on and put it on the Google Chrome Web Store for ease of use/convienence and also post it on here so people could access the code, but it's probably easier for everyone involved if I just post it here. I'm tired of spending hours trying to figure out how to use basic parts of their APIs because they think it's a good idea to try and support as many languages as possible (but fail to support a single one of them well). 
