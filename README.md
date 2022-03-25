# Fusion 360 Total Exporter
Need to export all of your fusion data in a bunch of formats? We've got you covered. This will export your designs across all hubs, projects, files, and components. It exports them in STL, STEP, and IGES format. It also exports the sketches as DXF files. In addition to those, an f3d/f3z file will be exported for each component/assembly file in all of your projects, across all hubs. It doesn't ask what you want written out, it just writes out the latest version of everything, to the folder of your choice.

## How do I use this?
1. Download this project from [here](https://github.com/Jnesselr/fusion-360-total-exporter/archive/master.zip), and unzip it (Or clone this repo, if you are familiar with git)
2. Open Fusion 360
3. Click on "Tools" then "Scripts/Add-ins"
4. Click the + button and select the unzipped folder
5. Double click on the "Fusion 360 Total Export" script
6. Acknowledge that this might take a while (There's a menu, but you should probably internalize that)
7. Select where you want the output to go - we suggest making a dedicated folder, to help keep everything nice and organized 
8. Go do something else for a while or enjoy a walk down memory lane as every single design you have is opened, exported, then closed again.

## Why did you make this?
Autodesk just announced that they were limiting features in their free tier to a level that made people a wee bit upset. I pay for Fusion 360, but I get that it's too much of an expense for some people. I had experience with exporting STL files for BotQueue (shhh spoilers) and figured that if I wrote a plugin, no one would have to do manual exports. Yay! Automation reigns supreme!

## What happens if I don't like what this plugin does?
No warrenty is implied, etc. etc. Go blame Autodesk for changing the free tier. If you want to blame me for anything, blame me and my sense of ethics for feeling like I need to write this program in the first place.

## What if I find a bug?
If an exception occurs, the script will let you know after it has exported everything that it can export. There will be a log file called output.log in your export folder. Submit an issue with that file attached. Please and thank you!

Also, if you can share the file that it failed on, that may help me, but it depends on what the exception actually shows.

## I don't like your style of writing
That wasn't a question. But yeah... me too some days.

# Update at 25.03.2022
I have edited (and somwhere rewrited) the script.
New features:
### +added progress file (to continue from the last file if it crashes). so you can continue from the last file if Fusion 360 crashes.

### +added export f2d (flat drawings) in PDF (https://forums.autodesk.com/t5/fusion-360-api-and-scripts/december-2020-product-update-api-quot-export-a-drawing-as-a-pdf/td-p/10050094)

### +added export in f3z (assemblies) thanks to this hub https://github.com/kantoku-code/Fusion360_Small_Tools_for_Developers/blob/master/TextCommands/TextCommands_txt_Ver2_0_8176.txt#L421

### +added support of Cyrillic symols in path names
