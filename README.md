# ezcad-2-to-3
EZCAD2 to EZCAD3 Marking Parameters Library Converter


## Table of contents
* [General info](#general-info)
* [Technologies](#technologies)
* [Setup](#setup)
* [Configuration](#configuration)
* [Usage](#usage)
* [FAQ](#faq)

## General info
This project is an EZCAD2 to EZCAD3 Marking Parameters Library converter (PARAM/MarkParam.lib to PARAM/MarkParamlib.ini).

## Features
* Opens the EZCAD2 `MarkParam.lib` file from the current folder
* Reads the following items from the EZCAD2 file:
  * **[Section Name]**
  * **MARKSPEED** (mm/sec)
  * **POWERRATIO** (% of the machine's maximum wattage)
  * **FREQ** (kHz)
  * **QPULSEWIDTH** (ns)
  * **WOBBLEMODE** (**0** or **1**, i.e. OFF or ON)
  * **WOBBLEDIAMETER** (mm)
  * **WOBBLEDIST** (mm)
* Converts the numeric values of these parameters from the scientific notation to a decimal fraction with 6 places after the decimal point (except for the WOBBLEMODE, which should be either 0 or 1).
* Reads the EZCAD3 section template and injects the converted items into it
* Saves the new EZCAD3 file as `MarkParamlib.ini` into the current folder
* If you put an existing `MarkParamlib.ini` file into the current folder, the new items will get appended to the bottom of it, preserving all of your existing material parameters
* Configurable Source Wattage + Lens Size and Target Wattage + Lens Size parameters, which make the Power Ratio recalculate accordingly
* Moved user-updatable parameters to the very top of the file
* Made the exported Description field customizable
* Added Wattage and Lens size to the output filename
* Added more output statements as to what's going on: imported filename, processed section names, output filename
* Introduced script settings to designate Source and Target machines as MOPA
* Fixed the mapping of the non-MOPA Source to MOPA Target Pulse Width to 200ns (Hallman's info from LMA). The LMA settings file is universally set to 10ns, which is wrong, but the non-MOPA machines simply ignore it
* If both the Source and the Target machines are MOPA, made the target value snap to one of the values that the machine is actually configured to use

## Technologies
Project is created with:
* VBScript 5.6, 5.7, 5.8 - so it runs on any version of Windows from **Windows 7** and up
* Microsoft Windows Script Host (WSH) 5.6, 5.7, 5.8
	
## Setup
To run this project, download it locally, copy the MarkParam.lib to the same folder. Open the command prompt with the [Admin privileges](https://blog.techinline.com/2019/08/14/run-command-prompt-as-administrator-windows-10/). Then run the following command to set your default output to the command line window, instead of message box popups (you only need to do that once):

> **\> cscript /h:cscript**

If you ever want to revert your Windows to throw the message box popups later, you can use:

> **\> cscript /h:wscript**

## Configuration
Edit the `section-template.txt` file to adjust the [TC (Time Correction and Delay values)](https://www.youtube.com/watch?v=gFvbrNnvijo) to match your machine.  

Edit the `ezcad-2-to-3.vbs` file to set the source machine and target machine `Wattage`, `Lens size`, `MOPA/non-MOPA type`, and the `Description` field

## Usage

Simply run from the command prompt:
> **\> ezcad-2-to-3.vbs**

Once the script runs your new EZCAD3 `MarkParamlib.ini` file will be saved in the same folder. Move it to the `PARAM` subfolder in your EZCAD3 installation.

## FAQ
**Q: Why instead of editing the configuration settings inside the script file don't you expose them as script parameters?**  

A: I expect the target audience of this converter to have no script usage experience, so it is easier for them to simply edit the files.
