<script type="text/javascript" src="https://cdnjs.buymeacoffee.com/1.0.0/button.prod.min.js" data-name="bmc-button" data-slug="Cidks7T2sy" data-color="#5F7FFF" data-emoji="🍺"  data-font="Cookie" data-text="Buy me a beer" data-outline-color="#000000" data-font-color="#ffffff" data-coffee-color="#FFDD00" ></script>

# Atera-Text-Alerts
Powershell script with WinForms GUI to query Atera's ticket syetem. It will text ticket information to the on-call technician if a ticket comes in after hours. Business hours are defined by you.

Utilizes Atera's API, Twilio API, and TinyURL API.

# Special Thanks
David Long for creating PSAtera, an easy to use module for utilizing the Atera API.
https://github.com/davejlong/PSAtera

# Requirements
Your Atera API key
https://support.atera.com/hc/en-us/articles/219083397-APIs

A Twilio account with a phone number. I use the "Pay as you Go" plan and add $20 when I need to.

The PSAtera module

```PS> Install-Module -Name PSAtera```

# Configuration
On first launch, go to the settings tab and configure your Atera API key, Twilio Token, Twilio SID, Twilio Number. Format of number is +12223334444

Add you on-call techs to the list. Format for cell phone is +12223334444 (Change country code as needed)

Set your refresh rate to check Atera for new tickets. (Default: 5 minutes, Minimum: 1 minute)

# Optional

If downloading the source .ps1 file, compile to an EXE using PS2Exe. I find this to be the best way to run.

```ps2exe "<path of ps1 file.ps1>" "<destination of exe.exe>"```

Add .exe to startup directory or task scheduler so it is always running at login.

# Screenshots

![alt text](https://i.imgur.com/3DSiw6G.png)

![alt text](https://i.imgur.com/JGSDL1r.png)

