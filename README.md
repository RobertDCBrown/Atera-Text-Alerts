# Atera-Text-Alerts
Powershell script to query Atera's ticket syetem. It will text ticket information to the on-call technician if a ticket comes in after hours. Business hours are defined by you.

Utilizes Atera's API, Twilio API, and TinyURL API. All required.

# Requirements
Your Atera API key
https://support.atera.com/hc/en-us/articles/219083397-APIs

A Twilio account with a phone number. I use the "Pay as you Go" plan and add $20 when I need to.

The PSAtera module
PS> Install-Module -Name PSAtera

# Configuration
On first launch, go to the settings tab and configure your Atera API key, Twilio Token, Twilio SID, Twilio Number. (format is +12223334444)
Add you on-call techs to the list. (format for cell phone is +12223334444)
Set your refresh rate to check Atera for new tickets. (Default: 5 minutes, Minimum: 1 minute)
