# DNS-Updater
DNS Updater

VBscript, ZoneEdit, DynDNS

Script checks to see if your dynamic IP address has changed by getting a response from icanhazip.com. It then reads a text file containing the last checked IP address. If current IP address is different, the script will update your DNS service (ZoneEdit, DollarDNS, DynDNS). Emails alert regardless, to let you know that the script ran and the status of the your IP, DNS. 

# Install Instructions
Script is designed for ZoneEdit dynamic DNS service. You may be able to modify the script to work with other DynDNS like services. Set your Dynamic DNS service settings for the variables. Set your email values (currently configured to work with gmail).
Create a scheduled task in Windows. Set the task to start on Start Up. Set a user for the task. 

