# Example of how to output a clean HTML table with the name and links to all that teamsite's subsites using SPS-GetTabledSubsites

# Include the Sharepoint Sheriff Suite
."C:\Users\lawsonje\Documents\Github\Sharepoint-Sheriff\SPS-Suite.ps1"

SPS-GetTabledSubsites -SiteUrl "https://yourspsite.sharepoint.com/parentSite" -OutFile "parentSite-subsites.html"
