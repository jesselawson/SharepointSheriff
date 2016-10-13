# Example of how to create a single subsite with SPS-CreateSubsite

# Include the Sharepoint Sheriff Suite
."C:\Users\lawsonje\Documents\Github\Sharepoint-Sheriff\SPS-Suite.ps1"

#SPS-BulkCreateTeamsitesFromCSV -SiteUrl "https://butteedu.sharepoint.com/departments" -CsvFile "ctfg.csv" -Sheriff $TheSheriff

SPS-CreateSubsite -SiteUrl "https://butteedu.sharepoint.com/departments" -SubsiteUrl "SLED" -Title "Student Learning & Economic Development" -Template "Team" -SamePermissions $true
