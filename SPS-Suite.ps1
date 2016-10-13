
<# 

    SPS-CreateSubsite

Purpose.
    Create a sigle subsite

Usage:
    See Example_SPS-CreateSubsite.ps1



#>
Function SPS-CreateSubsite {
    Param(
        [Parameter(Mandatory=$true,HelpMessage="The URL of the site collection",Position=0)][ValidateNotNull()]
        [string]$SiteUrl,

        [Parameter(Mandatory=$true,HelpMessage="The short URL of the subsite",Position=1)][ValidateNotNull()]
        [string]$SubsiteUrl,

        [Parameter(Mandatory=$true,HelpMessage="The site title",Position=2)][ValidateNotNull()]
        [string]$Title,

        [Parameter(Mandatory=$true,HelpMessage="The template to use, Project or Team",Position=3)][ValidateSet("Project","Team")]
        [string]$Template,

        [Parameter(Mandatory=$true,HelpMessage="Use Same Permissions As Parent Site",Position=4)]
        [boolean]$SamePermissions,

        [Parameter(Mandatory=$false,HelpMessage="Site description",Position=5)]
        [string]$Description

    )
    begin{
        # Get user authentication
        $Username = "lawsonje@butte.edu"

        $tmplt = ""

        # You can define what type of template you want to use here. For example, if you have a custom template, you can set the template identifier here to correspond with an easy-to-remember (and use)
        # template name. The default is a project site, but if you pass `--Template Team` to the cmdlet, then you'll get a STS#0, which, in SharePoint world, is a default Team Site. 
        If($Template -eq "Team"){
            $tmplt = "STS#0"
        }
        Else {
            $tmpt = "PROJECTSITE#0"
        }

        $subsite = $SiteUrl+"/"+$SubsiteUrl
        write-host "Creating subsite $subsite"

    }
    process{
        # Create a new sharepoint context
        $context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
        
        # Read in the password so we don't have to bake in the credentials. If you want, you can comment out the read-host line below and instead use the 
        # explicit password declaration if you're tired of typing in your password all the time.
        $Password = read-host -Prompt "Password" -AsSecureString
        #$Password = "myPassword"
        
        $context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Username, $Password)
   
        #Create SubSite
        $wci = New-Object Microsoft.SharePoint.Client.WebCreationInformation
        $wci.WebTemplate = $tmplt
        if ($Description.Length -gt 0){
            $wci.Description = $Description
        }
        $wci.UseSamePermissionsAsParentSite = $SamePermissions
        $wci.Title = $Title
        $wci.Url = $SubsiteUrl
        $wci.Language = "1033"

        $SubWeb = $context.Web.Webs.Add($wci)

        $noproblems = $true

        try {
            $context.ExecuteQuery()
        }
        catch {
            $noproblems = $false
            write-host "Unable to create new subsite!" -ForegroundColor Yellow
            # If you do get an error, it will look like this: `Exception calling "ExecuteQuery" with "0" argument(s): "The Web site address "something" is already in use."
            # The split() here will only give us the sentence that we care about. 
            write-host $_.Exception.Title -ForegroundColor Red
            write-host $_.Exception.Message -ForegroundColor Red
        }
        finally {
            if($SamePermissions -eq $false){
                try{
                    #Check to see if this is a root site collection that was passed
                    Get-SPOSite -Identity $SiteUrl
                }
                catch {
                #If here we are working with a subsite
                Write-Host "Oops, it looks like we're dealing with a subsite, I can handle that" -ForegroundColor Yellow 
                #reconstruct the SiteUrl parameter
                $SiteUrl = "https://butteedu.sharepoint.com/"+$SiteUrl.Split("/")[3]+"/"+$SiteUrl.Split("/")[4]
                }
                finally{
                    $ownerGroup = "$Title Owners"
                    $memberGroup = "$Title Members"
                    $visitorGroup = "$Title Visitors"
                    Write-Host "Hang tight, creating groups, updating ownership and adjusting permissions"
                    New-SPOSiteGroup -Site $SiteUrl -Group $ownerGroup -PermissionLevels "Read" | Out-Null
                    #Updating the group owner does not work from PowerShell, using the new Set-GroupOwner function
                    Set-GroupOwner -SiteUrl $SiteUrl -GroupToUpdate $ownerGroup.ToString() -GroupOwner $ownerGroup.ToString()
                    New-SPOSiteGroup -Site $SiteUrl -Group $memberGroup -PermissionLevels "Read" | Out-Null
                    Set-GroupOwner -SiteUrl $SiteUrl -GroupToUpdate $memberGroup.ToString() -GroupOwner $ownerGroup.ToString()
                    New-SPOSiteGroup -Site $SiteUrl -Group $visitorGroup -PermissionLevels "Read" | Out-Null
                    Set-GroupOwner -SiteUrl $SiteUrl -GroupToUpdate $visitorGroup.ToString() -GroupOwner $ownerGroup.ToString()
                    Set-PermissionsOnSite -Url $subsite -GroupName $ownerGroup -Roletype "Full Control" 
                    Set-PermissionsOnSite -Url $subsite -GroupName $memberGroup -Roletype "Contribute"  
                    Set-PermissionsOnSite -Url $subsite -GroupName $visitorGroup -Roletype "Read"  
                    Set-DefaultSiteGroups -SiteUrl $subsite -VisitorGroup $visitorGroup -MemberGroup $memberGroup -OwnerGroup $ownerGroup 
                }
            }
        }
    }
    end {
        $context.Dispose()
        if($noproblems -eq $true) {
            Write-Host -ForegroundColor Green "Subsite $subsite created successfully!"
        }
    }
}


# Generate subsites in SharePoint Online by using a CSV file
Function SPS-ImportSubsitesFromCSV {
    Param(
        [Parameter(Mandatory=$true,HelpMessage="The URL of the site collection",Position=0)][ValidateNotNull()]
        [string]$SiteUrl,
        [Parameter(Mandatory=$true,HelpMessage="The CSV containing the subsites we want to create",Position=0)][ValidateNotNull()]
        [string]$PathToCSV
    )

    begin {


    }

    process {


    }

    end {

    }
} # End of function