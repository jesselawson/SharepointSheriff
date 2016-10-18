
$AdminUsername = "lawsonje@butte.edu"

<#
    SPS-GetAdminPassword

    PURPOSE. 
        Provide a simple way to programatically embed different password collection styles based on the administrator using the script. Instead of having to 
        have a bunch of `$Password = read-host...` entries eveywhere OR having your plain-text password written multiple times in the same script, we can use this
        callback function to return our password capture method--either as a read-host or in plaintext to skip having to enter it in before every command.

    USAGE.
        $Password = SPS-GetAdminPassword

    NOTES.
        If you prefer to enter your admin password before every command, use the `read-host` line in the function. 
        If, however, you would prefer to just have your plain-text password baked into the scripts (as a lot of people find this more convenient), then simply 
            update this function to make it return your plain text password.

#>
Function SPS-GetAdminPassword {
    begin{
        
    }
    process {
        # Read in the password so we don't have to bake in the credentials. If you want, you can comment out the read-host line below and instead use the 
        # explicit password declaration if you're tired of typing in your password all the time.
        $Password = read-host -Prompt "Password for $AdminUsername" -AsSecureString
        #$Password = "PlainTextPassword"
    }
    end {
        return $Password
    }
}

<# 
    SPS-CreateSubsite

    PURPOSE.
    Create a single subsite. 

    USAGE.
        SPS-CreateSubsite -SiteUrl "https://<<your spo instance>>.sharepoint.com" `
        -SubsiteUrl "<<the-slug-of-the-subsite" -Title "The Subsite's Title" `
        -Template "Team" -SamePermissions $true
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
        $Username = $AdminUsername

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
        
        $Password = SPS-GetAdminPassword
        
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
               <# This should be un-needed, as we don't connect to <<spoinstance>>-admin.sharepoint.com. 
               try{
                    #Check to see if this is a root site collection that was passed
                    Get-SPOSite -Identity $SiteUrl
                }
                catch {
                #If here we are working with a subsite
                Write-Host "Oops, it looks like we're dealing with a subsite, I can handle that" -ForegroundColor Yellow 
                #reconstruct the SiteUrl parameter
                $sitefqdn = $SiteUrl.split("/")[2]
                $parentsite = "https://"+$sitefqdn+"/"
                $SiteUrl = $parentsite+$SiteUrl.Split("/")[3]+"/"+$SiteUrl.Split("/")[4]
                }#>
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
# This is a very opinionated function that makes a lot of assumptions for you:
#   
Function SPS-CreateSubsitesFromCSV {
    Param(
        [Parameter(Mandatory=$true,HelpMessage="The URL of the site collection",Position=0)][ValidateNotNull()]
        [string]$SiteUrl,
        [Parameter(Mandatory=$true,HelpMessage="The CSV containing the subsites we want to create",Position=1)][ValidateNotNull()]
        [string]$PathToCSV
    )

    begin {

        # Get the username and password   
        $Username = $AdminUsername
        $Password = SPS-GetAdminPassword
    
        # Check if csv file exists
        write-host "Checking CSV file..."
        $testpath = Test-Path -Path $PathToCSV    
    }

    process {
        if($testpath -eq $true) {
            # Import the CSV file
            write-host "Importing CSV file..."
            $subsites = Import-Csv $PathToCSV 

            write-host "Generating SPO context and credential..."
            $context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
            $context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Username, $Password)
            
            write-host "Beginning bulk creation..." 
            $numgood = 0
            $numbad = 0
            $success = $true
            foreach($subsite in $subsites) {
                $title = $subsite.Title
                $url = $subsite.Url
                write-host "Creating $title ($SiteUrl/$url)..." -NoNewline
                #Create SubSite
                $wci = New-Object Microsoft.SharePoint.Client.WebCreationInformation
                $wci.WebTemplate = "STS#0" # Note that I'm using a default Team site template here. Use whichever you want.
                $wci.UseSamePermissionsAsParentSite = $true
                $wci.Title = $subsite.Title
                $wci.Url = $subsite.Url
                $wci.Language = "1033"

                $SubWeb = $context.Web.Webs.Add($wci) 
                try {
                    $context.ExecuteQuery()
                }
                catch {
                    $success=$false
                    write-host "[ FAILED ]" -ForegroundColor Yellow
                    write-host ">> $_.Exception.Message" -ForegroundColor Red
                    $numbad += 1
                } 
                
                # Only write [OK] if success wasn't switched to false 
                if($success -eq $true) {
                    write-host "[ OK ]" -ForegroundColor Green
                    $numgood += 1
                }

                # Reset our success bool 
                $success = $true
            } # End foreach 

        } else {
            Write-Host "Oops! That CSV file doesn't seem to exist!"
        }
    }

    end {
        write-host "Finished with SPS-CreateSubsitesFromCSV. A total of $numgood subsites were created successfully."
        if($numbad -gt 0){
            write-host "Please note: $numbad subsites failed to be created!"
        }
    }
} # End of function

<#
    SPS-GetTabledSubsites.ps1

    PURPOSE.
        Given a teamsite, output a clean HTML table with the name and links to all that teamsite's subsites.

    USAGE.
        SPS-GetTabledSubsites --ParentSite "https://yourcollege.sharepoint.com/parentSite" --OutFile "parentSiteSubsites.html"

#>
  
Function SPS-GetTabledSubsites {
    Param(
        [Parameter(Mandatory=$true,HelpMessage="The URL of the site collection",Position=0)][ValidateNotNull()]
        [string]$SiteUrl,
        [Parameter(Mandatory=$true,HelpMessage="The name of the HTML file you wish to create",Position=1)][ValidateNotNull()]
        [string]$OutFile
    )

    begin {

        # Get the username and password   
        $Username = $AdminUsername
        $Password = SPS-GetAdminPassword

    }

    process {

        write-host "Generating SPO context and credential..."
        $context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
        $context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Username, $Password)
        
        write-host "Grabbing all the subsites for $SiteUrl... "
        
        $success = $true
        try {
            $context.Load($context.Web)
            $context.Load($context.Web.Webs)
            $context.ExecuteQuery()
        } catch {
            $success=$false
            write-host "[ FAILED ]" -ForegroundColor Yellow
            write-host ">> $_.Exception.Message" -ForegroundColor Red
        }

        # Only write [OK] if success wasn't switched to false 
        if($success -eq $true) {
            write-host "[ OK ]" -ForegroundColor Green
        }
        
        $outputhtml = ""

        write-host "Looping through subsites... "

        # Loop through the results and add the elements we want to a new PS Object
        for($i=0;$i -lt $context.Web.Webs.Count ;$i++) {

            $title = $context.Web.Webs[$i].Title
            $url = $context.Web.Webs[$i].Url

            write-host "Adding $title... " -NoNewline

            # Add this teamsite's information to the output html
            $outputhtml += "<p><a href='$url'>$title</a></p>"
            
            write-host "[ OK ]" -ForegroundColor Green
        }

        # Export to html file 
        $outputhtml > $OutFile

        write-host "All set!" 

    } # End of process{}
} # End of function