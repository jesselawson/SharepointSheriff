![Sharepoint Sheriff Logo](sps-badge.png)

# Sharepoint Sheriff
A set of powershell tools to automate common and bulk tasks for those of us who have had Sharepoint Online administration dumped into our laps.

**Quicklinks:**
* [Create a single subsite](#create-a-new-subsite)
* [Create multiple subsites from a csv](#create-multiple-subsites-from-a-csv-file)
* [Get all subsites and export them to an HTML list](#get-an-html-formatted-list-of-all-subsites)

# Features
A set of cmdlets to be used in the Sharepoint Online Management Shell that will help you administer your institution's Sharepoint Online instance from the command line. These cmdlets have been designed with automation, efficiency, and hurry-up-and-go (HUG) administration in mind. 

# Installation 

## Prerequisites

Using Powershell to interface with Sharepoint Online requires you to have a copy of the [Sharepoint Online Management Shell][1] installed on your computer. 

## Downloading

Get yourself a copy of Sharepoint Sheriff by downloading the latest version. 

## Importing the cmdlets

As is normal in Powershell, you'll need to import the cmdlets into your script before you can use them. One easy way of doing this is to use the dot-source method. 

Simply add the `SPS-Suite.ps1` fullpath to the top of your script, like so:

```powershell
."C:\Users\Roosevelt\SPS-Suite.ps1"
write-host "I've included a Sharepoint Sheriff suite cmdlet!"

```

# Usage

Before you use the Sharepoint Sheriff suite, you'll need to bake in your credentials OR replace my email address with your own. 

Example:

```powershell
    
        # Get user authentication
        $Username = "your.email@yourcompany.org"
        # Alternative: Comment out the explicit declaration and uncomment the below
        # $Username = read-host -Prompt "Admin Email"

```

## Create a new Subsite

Sharepoint sites are organized under *Site Collections*. Sharepoint has a tool called `New-SPSite`, which is how you create a new collection from the commandline--but is **not** how you create a teamsite. 

It's very common for an organization to have *one* site collection, and then off of that have a number of teamsites as part of that collection. 

One way to organize a complex organization is to create parent teamsites off of the site collection, and then create subsites of those parent sites to account for the many teamsites you might have to deal with.

Here's a structural representation of what you can create, and what cmdlet you would use to create them:

| URL                                                        | Powershell cmdlet to create |
| ---------------------------------------------------------- | --------------------------- |
| https://yourorg.sharepoint.com/depts                       | New-SPSite                  |
| https://yourorg.sharepoint.com/depts/hr                    | SPS-CreateSubsite           |
| https://yourorg.sharepoint.com/depts/research              | SPS-CreateSubsite           |
| https://yourorg.sharepoint.com/depts/research/bravoteam    | SPS-CreateSubsite           |

So in the above examples, we have a **Site Collection** called *yourorg.sharepoint.com*. Then, we have a teamsite called *depts* (which we can create with SPS-CreateSubsite and just pass along the root url of the site collection), two subsites of *depts* called *research* and *hr* (each of which we can also create with SPS-CreateSubsite), and a subsite of *research* called *bravoteam*. Notice the pattern? As long as you have already created your site collection, you can create a subsite.  

The command to create a new subsite is as follows:

```powershell
SPS-CreateSubsite -SiteUrl "https://yoursharepoint.sharepoint.com/aPrimarySite" -SubsiteUrl "MySubsite" -Title "My Special Subsite of aPrimarySite" 
```

The above will create `https://yoursharepoint.sharepoint.com/aPrimarySite/MySubsite`. 

## Create multiple subsites from a CSV file

If you have many subsites you need to create on the fly, create a CSV file containing the subsite url and title of each one you wish to create, then use `SPS-CreateSubsitesFromCSV` to read the file and bulk create all the teamsites your heart desires. 

For example, let's say I have a site collection called `jessecollege.sharepoint.com`, and in that site collection I have created a parent teamsite called `Departments` that will be the root site of all my instructional departments:

```powershell
SPS-CreateSubsite -SiteUrl "https://jessecollege.sharepoint.com" `
-SubsiteUrl "Departments" -Title "Jesse College Instructional Departments"
```

With my new root site, I now want to generate about six dozen teamsites, one for each of my departments. When you do this, be sure you set the permissions for the `Departments` site to whatever you want each of its subsites to be because the permissions are going to be inherited. For example, when I create a parent site, I'll add our Sharepoint Administrator, Sharepoint Technical Team members, and our User Support personnel as owners so that each site created can be fully controlled (and ownership can be delegated to a staff member in the future by) our techs. 

I'll need to first create a CSV file of all the sites I want to create. It might look like this:

```
Url,Title
ag,Agriculture
addiction,Addiction Studies
anything,Anything Studies
bees,Beehive Replication Studies
git,Github Studies
law,Law and Policy
```

The format for this csv file is `headerless` and like so: `subsite-url,Subsite Title`. For this example, I've created a csv file called `department-subsites.csv`.

With that CSV file created, simply plug it into `SPS-CreateSubsitesFromCSV` like so:

```powershell
SPS-CreateSubsitesFromCSV -SiteUrl "https://yoursharepoint.sharepoint.com/departments" -PathToCSV "C:\Users\Roosevelt\Documents\SharepointMigration\department-subsites.csv"
```

# Get an HTML-formatted list of all subsites 

Let's say you have a teamsite called `TunaSandwichClub`, and each state in the United States has it's own Tuna Sandwich Club Statewide Conference Center subsite. 

Let's also say that your boss wants to have a page somewhere in Sharepoint that lists all the subsites of the TunaSandwichClub site. Are you going to make that list by hand? Hell no!

Simply use `SPS-GetTabledSubsites.ps1` like so:

```powershell
SPS-GetTabledSubsites.ps1 -SiteUrl "https://mysharepointsite.sharepoint.com/TunaSandwichClub" -OutFile "tunaclub-subsites.html"
```

When you're done, check out the output file for some copy+pastable HTML code that you can slap right into a Sharepoint page. Yeehaw! 


[1]: https://www.microsoft.com/en-us/download/details.aspx?id=35588