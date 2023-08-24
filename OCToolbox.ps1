# version 1.1
# 2023.08.24
# https://github.com/digitalParkour/OCToolbox

param (
# Optional Inputs (otherwise will be prompted when needed)
    [string]$baseUrl = '', # baseUrl/v1 - https://ordercloud.io/knowledge-base/ordercloud-supported-regions    
    [string]$token = '', # Copy Access Token from https://portal.ordercloud.io/console (expand Base URL/Region box) - Ensure you have selected the appropriate Api Client context for target marketplace.
# Use -verbose option to see request details and other useful output when troubleshooting an issue
    [switch]$verbose = $false
    )

#region Comments
    
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~
# PRE-REQS
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~
# Get PSYaml module
#  - Goto: https://github.com/Phil-Factor/PSYaml
#  - Copy files to %USERPROFILE%\Documents\WindowsPowerShell\Modules\PSYaml
# SET THIS TO YOUR PSYaml FOLDER LOCATION
$PSYamlFolder = "$env:USERPROFILE\Documents\WindowsPowerShell\Modules\PSYaml"

# ~~~~~~~~~~~~~~~~~~~~~~~~~~~
# CODE OUTLINE [CTRL+M]
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~
# - With Powershell ISE use CTRL+M to toggle expanding and collapsing all code regions

#endregion

# ////////////////////////////////////////////////////////////////////////////////////////////////
#
#region 1. DEFINITIONS
#
# ////////////////////////////////////////////////////////////////////////////////////////////////

enum OcType {
    Entity
    Assignment
}

# This defines an OC entity, capturing properties necessary for scripting, such as dependencies and parent types.
class OcDefn
{
    [ValidateNotNullOrEmpty()][string]$Name # Name as given in Seed export file
    [ValidateNotNullOrEmpty()][OcType]$Type
    [string]$ApiName                        # Slug used in api routes (when different than Name)
    [string[]]$ImportDependsOn              # Names of definitions that need to be imported before this one
    [string[]]$LivesUnder                   # Ordered list of parent entities needed when making API requests
}

# Ordered Object Dependencies, import runs from top to bottom to avoid 404 missing object issues
$defns = [OcDefn[]]@(
    [OcDefn]@{
        Name = 'Catalogs'
        Type = [OcType]::Entity
    },
    [OcDefn]@{
        Name = 'Buyers' 
        Type = [OcType]::Entity
        ImportDependsOn = @('Catalogs') # for default catalog id
    },
    [OcDefn]@{
        Name = 'Users'
        Type = [OcType]::Entity
        ImportDependsOn = @('Buyers')
        LivesUnder = @('Buyers')
    },
    [OcDefn]@{
        Name = 'UserGroups'
        Type = [OcType]::Entity
        ImportDependsOn = @('Buyers')
        LivesUnder = @('Buyers')
    },
    [OcDefn]@{
        Name = 'AdminUsers'
        Type = [OcType]::Entity
    },
    [OcDefn]@{
        Name = 'AdminUserGroups'
        Type = [OcType]::Entity
        ApiName = 'usergroups'
    },
    [OcDefn]@{
        Name = 'Suppliers'
        Type = [OcType]::Entity
    },
    [OcDefn]@{
        Name = 'SupplierUsers'
        Type = [OcType]::Entity
        ApiName = 'users'
        ImportDependsOn = @('Suppliers')
        LivesUnder = @('Suppliers')
    },
    [OcDefn]@{
        Name = 'SupplierUserGroups'
        Type = [OcType]::Entity
        ApiName = 'usergroups'
        ImportDependsOn = @('Suppliers')
        LivesUnder = @('Suppliers')
    },
    [OcDefn]@{
        Name = 'IntegrationEvents'
        Type = [OcType]::Entity
    },
    [OcDefn]@{
        Name = 'ApiClients'
        Type = [OcType]::Entity
        ImportDependsOn = @('IntegrationEvents','AdminUsers','Users') # for default user
    },
    [OcDefn]@{
        Name = 'SecurityProfiles'
        Type = [OcType]::Entity
    },
    [OcDefn]@{
        Name = 'ImpersonationConfigs'
        Type = [OcType]::Entity
        ApiName = 'impersonationconfig'
        ImportDependsOn = @('ApiClients')
    },
    [OcDefn]@{
        Name = 'OpenIdConnects'
        Type = [OcType]::Entity
        ImportDependsOn = @('ApiClients', 'IntegrationEvents')
    },
    [OcDefn]@{
        Name = 'AdminAddresses'
        Type = [OcType]::Entity
        ApiName = 'addresses'
    },
    [OcDefn]@{
        Name = 'MessageSenders'
        Type = [OcType]::Entity
    },
    [OcDefn]@{
        Name = 'Locales'
        Type = [OcType]::Entity
    },
    [OcDefn]@{
        Name = 'Webhooks'
        Type = [OcType]::Entity
    },
    [OcDefn]@{
        Name = 'XpIndices'
        Type = [OcType]::Entity
    },
    [OcDefn]@{
        Name = 'Categories'
        Type = [OcType]::Entity
        ImportDependsOn = @('Catalogs')
        LivesUnder = @('Catalogs')
    },
    [OcDefn]@{
        Name = 'PriceSchedules'
        Type = [OcType]::Entity
    },
    [OcDefn]@{
        Name = 'Promotions'    
        Type = [OcType]::Entity
    },
    [OcDefn]@{
        Name = 'SupplierAddresses'
        Type = [OcType]::Entity
        ApiName = 'addresses'
        ImportDependsOn = @('Suppliers')
        LivesUnder = @('Suppliers')
    },
    [OcDefn]@{
        Name = 'Addresses'
        Type = [OcType]::Entity
        ImportDependsOn = @('Buyers')
        LivesUnder = @('Buyers')
    },
    [OcDefn]@{
        Name = 'Specs'
        Type = [OcType]::Entity
    },
    [OcDefn]@{
        Name = 'SpecOptions' # depends on specs
        Type = [OcType]::Entity
        ImportDependsOn = @('Specs')
        LivesUnder = @('Specs')
    },
    [OcDefn]@{
        Name = 'Products' # May depend on supplier (DefaultSupplierID), May depend on Address (ShipFromAddressID), May depend on PriceSchedule (DefaultPriceScheduleID)
        Type = [OcType]::Entity
        ImportDependsOn = @('Suppliers', 'Addresses', 'PriceSchedules')
    },
    @{
        Name = 'Variants' # depends on specs, specOptions, products, and also SpecProductAssignments
        Type = [OcType]::Entity
        ImportDependsOn = @('Specs', 'SpecOptions', 'Products', 'SpecProductAssignments')
        LivesUnder = @('Products')
    },
    [OcDefn]@{
        Name = 'InventoryRecords'
        Type = [OcType]::Entity
        ImportDependsOn = @('Products')
        LivesUnder = @('Products')
    },
    [OcDefn]@{
        Name = 'VariantInventoryRecords'
        Type = [OcType]::Entity
        ApiName = 'inventoryrecords'
        ImportDependsOn = @('Products','Variants')
        LivesUnder = @('Products','Variants')
    },
    [OcDefn]@{
        Name = 'Incrementors'
        Type = [OcType]::Entity
    },
    [OcDefn]@{
        Name = 'CostCenters'
        Type = [OcType]::Entity
        ImportDependsOn = @('Buyers')
        LivesUnder = @('Buyers')
    },
    [OcDefn]@{
        Name = 'CreditCards'
        Type = [OcType]::Entity
        ImportDependsOn = @('Buyers')
        LivesUnder = @('Buyers')
    },
    [OcDefn]@{
        Name = 'SpendingAccounts'
        Type = [OcType]::Entity
        ImportDependsOn = @('Buyers')
        LivesUnder = @('Buyers')
    },
    [OcDefn]@{
        Name = 'ApprovalRules'
        Type = [OcType]::Entity
        ImportDependsOn = @('Buyers')
        LivesUnder = @('Buyers')
    },
    [OcDefn]@{
        Name = 'ProductFacets'
        Type = [OcType]::Entity
    },
    [OcDefn]@{
        Name = 'SellerApprovalRules'
        Type = [OcType]::Entity
        ApiName = 'approvalrules'
    },
    # All Assignment types can be assumed to depend on all Entity types
    [OcDefn]@{
        Name = 'SecurityProfileAssignments'
        Type = [OcType]::Assignment
        ApiName = 'securityprofiles/assignments'
    },
    [OcDefn]@{
        Name = 'CatalogAssignments'
        Type = [OcType]::Assignment
        ApiName = 'catalogs/assignments'
    },
    [OcDefn]@{
        Name = 'PromotionAssignments'
        Type = [OcType]::Assignment
        ApiName = 'promotions/assignments'
    },
    [OcDefn]@{
        Name = 'ProductCatalogAssignment'
        Type = [OcType]::Assignment
        ApiName = 'catalogs/productassignments'
    },
    [OcDefn]@{
        Name = 'CategoryProductAssignments'
        Type = [OcType]::Assignment
        ApiName = 'categories/productassignments'
        LivesUnder = @('Catalogs')
    },
    [OcDefn]@{
        Name = 'SpecProductAssignments'
        Type = [OcType]::Assignment
        ApiName = 'specs/productassignments'
    },
    [OcDefn]@{
        Name = 'UserGroupAssignments'
        Type = [OcType]::Assignment
        ApiName = 'usergroups/assignments'
        LivesUnder = @('Buyers')
    },
    [OcDefn]@{
        Name = 'SpendingAccountAssignments'
        Type = [OcType]::Assignment
        ApiName = 'spendingaccounts/assignments'
        LivesUnder = @('Buyers')
    },
    [OcDefn]@{
        Name = 'AddressAssignments'
        Type = [OcType]::Assignment
        ApiName = 'addresses/assignments'
        LivesUnder = @('Buyers')
    },
    [OcDefn]@{
        Name = 'CostCenterAssignments'
        Type = [OcType]::Assignment
        ApiName = 'costcenters/assignments'
        LivesUnder = @('Buyers')
    },
    [OcDefn]@{
        Name = 'CreditCardAssignments'
        Type = [OcType]::Assignment
        ApiName = 'creditcards/assignments'
        LivesUnder = @('Buyers')
    },
    [OcDefn]@{
        Name = 'CategoryAssignments'
        Type = [OcType]::Assignment
        ApiName = 'categories/assignments'
        LivesUnder = @('Catalogs')
    },
    [OcDefn]@{
        Name = 'ProductSupplierAssignments'
        Type = [OcType]::Assignment
        ApiName = 'suppliers'
        LivesUnder = @('Products')
    },
    [OcDefn]@{
        Name = 'AdminUserGroupAssignments'
        Type = [OcType]::Assignment
        ApiName = 'usergroups/assignments'
    },
    [OcDefn]@{
        Name = 'ApiClientAssignments'
        Type = [OcType]::Assignment
        ApiName = 'apiclients/assignments'
    },
    [OcDefn]@{
        Name = 'LocaleAssignments'
        Type = [OcType]::Assignment
        ApiName = 'locales/assignments'
    },
    [OcDefn]@{
        Name = 'ProductAssignments'
        Type = [OcType]::Assignment
        ApiName = 'products/assignments'
    }
    [OcDefn]@{
        Name = 'SupplierUserGroupsAssignments'
        Type = [OcType]::Assignment
        ApiName = 'usergroups/assignments'
        LivesUnder = @('Suppliers')
    }
    [OcDefn]@{
        Name = 'SupplierBuyerAssignments'
        Type = [OcType]::Assignment
        ApiName = '/buyers'
        LivesUnder = @('Suppliers')
    }
 )

#endregion

# ////////////////////////////////////////////////////////////////////////////////////////////////
#
#region     1.b CONTEXT
#
# ////////////////////////////////////////////////////////////////////////////////////////////////

function Set-AccessToken
{
    do
    {
        Write-Host "Need Access Token"
        Write-Host " 1. Go to https://portal.ordercloud.io/console (CTRL+Click)" -ForegroundColor DarkGray
        Write-Host " 2. Set your context to the appropriate marketplace" -ForegroundColor DarkGray
        Write-Host " 3. Expand Base URL/Region box and copy Access Token" -ForegroundColor DarkGray
        $global:token = Read-Host -Prompt 'Paste OrderCloud Access Token'
    } until ($global:token)
}

function Set-BaseUrl
{
    do
    {        
        $endpoints = [ordered]@{
            'Sandbox - US West' = 'https://sandboxapi.ordercloud.io/v1';
            'Sandbox - US East' = 'https://useast-sandbox.ordercloud.io/v1';
            'Sandbox - Australia East' = 'https://australiaeast-sandbox.ordercloud.io/v1';
            'Sandbox - Europe West' = 'https://westeurope-sandbox.ordercloud.io/v1';
            'Sandbox - Japan East' = 'https://japaneast-sandbox.ordercloud.io/v1';
            'Staging - US West' = 'https://stagingapi.ordercloud.io/v1';
            'Staging - US East' = 'https://useast-staging.ordercloud.io/v1';
            'Staging - Australia East' = 'https://australiaeast-staging.ordercloud.io/v1';
            'Staging - Europe West' = 'https://westeurope-staging.ordercloud.io/v1';
            'Staging - Japan East' = 'https://japaneast-staging.ordercloud.io/v1';
            'Production - US West' = 'https://api.ordercloud.io/v1';
            'Production - US East' = 'https://useast-production.ordercloud.io/v1';
            'Production - Australia East' = 'https://australiaeast-production.ordercloud.io/v1';
            'Production - Europe West' = 'https://westeurope-production.ordercloud.io/v1';
            'Production - Japan East' = 'https://japaneast-production.ordercloud.io/v1';
        }

        $choice = Get-FromChoiceMenu "Specify OrderCloud Base URL" @($endpoints.Keys)
        $global:endpoint = $choice
        $global:baseUrl = $endpoints[$choice];
    } until ($global:baseUrl)  
}

function Get-MarketplaceID
{
    # Verify access, and show Marketplace
    $check = Invoke-OcRequest -request $([OcRequest]@{verb = "GET"; url = "/me"}) -forceRefresh $false;
    
    if($check -and $check.Seller.ID)
    {
        return $check.Seller.ID
    } else {
        return $null
    }
}

# Show context
function Show-Context
{
    $account = Get-MarketplaceID
    Write-Host
    Write-Host "Context:" -ForegroundColor Cyan
    Write-Host "Endpoint: $($global:endpoint) $($global:baseUrl)" -ForegroundColor Yellow
    if($account)
    {
        Write-Host "Marketplace: $account" -ForegroundColor Yellow
    } else {
        Write-Host "Marketplace: NEEDS TO BE SET" -ForegroundColor Red
    }
}

function Invoke-ConfirmContext
{
    $confirmed = $false
    do {
    
        if($global:baseUrl -and $global:token)
        {
            Show-Context
        }
                 
        $confirmed = $($global:baseUrl -and $global:token -and 0 -eq $host.ui.PromptForChoice(
            '',
            'Continue?',
            @(
                [System.Management.Automation.Host.ChoiceDescription]::new("&Yes", "Yes")
                [System.Management.Automation.Host.ChoiceDescription]::new("&No", "No")
            ),
            0))

        if(-not $confirmed)
        {
            if(-not $global:baseUrl -or (0 -ne $host.ui.PromptForChoice(
                "Endpoint",
                "Confirm endpoint: $($global:endpoint) $($global:baseUrl)",
                @(
                    [System.Management.Automation.Host.ChoiceDescription]::new("&Yes", "Yes")
                    [System.Management.Automation.Host.ChoiceDescription]::new("&No", "No")
                ),
                0)))
            {
                Set-BaseUrl
            }

            $marketplaceId = Get-MarketplaceID
            if(-not $marketplaceId)
            {
                Write-Host
                Write-Host "Marketplace"
            }
            if(-not $marketplaceId -or (0 -ne $host.ui.PromptForChoice(
                "Marketplace",
                "Confirm marketplace: $(Get-MarketplaceID)",
                @(
                    [System.Management.Automation.Host.ChoiceDescription]::new("&Yes", "Yes")
                    [System.Management.Automation.Host.ChoiceDescription]::new("&No", "No")
                ),
                0)))
            {
                Set-AccessToken
            }
        }
        
    } until($confirmed)
}

#endregion

# ////////////////////////////////////////////////////////////////////////////////////////////////
#
#region 2. OC REQUESTS
#
# ////////////////////////////////////////////////////////////////////////////////////////////////

# Convenience object to pass details of a REST request
class OcRequest
{
    [ValidateNotNullOrEmpty()][string]$verb
    [ValidateNotNullOrEmpty()][string]$url
    [string]$body
}

# Invoke OC REST Requst, ensure valid access token
function Invoke-OcRequest([OcRequest]$request, [bool]$forceRefresh = $true)
{   
    do {
        if(-not $global:token)
        {
            if($forceRefresh)
            {
                Set-AccessToken
            } else {
                return $null
            }
        }

        $headers = @{
        'Authorization' = "Bearer $($global:token)"
        'Content-Type' = 'application/json'
    }

        $url = "$($global:baseUrl)$($request.url)"
        if($global:verbose)
        {
            Write-Host
            Write-Host "REQUEST:" $url -ForegroundColor Gray
            # Write-Host "HEADERS:" $headers -ForegroundColor Gray
            Write-Host "BODY:" $request.body -ForegroundColor Gray
        }

        try {
            $response = Invoke-RestMethod -Headers $headers -Method $request.verb -Uri $url -Body $request.body -Verbose:$global:verbose -ContentType "application/json; charset=utf-8"
                
            if($global:verbose)
            {
                #Write-Host "RESPONSE:" $($response | ConvertTo-Json) -ForegroundColor Green
            }
            return $response;
        } catch {
            if($_.Exception.Response.StatusCode -eq 401)
            {
                $global:token = $null; # clearing token will force prompt next cycle
                Write-Host "Access Token likely expired. Provide a fresh token to retry and continue." -ForegroundColor Magenta
            } else
            {
                Write-Host "RESPONSE:" $($_.Exception.Response | ConvertTo-Json) -ForegroundColor Yellow
                if($_.Exception.Response.Headers -and $_.Exception.Response.Headers['X-oc-logid'])
                {
                    Write-Host "x-OC-LogId:" $_.Exception.Response.Headers['X-oc-logid'] -ForegroundColor Yellow
                }

                if($_.Exception.Response.StatusCode -eq 500)
                {
                    $Reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                    $Reader.BaseStream.Position = 0
                    $Reader.DiscardBufferedData()
                    $ResponseBody = $Reader.ReadToEnd()

                    Write-Host "BODY:" $ResponseBody -ForegroundColor Yellow
                }

                throw
            }
        }
    }
    until($global:token)
}

# ------------------------

# Resolve API route slug to use for $defn. Use $defn.ApiName when provided else fall back to $defn.Name (as most match the name).
function Get-ApiName ([OcDefn]$defn)
{
    if(-not $defn)
    {
        return $null
    }

    $apiName = $defn.Name.ToLowerInvariant()
    if($defn.ApiName)
    {
        $apiName = $defn.ApiName
    }
    return $apiName
}

# Package ids used in slug for easy access
class DependentSlug
{
    [string]$depSlug = ""
    [string[]]$depIds
}

# returns DependentSlug[] expanding dependent slug and collecting parallel array of ids for all permutations
function Get-DependentSlugs([string[]]$deps, [DependentSlug]$baseObj = $null)
{
    $depSlugs = [DependentSlug[]]@()

    $first, $rest = $deps
    
    if(-not $first)
    {
        return $depSlugs
    }

    $depDefn = $defns | Where-Object { $_.Name -eq $first }

    $depSlug = Get-ApiName $depDefn
    $depIds = Get-EntityIds -defn $depDefn -baseObj $baseObj # Let the inception begin!

    foreach($depId in $depIds)
    {
        $depObj = $([DependentSlug]@{
            depSlug = ""
            depIds = @()
        })
        if($baseObj)
        {
            $depObj.depSlug = $baseObj.depSlug
            $depObj.depIds = @($baseObj.depIds)
        }

        $depObj.depSlug = "$($depObj.depSlug)/$($depSlug)/$($depId)"
        $depObj.depIds = @($depObj.depIds) + @($depId)

        if($rest)
        {
            $crawlSlugs = Get-DependentSlugs -deps @($rest) -baseObj $depObj # and melt your brain ;)
            $depSlugs = $depSlugs + $crawlSlugs
        } else {
            $depSlugs = $depSlugs + @($depObj)
        }
    }

    return $depSlugs;
}


# Expand $request to collect result items form all pages. Assumes GET style request with no body.
function Invoke-OcRequestAll([OcRequest] $request)
{
    $pageSize = 100
    $items = @()
    $hasNextPage = $false
    
    $page = 0
    do {
        $page++
        $pagedRequest = [OcRequest]@{
            verb = $request.verb
            url = "$($request.url)?pageSize=$pageSize&page=$page"
        }

        $response = Invoke-OcRequest $pagedRequest
        
        $items = $items + $response.Items

        $hasNextPage = $response.Meta.Page -lt $response.Meta.TotalPages
        
        if($response.Meta.TotalPages -gt 1 -and $global:verbose){
            Write-Host "Queried page $page of $($response.Meta.TotalPages), $($response.Items.Length) items" -ForegroundColor DarkGray
        }
    } while ($hasNextPage)

    return $items
}

# Returns response.Items (across all pages) for given $defn type.
# Caches results for quick reuse on the same tool (heavily used for reporting tool on nested objects).
function Get-ListRequest([OcDefn]$defn)
{
    $results = Get-Cache $defn.Name
    if($results)
    {
        return $results;
    }

    $results = @()
    
    $slug = Get-ApiName $defn
    
    # Support requests with multiple IDs
    if($defn.LivesUnder)
    {
        $deps= Get-DependentSlugs -deps $defn.LivesUnder
        foreach($dep in $deps)
        {
            $request = @{
                verb = "GET"
                url = "$($dep.depSlug)/$($slug)"
            };

            $items = Invoke-OcRequestAll $request
            if($items)
            {
                foreach($item in $items)
                {
                    $item | Add-Member -NotePropertyName AncestorIDs -NotePropertyValue @($dep.depIds)
                }
                $results = $results + $items
            }
        }
    } else {
    
        $request = @{
            verb = "GET"
            url = "/$($slug)"
        };

        $results = Invoke-OcRequestAll $request

    }

    Set-Cache $defn.Name $results
    return $results
}

# Returns all IDs of given entity
# Usually called with only -defn input
# Internally is called with $baseObj to drill out nested API routes
function Get-EntityIds([OcDefn]$defn, [DependentSlug]$baseObj = $null)
{
    if($global:verbose)
    {
        Write-Host "Getting $($defn.Name) ids..."
    }
    $list = Get-ListRequest -defn $defn
    
    # refine list when crawling parent paths
    if($baseObj)
    {
        $list = $list | where { $($_.AncestorIDs | ?{ $baseObj.depIds -contains $_}) }
    }

    if($global:verbose)
    {
        $ttl = 0
        if($list -and $list.Length -gt 0)
        {
            $ttl = $list.Length
        }
        Write-Host "Found" $ttl $defn.Name "IDs"
    }
    
    if($defn.Name -eq 'XpIndices')
    {        
        return @($list | % {"$($_.ThingType)/$($_.Key)"})
    }
    return @($list | % {$_.ID})
}

#endregion

# ////////////////////////////////////////////////////////////////////////////////////////////////
#
#region 3. CACHING UTIL
#
#   - Save lists of OC entities to avoid redundant queries
#
# ////////////////////////////////////////////////////////////////////////////////////////////////

$cachePrefix = 'occache_'

# $global variables remain until powershell terminal session is closed or removed by Clear-Cache (below)
function Set-Cache($name,$value){
    if(-not $value)
    { return }

    Set-Variable -Scope Global -Name "$($cachePrefix)$($name)" -Value $value
    
    if($global:verbose)
    {
        Write-Host "set cached result for $name" -ForegroundColor DarkGreen
    }

    # return $value
}

function Get-Cache($name){
   $cachedResults = Get-Variable -Scope Global -Name "$($cachePrefix)$($name)" -ErrorAction SilentlyContinue -ValueOnly
   if($cachedResults){   
        if($global:verbose)
        {
            Write-Host "found cached result for $name" -ForegroundColor DarkGreen
        }
   }
   return $cachedResults
}

function Clear-Cache($name){
    if($global:verbose)
    {
        Write-Host "Clearing Cache: $name" -ForegroundColor DarkGreen
    }
    Remove-Variable -Scope Global -Name "$($cachePrefix)$($name)" -ErrorAction SilentlyContinue
}

# Find and clear all $global variables with given $cachePrefix
function Clear-AllCache {
    $keys = Get-Variable -Scope Global | where {$_.Name -like "$($cachePrefix)*"} | % {$_.Name.Replace($cachePrefix,'')}
    
    if($global:verbose -and $keys -and $keys.Length -gt 0)
    {
        Write-Host "Clearing All Cache Keys: $($keys -join ',')..." -ForegroundColor DarkGreen
    }
    foreach($key in $keys)
    {
        Clear-Cache $key
    }
}

#endregion

# ////////////////////////////////////////////////////////////////////////////////////////////////
#
#region 4. DISPLAY FUNCTIONS
#
# ////////////////////////////////////////////////////////////////////////////////////////////////

# Display automated commands to user consistently to show what script is doing
function Show-Command ($cmd)
{
    $line = "".PadRight($cmd.Length + 11,'-')
    Write-Host "|$line|" -ForegroundColor Cyan
    Write-Host "| COMMAND: $cmd |" -ForegroundColor Cyan
    Write-Host "|$line|" -ForegroundColor Cyan
}

# Splash screen
function Show-Logo {
    $logo1 = @"
       .                                                                        
       *#,                                                                      
        /##/                                                                    
.        ,####.                                                                 
 /#/       /####(                                                               
  *####*     /#####*                                                            
    (######,   .##%#%%/                                                 
      *########/.  ,#%#%##.                                                     
         /########%#/.  ,/#%#*                                                  
             *####%#%#%%%#*.  ,(#*                                              
       ,###*       ,#%%%%%%%%%%%#/                                     
"@
    $logo2 = @"                 
         .(######(/*.    .,*(%%%%%%%%%%%#(/,.                                   
            ,(######%%%%%%%#(/,...  .*/(##%%%%%%%%&%(, .(%&&%&&&&&&&%(,         
"@
    $logo3 = @"
                ./##%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%&%(  /&&&&&&#.,%&%       
                         ..,,**//(((#######((///((#%%%&%&&%&#. (&&&&%&&&&&&&&#. 
                    /#%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%&%%%%&%&&, (&&%&&&&&%&&%. 
                          ,(%%%%%%%%%%%%%%%%%%%%%&%%%&%%%&%%&&%&&. %&%,     .   
                                                     .,(&&&&%&&&&&* %           
                                                           *%&&%&&&/            
                               Americaneagle.com              ,%%&%&            
                                                                 #&&(           
                                OC TOOLBOX v1.0                   ,%%           
                                                                   .#           

"@
    Write-Host $logo1 -ForegroundColor Red
    Write-Host $logo2 -ForegroundColor White
    Write-Host $logo3 -ForegroundColor Cyan
}

function Show-ToolBox
{
    # Thanks to https://ascii.co.uk/art/toolbox
    Write-Host "          ___________  ______________________________" -ForegroundColor Red
    Write-Host "       .'           .'                              .'" -ForegroundColor Red
    Write-Host "    .'           .'                              .'  |" -ForegroundColor Red
    Write-Host " .'___________.'______________________________.'     |" -ForegroundColor Red
    Write-Host " |.----------.|.-----___-----.|.-----___-----.|      |" -ForegroundColor Red
    Write-Host " |]          |||_____________|||_____________||      |" -ForegroundColor Red
    Write-Host " ||          ||.-----___-----.|.-----___-----.|      |" -ForegroundColor Red
    Write-Host " ||          |||_____________|||_____________||      |" -ForegroundColor Red
    Write-Host " ||          ||.-----___-----.|.-----___-----.|      |" -ForegroundColor Red
    Write-Host " |]         o|||_____________|||_____________||      |" -ForegroundColor Red
    Write-Host " ||          ||.-----___-----.|.-----___-----.|      |" -ForegroundColor Red
    Write-Host " ||          |||             |||_____________||      |" -ForegroundColor Red
    Write-Host " ||          |||             ||.-----___-----.|     .'" -ForegroundColor Red
    Write-Host " |]          |||             |||             ||  .' o)" -ForegroundColor Red
    Write-Host " ||__________|||_____________|||_____________||'" -ForegroundColor Red
    Write-Host " ''----------'''-----------------------------''" -ForegroundColor Red
    Write-Host "  (o)                                      (o)" -ForegroundColor Red
}

#endregion

# ////////////////////////////////////////////////////////////////////////////////////////////////
#
#region     4.b Get-FromChoiceMenu
#
#    ie. Get-FromChoiceMenu "Which Color?" @('Red','Green','Blue', 'Yellow', 'White')
#    ie. Get-FromChoiceMenu "Which Colors?" @('Red','Green','Blue', 'Yellow', 'White') -allowMultiple -defaults @('Blue','White')
#
# ////////////////////////////////////////////////////////////////////////////////////////////////

# Returns user selection from list of options
# If -allowMultiple, then array will be returned
function Get-FromChoiceMenu ($title, $options, [switch]$allowMultiple = $false, $defaults = @())
{ 
    # Show warning, ISE doesn't let us control the cursor position
    if($Host.Name -like '*ISE*')
    {    
        Write-Host "ISE doesn't allow control of the cursor position, needed to draw the menu." -ForegroundColor Yellow
        throw "Run this in a powershell terminal outside of ISE"
    }

    # ////////////////////////////////////////////////////////////////////////////////////////////////
    #
    #  HELPERS SECTION
    #
    # ////////////////////////////////////////////////////////////////////////////////////////////////

    # Options can be dran individually, given proper cursor handling for performance
    function Write-Option {
        param (
            [string]$option,
            [switch]$highlighted = $false,
            [switch]$summary = $false
         )
        $color = "Cyan"
        if($highlighted)
        {
            $color = "White"
        } elseif($summary)
        {
            $color = "DarkGray"
        }
        
        if($highlighted) {
            Write-Host " > " -ForegroundColor $color -NoNewline
        } else {
            Write-Host "   " -ForegroundColor $color -NoNewline
        }
                
        if($allSelected -or ($selectedOptions -and $selectedOptions.Contains($option)))
        {
            if($summary) { $color = "White" }
            Write-Host "● " -ForegroundColor $color -NoNewline
        } elseif(-not $allowMultiple -or $summary) {
            Write-Host "  " -ForegroundColor $color -NoNewline
        } else {
            Write-Host "○ " -ForegroundColor $color -NoNewline
        }

        Write-Host $option.PadRight(50) -ForegroundColor $color # pad spaces to the right to overwrite old lines on weird scrolling case
    }

    # Master Draw function
    # Supports two modes
    #   Default - Interactive, fits as many options up to max window height keeping instructions in view, cursor is set to first option
    #   Summary - non-interactive, lists all options pushing the scroll, cursor left at end
    function Draw-Menu
    {
        param (
            [switch]$summary = $false,
            [int]$highlight = -1, # currently selected option index
            [switch]$firstDraw = $true # pass flag for initial draw to correct for offset
         )

        $offset = 0        
        $height = $instructions.length + $options.Length + 1 # +1 for title

        if(-not $firstDraw)
        {
            $offset = $height - ($options.Length - $selectedIndex)
        }
        [Console]::SetCursorPosition(0, (0,($host.Ui.RawUI.CursorPosition.Y - $offset) | Measure -Max).Maximum)
                
        # Write Question
        Write-Host "? " -NoNewline -ForegroundColor DarkGray
        Write-Host $title.PadRight(50)  -ForegroundColor White

        # Draw Instructions
        foreach ($line in $instructions)
        {
            Write-Host $line.PadRight(50) -ForegroundColor DarkGray
        }
        # Draw Options
        $count = 0;
        # Protect against small screens/long lists
        $maxLines = $Host.UI.RawUI.WindowSize.Height - $instructions.Length - 2 # -2 for title and height being one larger than actual
        foreach ($option in $options) {
            if(-not $summary -and $count -ge $maxLines)
            {
                break
            }
            Write-Option "$option" -summary:$summary -highlighted:$($count -eq $highlight)
            $count++
        }
        
        if(-not $summary)
        {
            [Console]::SetCursorPosition(0, (0,($host.Ui.RawUI.CursorPosition.Y - $count) | Measure -Max).Maximum)
        }
    }
    
    # ////////////////////////////////////////////////////////////////////////////////////////////////
    #
    #  INIT
    #
    # ////////////////////////////////////////////////////////////////////////////////////////////////
    
    # Instructions
    $instructions = @(
    "   Use the [Up/Down] arrows to navigate,")
    if($allowMultiple)
    {
        $instructions +=
    "           [Spacebar] to toggle selected,"
    }
    $instructions +=
    "       and [Enter] to complete."    
    $instructions += ""
    

    $selectedOptions = @()
    if($allowMultiple)
    {
        $options = @("All") + $options
        $selectedOptions = $defaults
    }    

    $allSelected = $false

    Write-Host

    # ////////////////////////////////////////////////////////////////////////////////////////////////
    #
    #  DO SELECTION
    #
    # ////////////////////////////////////////////////////////////////////////////////////////////////
    
    # Confirmation loop
    do {
        $selectedIndex = 0
        $redraw = $true
        $offset = 0
        $isFirstDraw = $true

        # Manage cursor and display until [Enter] is pressed
        $menu_active = $true
        :selectLoop while ($menu_active) {
         
            if($redraw)
            {
              Draw-Menu -highlight $selectedIndex -firstDraw:$isFirstDraw
              if($isFirstDraw)
              {
                $isFirstDraw = $false
              }      

              $redraw = $false    
            }

            if ([console]::KeyAvailable) {
                $key = $Host.UI.RawUI.ReadKey()

                # Move
                switch ($key.VirtualKeyCode) { 
                    38 {
                        #down key
                        if ($selectedIndex -gt 0) {                        
                            Write-Option $options[$selectedIndex]
                            
                            $selectedIndex = $selectedIndex - 1

                            [Console]::SetCursorPosition(0, (0,($host.Ui.RawUI.CursorPosition.Y - 2) | Measure -Max).Maximum)
                            Write-Option $options[$selectedIndex] -highlighted
                            [Console]::SetCursorPosition(0, (0,($host.Ui.RawUI.CursorPosition.Y - 1) | Measure -Max).Maximum)
                        } else {
                            continue :selectLoop;
                        }
                    }
 
                    40 {
                        #up key
                        if ($selectedIndex -lt $options.Length - 1) {                        
                            Write-Option $options[$selectedIndex]
                            
                            $selectedIndex = $selectedIndex + 1

                            Write-Option $options[$selectedIndex] -highlighted
                            [Console]::SetCursorPosition(0, (0,($host.Ui.RawUI.CursorPosition.Y - 1) | Measure -Max).Maximum)
                        } else {
                            continue :selectLoop;
                        }
                    }
                    32 {
                        #space bar
                        if ($selectedOptions -and $selectedOptions.Contains($options[$selectedIndex])) {
                            # toggle remove:
                            if($options[$selectedIndex] -eq "All")
                            {
                                $allSelected = $false
                                $selectedOptions = @()
                                
                                $redraw = $true
                            } elseif ($allSelected)
                            {
                                $allSelected = $false
                                $selectedOptions = @($options | Where-Object { $_ –ne "All" -and $_ –ne $options[$selectedIndex] })
                            } else {                            
                                $selectedOptions = @($selectedOptions | Where-Object { $_ –ne $options[$selectedIndex] })
                            }
                        } else {
                            # toggle add:
                            if($allowMultiple)
                            {
                                if($options[$selectedIndex] -eq "All")
                                {
                                    $allSelected = $true
                                    $selectedOptions = $options
                                    
                                    $redraw = $true
                                } else {
                                    $selectedOptions += $options[$selectedIndex]                                
                                }
                            } else {
                                $selectedOptions = @($options[$selectedIndex])
                            }
                        }
                        
                        [Console]::SetCursorPosition(0, $host.Ui.RawUI.CursorPosition.Y)
                        Write-Option $options[$selectedIndex] -highlighted                        
                        [Console]::SetCursorPosition(0, (0,($host.Ui.RawUI.CursorPosition.Y - 1) | Measure -Max).Maximum)
                    }
                    13 {
                        #enter key
                        if($allowMultiple)
                        {
                            if($selectedOptions)
                            {
                              $menu_active = $false
                            }
                        } else {
                            $selectedOptions = @($options[$selectedIndex])
                            $menu_active = $false
                        }
                    }
                }
            }
            Start-Sleep -Milliseconds 5 #Prevents CPU usage from spiking while looping
        }

        Draw-Menu -summary

        # Remove "All" from selected list if chosen
        if($allSelected)
        {
            $selectedOptions = @($options | Where-Object { $_ –ne "All" })
        } else {
        # Since user can select in random order, sort selected to match order provided with options list
            $sortedOptions = @()
            Foreach($option in $options)
            {
                if($option -ne "All" -and $selectedOptions -and $selectedOptions.Contains($option))
                {
                    $sortedOptions += $option
                }
            }
            $selectedOptions = $sortedOptions
        }
    }
    # Confirmation loop prompt
    Until(-not $allowMultiple -or 0 -eq $host.ui.PromptForChoice(
        "Confirm Selection?",
        ' ' + $($selectedOptions -join ', '),
        @(
            [System.Management.Automation.Host.ChoiceDescription]::new("&Yes", "Continue with selected")
            [System.Management.Automation.Host.ChoiceDescription]::new("&No", "Edit selections")
        ),
        0))

    Write-Host
    return $selectedOptions
}

#endregion

# ////////////////////////////////////////////////////////////////////////////////////////////////
#
#region 5. EXPORT PROGRAM
#
# ////////////////////////////////////////////////////////////////////////////////////////////////

function Invoke-ExportProgram
{    
    Write-Host "TOOL: Export Seed File"
    Write-Host
    Write-Host "We need 4 parameters to generate a seed file." -ForegroundColor Cyan
    
    # Gather inputs
    $current = Get-MarketplaceID
    if($current)
    {
        $marketplaceId = Read-Host -Prompt " 1. Source Marketplace ID (enter for current, $current)"
        if(-not $marketplaceId)
        {
            $marketplaceId = $current
        }
    } else {
        $marketplaceId = Read-Host -Prompt ' 1. Source Marketplace ID'
    }
    $defaultName = 'oc-export-seed'
    $outFile = Read-Host -Prompt " 2. Export name (enter for default, $defaultName)"
    $user = Read-Host -Prompt ' 3. OC Portal Username'
    $pass = Read-Host -Prompt ' 4. OC Portal Password'

    # Hide password from screen
    [Console]::SetCursorPosition(24, $host.Ui.RawUI.CursorPosition.Y - 1)
    Write-Host "*****".PadRight($pass.Length," ")

    if(-not $outFile)
    {
        $outFile = $defaultName
    }
    $outFile = $outFile.replace(" ","_")

    # Invoke Seed Export
    Show-Command "npx @ordercloud/seeding download -f='$($outFile).yml' -i='$marketplaceId' -u=$user -p=*****"
    npx @ordercloud/seeding download -f="$($outFile).yml" -i="$marketplaceId" -u="$user" -p="$pass"
    
    Write-Host "Export Done: $(pwd)\$($outFile).yml" -ForegroundColor Green
    Write-Host
    Write-Host "TIPS:" -ForegroundColor Cyan
    Write-Host " Seed exports do not include:"
    Write-Host " * Writeable ApiClient IDs" -ForegroundColor Yellow
    Write-Host "    - New api client ids are created in the target environment"
    Write-Host "    - Using import tool, any reference to a source api client id in seed is mapped to new id by name (so multiple runs are okay)"
    Write-Host " * Secrets" -ForegroundColor Yellow
    Write-Host '    - Identified by searching for "Redacted" in seed file'
    Write-Host "    - Can replace redacted values with real values if desired. The import tool will use them and set them if provided."
    Write-Host " * User Passwords" -ForegroundColor Yellow
    Write-Host "    - OC exports null values for passwords"
    Write-Host "    - either force user to reset password, or specify a value in the seed file. This import script will use and set them if provided."
    Write-Host
    Write-Host "Consider making edits to seed file before running import:" -ForegroundColor Cyan
    Write-Host " - Provide any redacted data (see above)"
    Write-Host " - Remove any records you don't want to deploy."
    Write-Host " - Apply any environment specific changes."
    Write-Host
}

#endregion

# ////////////////////////////////////////////////////////////////////////////////////////////////
#
#region 6. IMPORT PROGRAM
#
# ////////////////////////////////////////////////////////////////////////////////////////////////
function Invoke-ImportProgram
{
    Write-Host "TOOL: Import Seed file"

    # Validate Pre-req
    if(Test-Path $PSYamlFolder)
    {
        Import-Module $PSYamlFolder
    } else {
        Write-Host
        Write-Host "PRE-REQ needed!"
        Write-Host "Install PSYaml module (to parse yaml files)"
        Write-Host " - See one-time setup: https://github.com/Phil-Factor/PSYaml"
        Write-Host " - Copy files to $PSYamlFolder"
        return;
    }

    # Gather inputs
    $seedFile = ""
    do {
        # Get *.yml file names in current directory sorted most recent on top
        $files = @(Get-ChildItem -Path "$(pwd)\*.yml" | Sort-Object -Property LastWriteTime -Descending | % {$_.Name}) + @('Try Again')
        # User to pick one
        $seedFile = Get-FromChoiceMenu "Choose seed file from current directory (if none shown, run export or copy seed.yml to $(pwd) and try again):" $files
    } while ($seedFile -eq "Try Again")

    $seedFile = "$(pwd)\$seedFile"

    # Validate seed file
    do {
        Write-Host "Validating seed file"
        Show-Command "npx @ordercloud/seeding validate -d='$seedFile'"
        npx @ordercloud/seeding validate -d="$seedFile"
    } until(0 -eq $host.ui.PromptForChoice(
        "All Green?",
        ' If not, correct any errors then re-validate',
        @(
            [System.Management.Automation.Host.ChoiceDescription]::new("&Yes, continue", "Continue with file")
            [System.Management.Automation.Host.ChoiceDescription]::new("&No, revalidate", "Revalidate after correcting changes")
        ),
        0))
    
    # Verify target
    Write-Host
    Write-Host "Ready to Import."
    Write-Host "Verify Target:" -ForegroundColor Cyan

    Invoke-ConfirmContext
    
    Write-Host
    Write-Host "Seed file: $seedFile" -ForegroundColor Yellow
    
    # Confirm
    Write-Host
    if((0 -ne $host.ui.PromptForChoice(
        "",
        'Run Import?',
        @(
            [System.Management.Automation.Host.ChoiceDescription]::new("&Yes", "Yes")
            [System.Management.Automation.Host.ChoiceDescription]::new("&No", "No")
        ),
        0)))
    {
        return;
    }

    # Read Seed File

    # Get-Help ConvertFrom-Yaml -Full
    $seed = ConvertFrom-Yaml -Path $seedFile
    # seed object SHAPE:
    # .Meta
    # .Objects
    #   .{type}
    #     [entries]
    #   ...
    # .Assignments
    #   .{type}
    #     [entries]
    #   ...

    # Here Goes!
    # Run Objects before Assignments. Pass sorted keys to mitigate dependencies.
    Invoke-Sync 'Objects'     $seed.Objects     @($defns | Where-Object {$_.Type -eq [OcType]::Entity -and $_.Name -ne 'Variants' -and $_.Name -ne 'VariantInventoryRecords'} | % {$_.Name}) # skip variants to run separate later
    Invoke-Sync 'Assignments' $seed.Assignments @($defns | Where-Object {$_.Type -eq [OcType]::Assignment} | % {$_.Name})

    # Run Variant pieces after all others
    if($seed.Objects.Products -and $seed.Objects.Variants)
    {
        Invoke-SyncEntries -type 'GenerateVariants' -entries $seed.Objects.Products
        Invoke-SyncVariants -entries $seed.Objects.Variants
        
        Invoke-SyncEntries -type 'VariantInventoryRecords' -entries $seed.Objects.VariantInventoryRecords
    }
    
    Write-Host "Import Done" -ForegroundColor Green
}

#endregion

# ////////////////////////////////////////////////////////////////////////////////////////////////
#
#region     6.b IMPORT SYNC HANDLERS
#
# ////////////////////////////////////////////////////////////////////////////////////////////////

# Manage group level - delegate SyncEntries for each
function Invoke-Sync($groupName, $arrEntryGroup, $sortedKeys)
{
    Write-Host
    Write-Host "SYNC $($arrEntryGroup.Keys.Count) $($groupName):" -ForegroundColor Green
    if($global:verbose)
    {
        Write-Host ($arrEntryGroup.Keys -join ", ") -ForegroundColor Gray
    }

    # Warn if seed file is not introducing unregisterd types
    Foreach ($type in $arrEntryGroup.Keys) {
    
        if(-not $sortedKeys.Contains($type) -and $type -ne "Variants" -and $type -ne "VariantInventoryRecords")
        {
            Write-Host
            Write-Host "WARNING! Unregistered Type $type in seed will be skipped. Add type '$type' to the sorted keys array variable for $groupName then re-run." -ForegroundColor Red
        }
    }

    # Loop sorted keys to ensure proper dependencies are in place
    Foreach ($type in $sortedKeys) {

        if($arrEntryGroup.Keys -eq $null -or -not $arrEntryGroup.Keys.Contains($type))
        {
            continue; # allow incomplete seed files, skip type
        }

        $entries = $arrEntryGroup[$type]
        Invoke-SyncEntries -type $type -entries $entries
    }
}

# Special handling for Variants
# Generated variants get random ids which may not match those in seed file, so need to lookup and match them by spec values
function Invoke-SyncVariants($entries)
{
    $byProduct = $entries | Group-Object -Property {$_.ProductID}
    
    $g = 0
    $gTotal = $byProduct.Length;
            
    Write-Host
    Write-Host "Processing $($entries.Length) variants over $($gTotal) products..." -ForegroundColor Cyan

    foreach($group in $byProduct)
    {
        $g++
        $list = Invoke-OcRequest $([OcRequest]@{
            verb = "GET"
            url = "/products/$($group.Name)/variants"
        })
        $variants = $list.Items
        
        Write-Host
        Write-Host "Processing Product $($group.Name) ($g of $gTotal)..." -ForegroundColor Cyan
    
        # Loop Each entry in group
        $count = 0
        foreach($entry in $group.Group) {
            $count++
            Write-Host "$(([string]$count).PadRight(3)) $type >" $entry.ID -NoNewline

            # Prepare request
            $request = Get-PushRequest -type 'Variants' -data $entry
                        
            if($request -ne $null -and $request.url -ne $null -and $request.url.Length -and $request.verb -ne $null)
            {
                # Fix ID if different than seed (since it is auto generated)
                $hasID = $variants | Where-Object {$_.ID -eq $entry.ID}
                if($hasID)
                {
                    Write-Host ' - Matched' -NoNewline
                } else
                {
                    # Lookup ID by matching spec options
                    $matchKey = @($entry.Specs | Sort-Object -Property {$_.SpecID} | % {"$($_.SpecID)$($_.OptionID)"}) -join ""
                    
                    $match = $variants | Where-Object {@($_.Specs | Sort-Object -Property {$_.SpecID} | %{"$($_.SpecID)$($_.OptionID)"}) -join "" -eq $matchKey}
                    if($match)
                    {
                        Write-Host " - Replacing $($match.ID)" -NoNewLine
                        $request.url = "/products/$($group.Name)/variants/$($match.ID)"
                    } else {
                        Write-Host " - not matched :(" -NoNewline
                    }
                }
                Write-Host
                $resp = Invoke-OcRequest($request);
            } else {
                # Notify if request was not prepared as expected (not yet implemented)
                Write-Host "Aborting type, Not Yet Implemented" -ForegroundColor Red
                break
            }            
        }

    }   
}

# Iterate Entries of group, calling PUT methods to sync
function Invoke-SyncEntries($type, $entries)
{
    Write-Host
    Write-Host "Processing $($entries.Length) $type..." -ForegroundColor Cyan
    
    # Loop Each entry in group
    $count = 0
    Foreach ($entry in $entries) {
        $count++
        $label = $entry.ID;
        if($type -eq 'XpIndices')
        {
            $label = "$($entry.ThingType)|$($entry.Key)"
        }
        if(-not $label)
        {
            $label = "[ $(@($entry.Keys| Where-Object {$_ -like "*ID" -and $entry[$_] -ne 'null' } | % { $entry[$_] }) -join ' | ') ]"
        }
        Write-Host "$(([string]$count).PadRight(3)) $type >" $label

        # Prepare request
        $request = Get-PushRequest -type $type -data $entry
            
        if($request -ne $null -and $request.url -ne $null -and $request.url.Length -and $request.verb -ne $null)
        {
            $resp = Invoke-OcRequest($request);
        } else {
            # Notify if request was not prepared as expected (not yet implemented)
            Write-Host "Aborting type, Not Yet Implemented" -ForegroundColor Red
            break
        }            
    }
}

#endregion

# ////////////////////////////////////////////////////////////////////////////////////////////////
#
#region     6.c IMPORT REQUEST MAPPING HELPERS
#
# ////////////////////////////////////////////////////////////////////////////////////////////////

# Utility - handle "null" string in seed file, returning $null needed for json conversion
function Get-ValueOrNull($value)
{
    if($value -eq "null")
    {
        return $null;
    }

    return $value;
}

# Utility - ensure proper boolean value is returned and not a string
function Get-BoolOrFalse($value)
{
    if($value -eq $true -or $value -eq "true")
    {
        return $true;
    }

    return $false;
}

# Utility - massage xp obj into JSON, support null case
function ConvertTo-JsonWithXp($body, $xp)
{
    # Write-Host "BODY: $body"
    # Write-Host "XP: $xp"
    $xpJson = $(If($xp -eq $null -or $xp -eq 'null') {'null'} Else {ConvertTo-Json $xp -Compress}) #compress to help with 8,000 char limit
    $body["xp"] = 'XPSPLICE'
    return $(ConvertTo-Json $body).replace('"XPSPLICE"', $xpJson)
}

#endregion

# ////////////////////////////////////////////////////////////////////////////////////////////////
#
#region     6.d IMPORT REQUEST MAPPING
#
# ////////////////////////////////////////////////////////////////////////////////////////////////

# Map/Build Payload - Returns URL, HTTP Verb, and maps $data to appropriate request Body per $type
# Returns $null if type is not yet handled
function Get-PushRequest($type, $data)
{
    # MANAGE MAPPINGS HERE PER TYPE (order doesn't matter here)
    switch($type)
    {
      # === OBJECTS =====
      #SecurityProfiles
      'SecurityProfiles'{
        $body = @{
            ID = $data.ID
            Name = $data.Name
            Roles = $data.Roles
        }
        return @{
            verb = "PUT"
            url = "/$type/$($data.ID)"
            body = ConvertTo-Json $body
        };
      }
      #ImpersonationConfigs
      'ImpersonationConfigs'{
        $body = @{
            ID = $data.ID
            ImpersonationBuyerID = Get-ValueOrNull $data.ImpersonationBuyerID
            ImpersonationGroupID = Get-ValueOrNull $data.ImpersonationGroupID
            ImpersonationUserID = Get-ValueOrNull $data.ImpersonationUserID
            BuyerID = Get-ValueOrNull $data.BuyerID
            GroupID = Get-ValueOrNull $data.GroupID
            UserID = Get-ValueOrNull $data.UserID
            SecurityProfileID = Get-ValueOrNull $data.SecurityProfileID
            ClientID = $global:ApiClientIdLookupByOldId[$data.ClientID]
        }
        return @{
            verb = "PUT"
            url = "/impersonationconfig/$($data.ID)"
            body = ConvertTo-Json $body
        };
      }
      #OpenIdConnects
      'OpenIdConnects'{
        $scopes = $null
        if($data.AdditionalIdpScopes -ne "null" -and $data.AdditionalIdpScopes -ne $null)
        {
            $scopes = @($data.AdditionalIdpScopes); # force array - for case of single item array collapsing into a string thanks to powershell auto unboxing
        }
        $body = @{
            ID = Get-ValueOrNull $data.ID
            OrderCloudApiClientID = $global:ApiClientIdLookupByOldId[$data.OrderCloudApiClientID]
            ConnectClientID = $data.ConnectClientID
            ConnectClientSecret = $data.ConnectClientSecret
            AppStartUrl = $data.AppStartUrl
            AuthorizationEndpoint = $data.AuthorizationEndpoint
            TokenEndpoint = $data.TokenEndpoint
            UrlEncoded = Get-BoolOrFalse $data.UrlEncoded
            IntegrationEventID = Get-ValueOrNull $data.IntegrationEventID
            # IntegrationEventName = Get-ValueOrNull $data.IntegrationEventName
            CallSyncUserIntegrationEvent = Get-BoolOrFalse $data.CallSyncUserIntegrationEvent
            AdditionalIdpScopes = $scopes
            CustomErrorUrl = Get-ValueOrNull $data.CustomErrorUrl
        }
        return @{
            verb = "PUT"
            url = "/$type/$($data.ID)"
            body = ConvertTo-Json $body
        };
      }
      #AdminUsers
      'AdminUsers'{
        $body = @{
            ID = $data.ID
            #CompanyID = $data.CompanyID <MarketplaceID placeholder>
            Username = $data.Username
            Password = Get-ValueOrNull $data.Password # Exports do not include passwords, but allow import if value is specified in file
            FirstName = $data.FirstName
            LastName = $data.LastName
            Email = $data.Email
            Phone = Get-ValueOrNull $data.Phone
            TermsAccepted = Get-ValueOrNull $data.TermsAccepted
            Active = Get-BoolOrFalse $data.Active
            #AvailableRoles = @($data.AvailableRoles)
            #DateCreated = $data. '2022-08-11T13 = $data.21 = $data.57.687+00 = $data.00'
            #PasswordLastSetDate = $data. '2022-08-11T13 = $data.21 = $data.57.733+00 = $data.00'
        }
        return @{
            verb = "PUT"
            url = "/$type/$($data.ID)"
            body = ConvertTo-JsonWithXp -body $body -xp $data.xp
        };
      }
      #AdminUserGroups
      'AdminUserGroups'{      
        $body = @{
            ID = $data.ID
            Name = $data.Name
            Description = Get-ValueOrNull $data.Description
        }
        return @{
            verb = "PUT"
            url = "/usergroups/$($data.ID)"
            body = ConvertTo-JsonWithXp -body $body -xp $data.xp
        };
      }
      #AdminAddresses
      'AdminAddresses'{
        $body = @{
            ID = $data.ID
            AddressName = Get-ValueOrNull $data.AddressName
            #DateCreated = $data.DateCreated
            CompanyName = Get-ValueOrNull $data.CompanyName
            FirstName = Get-ValueOrNull $data.FirstName
            LastName = Get-ValueOrNull $data.LastName
            Street1 = $data.Street1
            Street2 = Get-ValueOrNull $data.Street2
            City = $data.City
            State = $data.State
            Zip = $data.Zip
            Country = $data.Country
            Phone = Get-ValueOrNull $data.Phone
        }
        return @{
            verb = "PUT"
            url = "/addresses/$($data.ID)"
            body = ConvertTo-JsonWithXp -body $body -xp $data.xp
        };
      }
      #SupplierUsers
      'SupplierUsers'{
        $body = @{
            ID = $data.ID
            #CompanyID = $data.CompanyID <MarketplaceID placeholder>
            Username = $data.Username
            Password = Get-ValueOrNull $data.Password # Exports do not include passwords, but allow import if value is specified in file
            FirstName = $data.FirstName
            LastName = $data.LastName
            Email = $data.Email
            Phone = Get-ValueOrNull $data.Phone
            TermsAccepted = Get-ValueOrNull $data.TermsAccepted
            Active = Get-BoolOrFalse $data.Active
            #AvailableRoles = @($data.AvailableRoles)
            #DateCreated = $data. '2022-08-11T13 = $data.21 = $data.57.687+00 = $data.00'
            #PasswordLastSetDate = $data. '2022-08-11T13 = $data.21 = $data.57.733+00 = $data.00'
        }
        return @{
            verb = "PUT"
            url = "/suppliers/$($data.SupplierID)/users/$($data.ID)"
            body = ConvertTo-JsonWithXp -body $body -xp $data.xp
        };
      }
      #SupplierUserGroups
      'SupplierUserGroups'{      
        $body = @{
            ID = $data.ID
            Name = $data.Name
            Description = Get-ValueOrNull $data.Description
        }
        return @{
            verb = "PUT"
            url = "/suppliers/$($data.SupplierID)/usergroups/$($data.ID)"
            body = ConvertTo-JsonWithXp -body $body -xp $data.xp
        };
      }      
      #SupplierAddresses
      'SupplierAddresses'{
        $body = @{
            ID = $data.ID
            AddressName = Get-ValueOrNull $data.AddressName
            #DateCreated = $data.DateCreated
            CompanyName = Get-ValueOrNull $data.CompanyName
            FirstName = Get-ValueOrNull $data.FirstName
            LastName = Get-ValueOrNull $data.LastName
            Street1 = $data.Street1
            Street2 = Get-ValueOrNull $data.Street2
            City = $data.City
            State = $data.State
            Zip = $data.Zip
            Country = $data.Country
            Phone = Get-ValueOrNull $data.Phone
        }
        return @{
            verb = "PUT"
            url = "/suppliers/$($data.SupplierID)/addresses/$($data.ID)"
            body = ConvertTo-JsonWithXp -body $body -xp $data.xp
        };
      }
      #MessageSenders
      'MessageSenders'{      
        $types = $null
        if($data.MessageTypes -ne "null" -and $data.MessageTypes -ne $null)
        {
            $types = @($data.MessageTypes); # force array - for case of single item array collapsing into a string thanks to powershell auto unboxing
        }
        $roles = $null
        if($data.ElevatedRoles -ne "null" -and $data.ElevatedRoles -ne $null)
        {
            $roles = @($data.ElevatedRoles); # force array - for case of single item array collapsing into a string thanks to powershell auto unboxing
        }
        $body = @{
            ID = $data.ID
            Name = $data.Name
            MessageTypes = $types
            Description = Get-ValueOrNull $data.Description
            URL = $data.URL
            ElevatedRoles = $roles
            SharedKey = Get-ValueOrNull $data.SharedKey
        }
        return @{
            verb = "PUT"
            url = "/$type/$($data.ID)"
            body = ConvertTo-JsonWithXp -body $body -xp $data.xp
        };            
        }
      #ApiClients
      'ApiClients'{
        # Can't write IDs, so create new id unless we can match existing by name
        $clientId = Get-OrCreateApiClientId $data.AppName
        $global:ApiClientIdLookupByOldId[$data.ID] = $clientId # store ID lookup for replacements later. ASSUMPTION: all API Clients are processed before OldID lookup is called.
        $body = @{      
            ClientSecret = Get-ValueOrNull $data.ClientSecret # REDACTED, but seed file can be edited to contain the secret
            AccessTokenDuration = $data.AccessTokenDuration
            Active = Get-BoolOrFalse $data.Active
            AppName = $data.AppName
            RefreshTokenDuration = $data.RefreshTokenDuration
            DefaultContextUserName = Get-ValueOrNull $data.DefaultContextUserName
            AllowAnyBuyer = Get-BoolOrFalse $data.AllowAnyBuyer
            AllowAnySupplier = Get-BoolOrFalse $data.AllowAnySupplier
            AllowSeller = Get-BoolOrFalse $data.AllowSeller
            IsAnonBuyer = Get-BoolOrFalse $data.IsAnonBuyer
            OrderCheckoutIntegrationEventID = Get-ValueOrNull $data.OrderCheckoutIntegrationEventID
            OrderReturnIntegrationEventID = Get-ValueOrNull $data.OrderReturnIntegrationEventID
            AddToCartIntegrationEventID = Get-ValueOrNull $data.AddToCartIntegrationEventID
            MinimumRequiredRoles = $data.MinimumRequiredRoles
            MinimumRequiredCustomRoles = $data.MinimumRequiredCustomRoles
            MaximumGrantedRoles = $data.MaximumGrantedRoles
            MaximumGrantedCustomRoles = $data.MaximumGrantedCustomRoles
            #OrderCheckoutIntegrationEventName = Get-ValueOrNull $data.OrderCheckoutIntegrationEventName
            #AssignedBuyerCount = $data.AssignedBuyerCount
            #AssignedSupplierCount = $data.AssignedSupplierCount
            #OrderReturnIntegrationEventName = Get-ValueOrNull $data.OrderReturnIntegrationEventName
            #AddToCartIntegrationEventName = Get-ValueOrNull $data.AddToCartIntegrationEventName
        }
        return @{
            verb = "PUT"
            url = "/apiclients/$clientId"
            body = ConvertTo-JsonWithXp -body $body -xp $data.xp
        };      
      }
      #Locales
      'Locales'{
        $body = @{
            ID = $data.ID
            #OwnerID = <MarketplaceID placeholder> # By not specifying it will default to target marketplace
            Currency = $data.Currency
            Language = Get-ValueOrNull $data.Language
        }
        return @{
            verb = "PUT"
            url = "/$type/$($data.ID)"
            body = $body | ConvertTo-Json
        };
      }
      #Webhooks
      'Webhooks'{
        $roles = $null
        if($data.ElevatedRoles -ne "null" -and $data.ElevatedRoles -ne $null)
        {
            $roles = @($data.ElevatedRoles); # force array - for case of single item array collapsing into a string thanks to powershell auto unboxing
        }
        $body = @{
            ID = $data.ID
            Name = $data.Name
            Description = Get-ValueOrNull $data.Description
            Url = $data.Url
            HashKey = $data.HashKey # REDACTED, but seed file can be edited to contain the secret
            ElevatedRoles = $roles
            #ConfigData = $null # Handled below similar to xp, just renaming back to ConfigData
            BeforeProcessRequest = Get-BoolOrFalse $data.BeforeProcessRequest
            ApiClientIDs = @(Get-NewClientIds $data.ApiClientIDs)
            WebhookRoutes = "WEBHOOKPLACEHOLDER"
        }
        $webhookRoutes = $(if($data.WebhookRoutes -eq $null -or $data.WebhookRoutes -eq $null){"null"} else {ConvertTo-Json @($data.WebhookRoutes) -Compress})
        return @{
            verb = "PUT"
            url = "/$type/$($data.ID)"
            body = $(ConvertTo-JsonWithXp -body $body -xp $data.ConfigData).replace('"xp"','"ConfigData"').replace('"WEBHOOKPLACEHOLDER"',$webhookRoutes)
        };
      }
      #IntegrationEvents
      'IntegrationEvents'{      
        $body = @{
            ID = $data.ID
            Name = $data.Name
            ElevatedRoles = $data.ElevatedRoles
            #ConfigData = $null # Handled below similar to xp, just renaming back to ConfigData
            EventType = $data.EventType
            CustomImplementationUrl = Get-ValueOrNull $data.CustomImplementationUrl
            HashKey = $data.HashKey # REDACTED, but seed file can be edited to contain the secret
        }
        return @{
            verb = "PUT"
            url = "/$type/$($data.ID)"
            body = $(ConvertTo-JsonWithXp -body $body -xp $data.ConfigData).replace('"xp"','"ConfigData"')
        };
      }
      #Users
      'Users'{           
        $body = @{
            ID = $data.ID
            # CompanyID = $data.CompanyID
            Username = Get-ValueOrNull $data.Username
            # Password = Get-ValueOrNull $data.Password
            FirstName = Get-ValueOrNull $data.FirstName
            LastName = Get-ValueOrNull $data.LastName
            Email = Get-ValueOrNull $data.Email
            Phone = Get-ValueOrNull $data.Phone
            TermsAccepted = Get-ValueOrNull $data.TermsAccepted
            Active = Get-BoolOrFalse $data.Active
            AvailableRoles = $data.AvailableRoles
            # DateCreated = $data.DateCreated
            # PasswordLastSetDate = $data.PasswordLastSetDate
            # BuyerID = $data.BuyerID
        }
        return @{
            verb = "PUT"
            url = "/buyers/$($data.BuyerID)/$type/$($data.ID)"
            body = ConvertTo-JsonWithXp -body $body -xp $data.xp
        };
      }
      #UserGroups
      'UserGroups'{  
        $body = @{
            ID = $data.ID
            Name = Get-ValueOrNull $data.Name
            Description = Get-ValueOrNull $data.Description
            #BuyerID = $data.BuyerID
        }
        return @{
            verb = "PUT"
            url = "/buyers/$($data.BuyerID)/$type/$($data.ID)"
            body = ConvertTo-JsonWithXp -body $body -xp $data.xp
        };
      }      
      #Addresses
      'Addresses'{
        $body = @{
            ID = $data.ID
            AddressName = Get-ValueOrNull $data.AddressName
            #DateCreated = $data.DateCreated
            CompanyName = Get-ValueOrNull $data.CompanyName
            FirstName = Get-ValueOrNull $data.FirstName
            LastName = Get-ValueOrNull $data.LastName
            Street1 = $data.Street1
            Street2 = Get-ValueOrNull $data.Street2
            City = $data.City
            State = $data.State
            Zip = $data.Zip
            Country = $data.Country
            Phone = Get-ValueOrNull $data.Phone
        }
        return @{
            verb = "PUT"
            url = "/buyers/$($data.BuyerID)/addresses/$($data.ID)"
            body = ConvertTo-JsonWithXp -body $body -xp $data.xp
        };
      }
      #XpIndices
      'XpIndices'{  
        $body = @{
            ThingType = $data.ThingType
            Key = $data.Key
        }
        return @{
            verb = "PUT"
            url = "/$type"
            body = ConvertTo-Json $body
        };
      }
      #Buyers
      'Buyers'{    
        $body = @{
            ID = $data.ID
            Name = Get-ValueOrNull $data.Name
            DefaultCatalogID = Get-ValueOrNull $data.DefaultCatalogID
            Active = Get-BoolOrFalse $data.Active
            #DateCreated = $data.DateCreated
        }
        return @{
            verb = "PUT"
            url = "/$type/$($data.ID)"
            body = ConvertTo-JsonWithXp -body $body -xp $data.xp
        };
      }
      #Catalogs
      'Catalogs'{
        $body = @{
            ID = $data.ID
            #OwnerID: <MarketplaceID placeholder>
            Name = Get-ValueOrNull $data.Name
            Description = Get-ValueOrNull $data.Description
            Active = Get-BoolOrFalse $data.Active
            #CategoryCount: 36
        }
        return @{
            verb = "PUT"
            url = "/$type/$($data.ID)"
            body = ConvertTo-JsonWithXp -body $body -xp $data.xp
        };
      }
      #Categories
      'Categories'{      
        $body = @{
            ID = $data.ID
            Name = Get-ValueOrNull $data.Name
            Description = Get-ValueOrNull $data.Description
            ListOrder = $data.ListOrder
            Active = Get-BoolOrFalse $data.Active
            ParentID = Get-ValueOrNull $data.ParentID
        }
        return @{
            verb = "PUT"
            url = "/catalogs/$($data.CatalogID)/$type/$($data.ID)"
            body = ConvertTo-JsonWithXp -body $body -xp $data.xp
        };
      }
      #PriceSchedules
      'PriceSchedules'{   
        $priceBreaks = @();
        if($data.PriceBreaks -and $data.PriceBreaks.Length -gt 0)
        {
            Foreach($price in $data.PriceBreaks)
            {
                $priceBreaks += @{
                    Quantity = Get-ValueOrNull $price.Quantity
                    Price = Get-ValueOrNull $price.Price
                    SalePrice = Get-ValueOrNull $price.SalePrice
                    SubscriptionPrice = Get-ValueOrNull $price.SubscriptionPrice
                }
            }
        } 
        $body = @{
            #OwnerID = <MarketplaceID placeholder> # When not specified, Defaults to target marketplace
            ID = $data.ID
            Name = Get-ValueOrNull $data.Name
            ApplyTax = Get-BoolOrFalse $data.ApplyTax
            ApplyShipping = Get-BoolOrFalse $data.ApplyShipping
            MinQuantity = Get-ValueOrNull $data.MinQuantity
            MaxQuantity = Get-ValueOrNull $data.MaxQuantity
            UseCumulativeQuantity = Get-BoolOrFalse $data.UseCumulativeQuantity
            RestrictedQuantity = Get-BoolOrFalse $data.RestrictedQuantity
            PriceBreaks = $priceBreaks
            Currency = Get-ValueOrNull $data.Currency
            SaleStart = Get-ValueOrNull $data.SaleStart
            SaleEnd = Get-ValueOrNull $data.SaleEnd
            IsOnSale = Get-BoolOrFalse $data.IsOnSale
        }
        return @{
            verb = "PUT"
            url = "/$type/$($data.ID)"
            body = ConvertTo-JsonWithXp -body $body -xp $data.xp
        };
      }
      #Promotions
      'Promotions'{
        $body = @{
            #OwnerID = <MarketplaceID placeholder> # When not specified, Defaults to target marketplace
            ID = $data.ID
            LineItemLevel = Get-BoolOrFalse $data.LineItemLevel
            Code = Get-ValueOrNull $data.Code
            Name = Get-ValueOrNull $data.Name
            RedemptionLimit = Get-ValueOrNull $data.RedemptionLimit
            RedemptionLimitPerUser = Get-ValueOrNull $data.RedemptionLimitPerUser
            # RedemptionCount = $data.RedemptionCount
            Description = Get-ValueOrNull $data.Description
            FinePrint = Get-ValueOrNull $data.FinePrint
            StartDate = Get-ValueOrNull $data.StartDate
            ExpirationDate = Get-ValueOrNull $data.ExpirationDate
            EligibleExpression = Get-ValueOrNull $data.EligibleExpression
            ValueExpression = Get-ValueOrNull $data.ValueExpression
            CanCombine = Get-BoolOrFalse $data.CanCombine
            AllowAllBuyers = Get-BoolOrFalse $data.AllowAllBuyers
            Active = Get-BoolOrFalse $data.Active
            AutoApply = Get-BoolOrFalse $data.AutoApply
            Priority = Get-ValueOrNull $data.Priority
        }
        return @{
            verb = "PUT"
            url = "/$type/$($data.ID)"
            body = ConvertTo-JsonWithXp -body $body -xp $data.xp
        };
      }
      #Products
      'Products'{
        $inventory = $null;
        if($data.Inventory -ne $null -and $data.Inventory -ne 'null')
        {
            $inventory = @{
                Enabled = Get-BoolOrFalse $data.Inventory.Enabled
                NotificationPoint = Get-ValueOrNull $data.Inventory.NotificationPoint
                VariantLevelTracking = Get-BoolOrFalse $data.Inventory.VariantLevelTracking
                OrderCanExceed = Get-BoolOrFalse $data.Inventory.OrderCanExceed
                QuantityAvailable = Get-ValueOrNull $data.Inventory.QuantityAvailable
            }
        }

        $body = @{
            #OwnerID = <MarketplaceID placeholder> # When not specified, Defaults to target marketplace
            ID = $data.ID
            Name = Get-ValueOrNull $data.Name
            Description = Get-ValueOrNull $data.Description
            DefaultPriceScheduleID = Get-ValueOrNull $data.DefaultPriceScheduleID
            Active = Get-BoolOrFalse $data.Active
            AutoForward = Get-BoolOrFalse $data.AutoForward
            ParentID = Get-ValueOrNull $data.ParentID
            IsParent = Get-BoolOrFalse $data.IsParent
            QuantityMultiplier = $data.QuantityMultiplier
            ShipWeight = Get-ValueOrNull $data.ShipWeight
            ShipHeight = Get-ValueOrNull $data.ShipHeight
            ShipWidth = Get-ValueOrNull $data.ShipWidth
            ShipLength = Get-ValueOrNull $data.ShipLength
            #SpecCount: 0
            #VariantCount: 0
            ShipFromAddressID = Get-ValueOrNull $data.ShipFromAddressID
            Inventory = $inventory
            #LastUpdated: '2021-11-18T17:35:41.737+00:00'
            DefaultSupplierID = Get-ValueOrNull $data.DefaultSupplierID
            AllSuppliersCanSell = Get-BoolOrFalse $data.AllSuppliersCanSell
            Returnable = Get-BoolOrFalse $data.Returnable
        }
        return @{
            verb = "PUT"
            url = "/$type/$($data.ID)"
            body = ConvertTo-JsonWithXp -body $body -xp $data.xp
        };
      }
      #GenerateVariants - custom key
      'GenerateVariants' {      
        return @{
            verb = "POST"
            url = "/products/$($data.ID)/variants/generate?overwriteExisting=true"
        };
      }
      #Specs
      'Specs'{
        $body = @{
            #OwnerID = <MarketplaceID placeholder> # When not specified, Defaults to target marketplace
            ID = $data.ID
            Name = Get-ValueOrNull $data.Name
            ListOrder = Get-ValueOrNull $data.ListOrder
            DefaultValue = Get-ValueOrNull $data.DefaultValue
            Required = Get-BoolOrFalse $data.Required
            AllowOpenText = Get-BoolOrFalse $data.AllowOpenText
            DefaultOptionID = Get-ValueOrNull $data.DefaultOptionID
            DefinesVariant = Get-BoolOrFalse $data.DefinesVariant
        }
        return @{
            verb = "PUT"
            url = "/$type/$($data.ID)"
            body = ConvertTo-JsonWithXp -body $body -xp $data.xp
        };
      }
      #Variants
      'Variants'{
        $inventory = $null;
        if($data.Inventory -ne $null -and $data.Inventory -ne 'null')
        {
            $inventory = @{
                QuantityAvailable = Get-ValueOrNull $data.Inventory.QuantityAvailable
            }
        }
        $body = @{
            #OwnerID = <MarketplaceID placeholder> # When not specified, Defaults to target marketplace
            ID = $data.ID
            Name = Get-ValueOrNull $data.Name
            Description = Get-ValueOrNull $data.Description
            Active = Get-BoolOrFalse $data.Active
            ShipWeight = Get-ValueOrNull $data.ShipWeight
            ShipHeight = Get-ValueOrNull $data.ShipHeight
            ShipWidth = Get-ValueOrNull $data.ShipWidth
            ShipLength = Get-ValueOrNull $data.ShipLength
            Inventory = $inventory
        }
        return @{
            verb = "PUT"
            url = "/products/$($data.ProductID)/variants/$($data.ID)"
            body = ConvertTo-JsonWithXp -body $body -xp $data.xp
        };
      }
      #SpecOptions
      'SpecOptions'{ 
        $body = @{
            #OwnerID = <MarketplaceID placeholder> # When not specified, Defaults to target marketplace
            ID = $data.ID
            Value = $data.Value
            ListOrder = Get-ValueOrNull $data.ListOrder
            IsOpenText = Get-BoolOrFalse $data.IsOpenText
            PriceMarkupType = Get-ValueOrNull $data.PriceMarkupType
            PriceMarkup = Get-ValueOrNull $data.PriceMarkup
        }
        return @{
            verb = "PUT"
            url = "/specs/$($data.SpecID)/options/$($data.ID)"
            body = ConvertTo-JsonWithXp -body $body -xp $data.xp
        };
      }
      #Variants
      'Variants'{
        $inventory = $null;
        if($data.Inventory -ne $null -and $data.Inventory -ne 'null')
        {
            $inventory = @{
                QuantityAvailable = Get-ValueOrNull $data.Inventory.QuantityAvailable
            }
        }
        $body = @{
            #OwnerID = <MarketplaceID placeholder> # When not specified, Defaults to target marketplace
            ID = $data.ID
            Name = $data.Name
            Description = Get-ValueOrNull $data.Description
            Active = Get-BoolOrFalse $data.Active
            ShipWeight = Get-ValueOrNull $data.ShipWeight
            ShipHeight = Get-ValueOrNull $data.ShipHeight
            ShipWidth = Get-ValueOrNull $data.ShipWidth
            ShipLength = Get-ValueOrNull $data.ShipLength
            Inventory = $inventory
        }
        return @{
            verb = "PUT"
            url = "/products/$($data.ProductID)/variants/$($data.ID)"
            body = ConvertTo-JsonWithXp -body $body -xp $data.xp
        };
      }      
      #InventoryRecords
      'InventoryRecords'{
        $body = @{
            #OwnerID = <MarketplaceID placeholder> # When not specified, Defaults to target marketplace
            ID = $data.ID
            AllowAllBuyers = Get-BoolOrFalse $data.AllowAllBuyers
            AddressID = Get-ValueOrNull $data.AddressID
            OrderCanExceed = Get-BoolOrFalse $data.OrderCanExceed
            QuantityAvailable = Get-ValueOrNull $data.QuantityAvailable
        }
        return @{
            verb = "PUT"
            url = "/products/$($data.ProductID)/$($type)/$($data.ID)"
            body = ConvertTo-JsonWithXp -body $body -xp $data.xp
        };
      }      
      #VariantInventoryRecords
      'VariantInventoryRecords'{
        $body = @{
            #OwnerID = <MarketplaceID placeholder> # When not specified, Defaults to target marketplace
            ID = $data.ID
            AllowAllBuyers = Get-BoolOrFalse $data.AllowAllBuyers
            AddressID = Get-ValueOrNull $data.AddressID
            OrderCanExceed = Get-BoolOrFalse $data.OrderCanExceed
            QuantityAvailable = Get-ValueOrNull $data.QuantityAvailable
        }
        return @{
            verb = "PUT"
            url = "/products/$($data.ProductID)/variants/$($data.VariantID)/inventoryrecords/$($data.ID)"
            body = ConvertTo-JsonWithXp -body $body -xp $data.xp
        };
      }
      #Incrementors
      'Incrementors'{      
        $body = @{
            #OwnerID = <MarketplaceID placeholder> # When not specified, Defaults to target marketplace
            ID = $data.ID
            Name = Get-ValueOrNull $data.Name
            LastNumber = Get-ValueOrNull $data.LastNumber
            LeftPaddingCount = Get-ValueOrNull $data.LeftPaddingCount
        }
        return @{
            verb = "PUT"
            url = "/$type/$($data.ID)"
            body = ConvertTo-Json $body
        };
      }
      #Suppliers
      'Suppliers'{      
        $body = @{
            #OwnerID = <MarketplaceID placeholder> # When not specified, Defaults to target marketplace
            ID = $data.ID
            Name = Get-ValueOrNull $data.Name
            DateCreated = Get-ValueOrNull $data.DateCreated
            Active = Get-BoolOrFalse $data.Active
            AllBuyersCanOrder = Get-BoolOrFalse $data.AllBuyersCanOrder
        }
        return @{
            verb = "PUT"
            url = "/$type/$($data.ID)"
            body = ConvertTo-JsonWithXp -body $body -xp $data.xp
        };
      }
      #ProductFacets
      'ProductFacets'{      
        $body = @{
            #OwnerID = <MarketplaceID placeholder> # When not specified, Defaults to target marketplace
            ID = $data.ID
            Name = Get-ValueOrNull $data.Name
            XpPath = Get-ValueOrNull $data.XpPath
            ListOrder = Get-ValueOrNull $data.ListOrder
            MinCount = Get-ValueOrNull $data.MinCount
        }
        return @{
            verb = "PUT"
            url = "/$type/$($data.ID)"
            body = ConvertTo-JsonWithXp -body $body -xp $data.xp
        };
      }
      #SellerApprovalRules
      'SellerApprovalRules'{      
        $body = @{
            #OwnerID = <MarketplaceID placeholder> # When not specified, Defaults to target marketplace
            ID = $data.ID
            Name = Get-ValueOrNull $data.Name
            ApprovalType = Get-ValueOrNull $data.ApprovalType
            Description = Get-ValueOrNull $data.Description
            ApprovingGroupID = Get-ValueOrNull $data.ApprovingGroupID
            RuleExpression = Get-ValueOrNull $data.RuleExpression
        }
        return @{
            verb = "PUT"
            url = "/approvalrules/$($data.ID)"
            body = ConvertTo-JsonWithXp -body $body -xp $data.xp
        };
      }
      #SpendingAccounts
      'SpendingAccounts'{      
        $body = @{
            #OwnerID = <MarketplaceID placeholder> # When not specified, Defaults to target marketplace
            ID = $data.ID
            Name = Get-ValueOrNull $data.Name
            Balance = Get-ValueOrNull $data.Balance
            AllowAsPaymentMethod = Get-BoolOrFalse $data.AllowAsPaymentMethod
            RedemptionCode = Get-ValueOrNull $data.RedemptionCode
            StartDate = Get-ValueOrNull $data.StartDate
            EndDate = Get-ValueOrNull $data.EndDate
        }
        return @{
            verb = "PUT"
            url = "/buyers/$($data.BuyerID)/$($type)/$($data.ID)"
            body = ConvertTo-JsonWithXp -body $body -xp $data.xp
        };
      }
      #CostCenters
      'CostCenters'{      
        $body = @{
            #OwnerID = <MarketplaceID placeholder> # When not specified, Defaults to target marketplace
            ID = $data.ID
            Name = Get-ValueOrNull $data.Name
            Description = Get-ValueOrNull $data.Description
        }
        return @{
            verb = "PUT"
            url = "/buyers/$($data.BuyerID)/$($type)/$($data.ID)"
            body = ConvertTo-JsonWithXp -body $body -xp $data.xp
        };
      }
      #ApprovalRules
      'ApprovalRules'{      
        $body = @{
            #OwnerID = <MarketplaceID placeholder> # When not specified, Defaults to target marketplace
            ID = $data.ID
            Name = Get-ValueOrNull $data.Name
            Description = Get-ValueOrNull $data.Description
            ApprovingGroupID = Get-ValueOrNull $data.ApprovingGroupID
            RuleExpression = Get-ValueOrNull $data.RuleExpression
        }
        return @{
            verb = "PUT"
            url = "/buyers/$($data.BuyerID)/$($type)/$($data.ID)"
            body = ConvertTo-JsonWithXp -body $body -xp $data.xp
        };
      }
      #CreditCards
      'CreditCards'{      
        $body = @{
            #OwnerID = <MarketplaceID placeholder> # When not specified, Defaults to target marketplace
            ID = $data.ID
            Token = Get-ValueOrNull $data.Token
            DateCreated = Get-ValueOrNull $data.DateCreated
            CardType = Get-ValueOrNull $data.CardType
            PartialAccountNumber = Get-ValueOrNull $data.PartialAccountNumber
            CardholderName = Get-ValueOrNull $data.CardholderName
            ExpirationDate = Get-ValueOrNull $data.ExpirationDate
        }
        return @{
            verb = "PUT"
            url = "/buyers/$($data.BuyerID)/$($type)/$($data.ID)"
            body = ConvertTo-JsonWithXp -body $body -xp $data.xp
        };
      }

      # === Assignments =====
        #SecurityProfileAssignments
        'SecurityProfileAssignments'{
            $body = @{
                SecurityProfileID = $data.SecurityProfileID
                BuyerID = Get-ValueOrNull $data.BuyerID
                SupplierID = Get-ValueOrNull $data.SupplierID
                UserID = Get-ValueOrNull $data.UserID
                UserGroupID = Get-ValueOrNull $data.UserGroupID
            }
            return @{
                verb = "POST"
                url = "/Securityprofiles/assignments"
                body = ConvertTo-Json $body
            };
        }
        #CatalogAssignments
        'CatalogAssignments'{
            $body = @{
                CatalogID = $data.CatalogID
                BuyerID = $data.BuyerID
                ViewAllCategories = Get-BoolOrFalse $data.ViewAllCategories
                ViewAllProducts = Get-BoolOrFalse $data.ViewAllProducts
            }
            return @{
                verb = "POST"
                url = "/catalogs/assignments"
                body = ConvertTo-Json $body
            };
        }
        #PromotionAssignments
        'PromotionAssignments'{
            $body = @{
                PromotionID = $data.PromotionID
                BuyerID = $data.BuyerID
                UserGroupID = Get-ValueOrNull $data.UserGroupID
                #UserID = Get-ValueOrNull $data.UserID
            }
            return @{
                verb = "POST"
                url = "/promotions/assignments"
                body = ConvertTo-Json $body
            };
        }
        #CategoryProductAssignments
        'CategoryProductAssignments'{
            $body = @{
                CategoryID = $data.CategoryID
                ProductID = $data.ProductID
                ListOrder = Get-ValueOrNull $data.ListOrder
            }
            return @{
                verb = "POST"
                url = "/catalogs/$($data.CatalogID)/categories/productassignments"
                body = ConvertTo-Json $body
            };
        }
        #ProductCatalogAssignment
        'ProductCatalogAssignment'{
            $body = @{
                CatalogID = $data.CatalogID
                ProductID = $data.ProductID
            }
            return @{
                verb = "POST"
                url = "/catalogs/productassignments"
                body = ConvertTo-Json $body
            };
        }  
        #SpecProductAssignments
        'SpecProductAssignments'{
            $body = @{
                SpecID = $data.SpecID
                ProductID = $data.ProductID
                DefaultValue = Get-ValueOrNull $data.DefaultValue
                DefaultOptionID = Get-ValueOrNull $data.DefaultOptionID
            }
            return @{
                verb = "POST"
                url = "/specs/productassignments"
                body = ConvertTo-Json $body
            };
        }


        #UserGroupAssignments
        'UserGroupAssignments'{        
            $body = @{
                UserGroupID = $data.UserGroupID
                UserID = $data.UserID
            }
            return @{
                verb = "POST"
                url = "/buyers/$($data.BuyerID)/usergroups/assignments"
                body = ConvertTo-Json $body
            };
        }
        #SpendingAccountAssignments
        'SpendingAccountAssignments'{
            $body = @{
                SpendingAccountID = $data.SpendingAccountID
                UserID = Get-ValueOrNull $data.UserID
                UserGroupID = Get-ValueOrNull $data.UserGroupID
                AllowExceed = Get-BoolOrFalse $data.AllowExceed
            }
            return @{
                verb = "POST"
                url = "/buyers/$($data.BuyerID)/spendingaccounts/assignments"
                body = ConvertTo-Json $body
            };
        }
        #AddressAssignments
        'AddressAssignments'{
            $body = @{
                AddressID = $data.AddressID
                UserID = Get-ValueOrNull $data.UserID
                UserGroupID = Get-ValueOrNull $data.UserGroupID
                IsShipping = Get-BoolOrFalse $data.IsShipping
                IsBilling = Get-BoolOrFalse $data.IsBilling
            }
            return @{
                verb = "POST"
                url = "/buyers/$($data.BuyerID)/addresses/assignments"
                body = ConvertTo-Json $body
            };
        }
        #CostCenterAssignments
        'CostCenterAssignments'{
            $body = @{
                CostCenterID = $data.CostCenterID
                UserID = Get-ValueOrNull $data.UserID
                UserGroupID = Get-ValueOrNull $data.UserGroupID
            }
            return @{
                verb = "POST"
                url = "/buyers/$($data.BuyerID)/costcenters/assignments"
                body = ConvertTo-Json $body
            };
        }
        #CreditCardAssignments
        'CreditCardAssignments'{
            $body = @{
                CreditCardID = $data.CreditCardID
                UserID = Get-ValueOrNull $data.UserID
                UserGroupID = Get-ValueOrNull $data.UserGroupID
                IsShipping = Get-BoolOrFalse $data.IsShipping
            }
            return @{
                verb = "POST"
                url = "/buyers/$($data.BuyerID)/creditcards/assignments"
                body = ConvertTo-Json $body
            };
        }
        #CategoryAssignments
        'CategoryAssignments'{
            $body = @{
                CategoryID = $data.CategoryID
                UserID = Get-ValueOrNull $data.UserID
                UserGroupID = Get-ValueOrNull $data.UserGroupID
                BuyerID = Get-ValueOrNull $data.BuyerID
                Visible = Get-BoolOrFalse $data.Visible
                ViewAllProducts = Get-BoolOrFalse $data.ViewAllProducts
            }
            return @{
                verb = "POST"
                url = "/catalogs/$($data.CatalogID)/categories/assignments"
                body = ConvertTo-Json $body
            };
        }
        #ProductSupplierAssignments
        'ProductSupplierAssignments'{
            $body = @{
                DefaultPriceScheduleID = Get-ValueOrNull $data.DefaultPriceScheduleID
            }
            return @{
                verb = "PUT"
                url = "/products/$($data.ProductID)/suppliers/$($data.SupplierID)"
                body = ConvertTo-Json $body
            };
        }
        #AdminUserGroupAssignments
        'AdminUserGroupAssignments'{
            $body = @{
                UserID = Get-ValueOrNull $data.UserID
                UserGroupID = Get-ValueOrNull $data.UserGroupID
            }
            return @{
                verb = "POST"
                url = "/usergroups/assignments"
                body = ConvertTo-Json $body
            };
        }
        #ApiClientAssignments
        'ApiClientAssignments'{
            $body = @{
                ApiClientID = Get-NewClientId $data.ApiClientID
                BuyerID = Get-ValueOrNull $data.BuyerID
                SupplierID = Get-ValueOrNull $data.SupplierID
            }
            return @{
                verb = "POST"
                url = "/apiclients/assignments"
                body = ConvertTo-Json $body
            };
        }
        #LocaleAssignments
        'LocaleAssignments'{
            $body = @{
                LocaleID = $data.LocaleID
                BuyerID = Get-ValueOrNull $data.BuyerID
                UserGroupID = Get-ValueOrNull $data.UserGroupID
            }
            return @{
                verb = "POST"
                url = "/locales/assignments"
                body = ConvertTo-Json $body
            };
        }
        #ProductAssignments
        'ProductAssignments'{
            $body = @{
                SellerID = Get-ValueOrNull $data.SellerID
                ProductID = Get-ValueOrNull $data.ProductID
                BuyerID = Get-ValueOrNull $data.BuyerID
                UserGroupID = Get-ValueOrNull $data.UserGroupID
                PriceScheduleID = Get-ValueOrNull $data.PriceScheduleID
            }
            return @{
                verb = "POST"
                url = "/products/assignments"
                body = ConvertTo-Json $body
            };
        }  
        #SupplierUserGroupsAssignments
        'SupplierUserGroupsAssignments'{
            $body = @{
                UserID = Get-ValueOrNull $data.UserID
                UserGroupID = Get-ValueOrNull $data.UserGroupID
            }
            return @{
                verb = "POST"
                url = "/suppliers/$($data.SupplierID)/usergroups/assignments"
                body = ConvertTo-Json $body
            };
        }  
        #SupplierBuyerAssignments
        'SupplierBuyerAssignments'{
            return @{
                verb = "PUT"
                url = "/suppliers/$($data.SupplierID)/buyers/$($data.BuyerID)"
            };
        }  

      default{return $null}
    }
}

#endregion

# ////////////////////////////////////////////////////////////////////////////////////////////////
#
#region     6.e API CLIENT IMPORT HANDLER
#
# ////////////////////////////////////////////////////////////////////////////////////////////////

# API Client ID lookup
# IDs are unique per environment, so for import need a bridge to match appropriately. Cache the lookup in this hash table.
$global:ApiClientIdLookupByName = @{}
$global:ApiClientIdLookupByOldId = @{}

# ------------------------

# Gets matching Api Client Id by name, or creates one if not exists in this environment.
function Get-OrCreateApiClientId($name)
{
    # CASE 1 - already cached, return lookup
    if($global:ApiClientIdLookupByName.ContainsKey($name))
    {
        Write-Host "ApiClientId - Matched $name to $($global:ApiClientIdLookupByName[$name])" -ForegroundColor Magenta
        return $global:ApiClientIdLookupByName[$name];
    }

    # CASE 2 - query and match by name (add to cache for next time)
    $uri = "/apiclients?search=$name&searchOn=AppName"
    $response = Invoke-OcRequest $([OcRequest]@{verb = "Get"; url = $uri })

    if($global:verbose)
    {
        $log = $response | ConvertTo-Json
        Write-Host "CASE 2 Query Result:" -ForegroundColor Gray
        Write-Host $log -ForegroundColor Gray
    }
    if($response.Meta.TotalCount -gt 0)
    {
        $firstId = $response.Items[0].ID
        Write-Host "ApiClientId - Queried $name to be $firstId" -ForegroundColor Magenta
        $global:ApiClientIdLookupByName[$name] = $firstId
        return $firstId
    }
        
    # CASE 3 - create new API Client as match (stub only, and add to cache)        
    $body = @{
        AppName = $name
        AccessTokenDuration = 600
    }
    $json = $body | ConvertTo-Json
    $newClient = Invoke-OcRequest $([OcRequest]@{verb = "POST"; url = "/apiclients"; body = $json })

    if($global:verbose)
    {
        $log = $newClient | ConvertTo-Json
        Write-Host "CASE 3:" -ForegroundColor Gray
        Write-Host $log -ForegroundColor Gray
    }
    $newId = $newClient.ID
    Write-Host "ApiClientId - Created $name as $newId" -ForegroundColor Magenta
    $global:ApiClientIdLookupByName[$name] = $newId
    return $newId
}

# Utility - return array or new ids, given array of old ApiClient ids
# ** Assumes Get-OrCreateApiClientId() has already run for all ApiClients
function Get-NewClientIds($arrOldIds)
{
    if(-not $arrOldIds)
    {
        return $null
    }
    else
    {
        return @(@($arrOldIds) | % {$global:ApiClientIdLookupByOldId[$_]})
    }
}

# Utility - return New Id by Old Id
# ** Assumes Get-OrCreateApiClientId() has already run for all ApiClients
function Get-NewClientId($oldId)
{
    if(-not $oldId)
    {
        return $null
    }
    else
    {
        return $global:ApiClientIdLookupByOldId[$oldId]
    }
}

#endregion

# ////////////////////////////////////////////////////////////////////////////////////////////////
#
#region 7. DELETE PROGRAM
#
# ////////////////////////////////////////////////////////////////////////////////////////////////
function Invoke-DeleteProgram
{
    Write-Host "TOOL: Delete Data by Root Type"

    Invoke-ConfirmContext

    # Only choose from root types (when OC deletes a root type it removes all subtypes)
    # ie. deleting product deletes all related variants, variant inventory, etc
    $choices = $defns | where { $_.Type -eq [OcType]::Entity -and -not $_.LivesUnder } | % { $_.Name } | Sort-Object
    $typesToPurge = Get-FromChoiceMenu "Choose Types To Purge?" $choices -allowMultiple

    # Sort to honor dependencies, loop backwards, deleting non-dependent first
    $reverseSortedTypes = @()
    $testString = "abcdef"
    for ($i=$defns.Length-1; $i -ge 0; $i--) {
        if($typesToPurge.Contains($defns[$i].Name))
        {
            $reverseSortedTypes += $defns[$i];
        }
    }

    Show-Context

    if(0 -ne $host.ui.PromptForChoice(
        "Confirm Delete?",
        " Continue to Purge $($reverseSortedTypes.Length) types (and all their sub data)?",
        @(
            [System.Management.Automation.Host.ChoiceDescription]::new("&Yes", "Yes, delete")
            [System.Management.Automation.Host.ChoiceDescription]::new("&No", "No, cancel")
        ),
        0))
        { return }

    # Loop types and attempt deletions
    foreach($defn in $reverseSortedTypes)
    {
        $ids = Get-EntityIds $defn
        Write-Host
        Write-Host "Purging $($defn.Name) - $($ids.Length) records" -ForegroundColor Cyan
        $count = 0
        foreach($id in $ids)
        {         
            $count++   
            Write-Host ([string]$count).PadRight(3) " Deleting $id"
            $request = Get-DeleteRequest $defn $id
            if($request)
            {
                $resp = Invoke-OcRequest($request);
            } else {
                Write-Host "Type $($defn.Name) not mapped. Skipping" -ForegroundColor Red
                break;
            }
        }
        Clear-Cache $($defn.Name)
        
    }

    Write-Host
    Write-Host "Purge Done" -ForegroundColor Green
}

#endregion

# ////////////////////////////////////////////////////////////////////////////////////////////////
#
#region 7.b DELETE REQUEST MAPPING
#
# ////////////////////////////////////////////////////////////////////////////////////////////////

# Map $defn to its Delete Request. Most follow the same pattern, but any outliers are handled here as well.
# Provide $defn and $id of record to delete.
# $id supports a format for cases of nested entities, ie: "PRODUCTID:VARIANTID"
function Get-DeleteRequest([OcDefn]$defn, $id)
{

    $slug = $defn.Name.ToLowerInvariant()
    if($defn.ApiName)
    {
        $slug = $defn.ApiName
    }

    $url = "/$($slug)/$($id)"

    # Support requests with multiple IDs
    if($defn.LivesUnder -or $id.Contains(':'))
    {
        $ids = $id.Split(':')
        if($ids.Length -ne $defn.LivesUnder.Length)
        {
            throw "Type $($defn.Name) defn expected $($defn.LivesUnder.Length) multiple IDs but got mismatched input of $($ids.Length)"
        }

        $depUrl = ""
        for($x = 0; $x -lt $ids.Length; $x++)
        {
            $depUrl = "$depUrl/$($defn.LivesUnder[$x])/$($ids[$x])"
        }
        $url = "$($depUrl)$($url)"
    }

    $request = @{
        verb = "DELETE"
        url = $url
    };

    return $request;
}

#endregion

# ////////////////////////////////////////////////////////////////////////////////////////////////
#
#region 8. REPORT PROGRAM
#
# ////////////////////////////////////////////////////////////////////////////////////////////////

function Invoke-ReportProgram
{    
    Write-Host "TOOL: Reporting"
    
    Invoke-ConfirmContext
    
    Write-Host
    Write-Host "Choose report type, name, and columns." -ForegroundColor Cyan

    # 1. Choose Type
    $entityOptions = $defns | where { $_.Type -eq [OcType]::Entity } | % { $_.Name } | Sort-Object
    $assignmentOptions = $defns | where { $_.Type -eq [OcType]::Assignment } | % { $_.Name } | Sort-Object

    $viewOption = 1
    do
    {
        if($viewOption -eq 1)
        {
            $choices = @("See Assignment Options...") + $entityOptions
        } else {        
            $choices = @("See Entity Options...") + $assignmentOptions
        }
        $choice = Get-FromChoiceMenu "1. Report Type:" $choices

        if($choice -eq "See Assignment Options...")
        {
            $viewOption = 2
        } elseif($choice -eq "See Entity Options...") {
            $viewOption = 1
        } else {
            $viewOption = 0
        }
    } while($viewOption -gt 0)

    $defn = $defns | where { $_.Name -eq $choice }

    # 2. Choose Name
    $defaultName = 'report'
    $outFile = Read-Host -Prompt "2. Report Name (enter for default, $defaultName)"
    if(-not $outFile)
    {
        $outFile = $defaultName
    }
    $outFile = $outFile.replace(" ","_")
    if(-not $outFile.EndsWith('.csv'))
    {
        $outFile = "$($outFile).csv"
    }
    $outFile = "$(pwd)\$($outFile)"

    # Get Data
    Write-Host "Querying data..." -ForegroundColor DarkGray
    $list = Get-ListRequest $defn
    
    # $list[0] | Get-Member | Format-Table
    
    if(-not $list)
    {
        Write-Host "No Data"
    } else {
    # 3. Choose Columns
        $listSample = $list[0];
        $point = 97 # a.
        $depCols = New-Object System.Collections.ArrayList # parellel array to LivesUnder, where each matching position is a list of cols to report
        $depHeader = ""

        # Pick ancestor properties
        if($listSample.AncestorIDs)
        {
            for($a = 0; $a -lt @($listSample.AncestorIDs).Length; $a++)
            {
                # $ancestorID = @($listSample.AncestorIDs)[$a]
                $ancestorDefn =  $defns | Where-Object { $_.Name -eq $defn.LivesUnder[$a] }
                $ancesterSet = Get-ListRequest $ancestorDefn # This uses cache
                $ancestorSample = $ancesterSet[0]

                $props = Get-PropertyNames $ancestorSample
                $defaults = Get-DefaultSelectedColumns -d $ancestorDefn -names $props
                $cols = Get-FromChoiceMenu "3.$([char]$point) Choose Ancestor $($ancestorDefn.Name) Properties for Report:" $props -allowMultiple $defaults
                $point++

                [void]$depCols.Add(@($cols));
                if($cols)
                {
                    if($depHeader)
                    {
                        $depHeader = "$($depHeader),"
                    }
                    $depHeader = "$($depHeader)$(@($cols | % {"$($ancestorDefn.Name).$_"}) -join ",")"
                }
            }
        }

        # Pick entity properties
        $props = Get-PropertyNames $listSample
        $defaults = Get-DefaultSelectedColumns -d $defn -names $props
        $bullet = '3.'
        if($point -gt 97)
        {
            $bullet = "$bullet$([char]$point)"
        }
        $cols = Get-FromChoiceMenu "$bullet Choose $($defn.Name) Properties for Report:" $props -allowMultiple $defaults

    # 4. Write file
        # Start file        
        Write-Host "Writing report to: $outFile" -ForegroundColor Cyan

        $header = @(@($depHeader) + $cols) -join ","
        Set-Content -Path "$outFile" -Value $header

        foreach($entry in $list)
        {
            $row = ""
            $isFirst = $true

            for($s=0; $s -lt $defn.LivesUnder.Length; $s++)
            {
                $sub = $defn.LivesUnder[$s]
                if($depCols[$s])
                {
                    $ancestorID = @($entry.AncestorIDs)[$s]                  
                    $ancestorDefn =  $defns | Where-Object { $_.Name -eq $sub }
                    $ancesterEntry = Get-ListRequest $ancestorDefn | where {$_.ID -eq $ancestorID} # This uses cache
                    $row = Get-CellValues -row $row -cols $depCols[$s] -entry $ancesterEntry
                }
            }
            
            $row = Get-CellValues -row $row -cols $cols -entry $entry
            
            Add-Content -Path "$outFile" -Value $row
        }
    }
            
    Write-Host "Report Done" -ForegroundColor Green
}

# Helper to extract smart defaults given list of property options
function Get-DefaultSelectedColumns([OcDefn]$d, [string[]]$names)
{
    if($d.Type -eq [OcType]::Entity)
    {
        return @($names | where {$_ -eq 'Name' -or $_ -eq 'ID' -or $_ -eq 'Active'})
    } else {
        return @($names | where {$_ -like '*ID'})
    }
}

# Appends partial row to $row, given subset of $cols for $entry
function Get-CellValues($row,$cols,$entry)
{
    foreach($col in $cols)
    {
        if($row)
        {
            $row = "$($row),"
        }
                
        $cell = ""
        if($col.Contains('.'))
        {
            # Handle nested properties
            $xpath = $col.Split('.')
            $cell = $entry
            foreach($segment in $xpath)
            {
                $cell = $cell."$segment"
            }
        }
        else
        {   # Handle direct property
            $cell = $($entry."$col")
        }
                
        $row = "$($row)$($cell)"
    }

    return $row
}

# Given an OC response object for an entity, extract the properties we might want to list in a report
# Crawl and nested objects and provide dot notation for each explicit option.
function Get-PropertyNames($obj)
{    
    $props = $obj | get-member | where {$_.MemberType -eq 'NoteProperty'}
    
    $result = @()
    foreach($prop in $props)
    {
        if($prop.Definition -like '*object*' -and $prop.Name -ne 'AncestorIDs' -and $obj."$($prop.Name)" -ne $null)
        {
            # recursively expand nested object properties
            $result = $result + $(@(Get-PropertyNames $($obj."$($prop.Name)")) | % {"$($prop.Name).$($_)"})
        } else {
            $result += $prop.Name
        }
    }
    return $result
}

#endregion

# ////////////////////////////////////////////////////////////////////////////////////////////////
#
#region 9 MAIN PROCESS
#
# ////////////////////////////////////////////////////////////////////////////////////////////////

# Stop on first error
$ErrorActionPreference = "Stop"

Clear-host

# Capture script inputs
if($baseUrl)
{
    $global:baseUrl = $baseUrl
}
if($token)
{
    $global:token = $token
}
$global:verbose = $verbose

Show-Logo

$toolsMenu = @(
    'Report...',
    'Export...',
    'Import...',
    'Delete...',
    'Exit' )

$exit = $false
do {

    # Clear all caches to avoid being stale, minding external activities
    Clear-AllCache

    Show-ToolBox

    # Tool Selector
    $tool = Get-FromChoiceMenu "Pick a tool:" $toolsMenu
    switch($tool)
    {
        'Report...' {
            Invoke-ReportProgram
        }
        'Export...' {
            Invoke-ExportProgram
        }
        'Import...' {
            Invoke-ImportProgram
        }
        'Delete...' {
            Invoke-DeleteProgram
        }
        'Exit' {
            Write-Host 'Goodbye' -ForegroundColor Cyan
            $exit = $true
        }
    }
} until ($exit)

Write-Host

#endregion