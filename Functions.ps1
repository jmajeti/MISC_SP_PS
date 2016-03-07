############################################################################
# Create SharePoint Health Check Dashboard Using PowerShell Script 
# Created By Bilel Marouen
# Date: 18/10/2015
############################################################################



Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
#Url of the new Test site that will be created using this script and were we will perform all tests
$url= <TestSiteUrl>

Function UpdateListEQ($res,$id,$msg)
{
    #update list
    #<DashboardSiteUrl> that will contain the dashboard Should be created manually before running the script
    $SpSite=Get-spweb <DashboardSiteUrl>
    $list=$SpSite.lists[Patching]

     if($res -eq $null)
    {
          $item=$list.GetItemById($id)
          $item[Status]=  "<img src='<RedSign.jpeg>'></img>"
          $item.Update()f
          write-host Error $msg -ForegroundColor Red 
    }
    else
    {
          $item=$list.GetItemById($id)
          $item[Status]= "<img src='<GreenSign.jpeg>'></img>"
          $item.Update()
          write-host success $msg -ForegroundColor Green 
    }
}
Function UpdateListNE($res,$id,$msg)
{
    #update list
    $SpSite=Get-spweb <DashboardSiteUrl>
    $list=$SpSite.lists[Patching]
     if($res -ne $null)
    {
          $item=$list.GetItemById($id)
          $item[Status]= "<img src='<RedSign.jpeg>'></img>"
          $item.Update()
          write-host Error $msg -ForegroundColor Red 
    }
    else
    {
          $item=$list.GetItemById($id)
          $item[Status]= "<img src='<GreenSign.jpeg>'></img>"
          $item.Update()
          write-host success $msg -ForegroundColor Green 
    }
}
function createSC()
{
    $template = Get-SPWebTemplate STS#1
    #<WebAppUrl> that will contains the Site test Collection 
    $webapp= Get-WebApplication <WebAppURl>
    $url=<TestSiteUrl>
    $res=  New-SPSite -url $url -OwnerAlias <UserAccount> -HostHeaderWebApplication $webapp -Name Test Site -Template $template  -ErrorAction SilentlyContinue
    
    #Update List
    UpdateListEQ $res 1 createSC

}

function DeleteSC()
{
    $site=Get-SPSite <TestSiteUrl>
    $resdelete= Remove-SPSite -Identity $site -GradualDelete -Confirm$false
    
    #Update List
   # UpdateListNE $res 18 DeleteSC
}

function SCAdmin()
{
    $user= <UserAccount>
    $site=Get-spsite <TestSiteUrl>
    $web=$site.RootWeb
    $owner=$web.EnsureUser($user)
    $res=($site.Owner= $owner)

    
    #Update List
    UpdateListEQ $res 4 SCAdmin
}

function CreateWeb()
{
    $template = Get-SPWebTemplate STS#1
    #<TestSubSiteUrl> the Subsite will be created under the test site colletion created previously
    $url=<TestSubSiteUrl>
    $res=  New-SPWeb -url $url  -Name Subsite -Template $template -ErrorAction SilentlyContinue
   
    #Update List
    UpdateListEQ $res 13 CreateWeb
}

function CreateSPG()
{
    $url=v
    $web=Get-SPWeb $url
    $user=$web.EnsureUser("<UserAccount>")
    $res= $web.SiteGroups.Add(“TestGRP”,$user,$user, “Test group”) 
    
    #Update List
    UpdateListNE $res 14 CreateSPG

    #####Add users in the group
    $ownerGroup = $web.SiteGroups[TestGRP] 
    $ownerGroup.AllowMembersEditMembership=$true
    #$ownerGroup.Update() 
    $user=$web.Site.RootWeb.EnsureUser(<AnotherUserAccount>)
    $res=$ownerGroup.Users.Add($user,,,)    
    #$ownerGroup.Update()

    #Update List
    UpdateListNE $res 15 AddUserToGrp
}

function CreateSiteColumn()
{
    #Get the site collection and web object
    $siteColl = Get-SPSite -Identity <TestSiteUrl>
    $rootWeb = $siteColl.RootWeb
    #Assign fieldXMLString variable with field XML for site column
    $fieldXMLString = 'Field Type=Text
    Name=Site Column Test
    Description=Site Column Test
    DisplayName=Site Column Test
    StaticName=Site Column Test
    Group=Portman Applications
    Hidden=FALSE
    Required=FALSE
    Sealed=FALSE
    ShowInDisplayForm=TRUE
    ShowInEditForm=TRUE
    ShowInListSettings=TRUE
    ShowInNewForm=TRUEField'
    #See field XML on console
    #write-host $fieldXMLString
    #Create site column from XML string
    $res=$rootWeb.Fields.AddFieldAsXml($fieldXMLString) 

    #Update List
    UpdateListEQ $res 16 CreateSiteColumn
}

function CreateContentType()
{
    $siteColl = Get-SPSite -Identity <TestSiteUrl>
    $rootWeb = $siteColl.RootWeb
    $parent = $rootWeb.AvailableContentTypes[Document]
    $contentType =  New-Object Microsoft.SharePoint.SPContentType -ArgumentList @($parent,$rootWeb.ContentTypes,Test Content Type)
    $contentType.Group = Test Content Type
    $contentType.Description = Test Content Type
    $res=$rootWeb.ContentTypes.Add($contentType)
     
    #Update List
    UpdateListEQ $res 17 CreateContentType
}

function CreateList()
{
    $siteColl = Get-SPWeb -Identity <TestSiteUrl>  
    $ListUrl=TestList
    $ListTitle=Test List
    #Create new List
    $listTemplate = $siteColl.ListTemplates[Custom List]
    $res=$siteColl.Lists.Add($ListUrl,,$listTemplate)
    #Update List
    UpdateListEQ $res 18 CreateList
}

function CreateDocumentLibrary()
{
    $siteColl = Get-SPWeb -Identity <TestSiteUrl>
    #Create new Document library
    $listTemplate = [Microsoft.SharePoint.SPListTemplateType]DocumentLibrary 
    $res=$siteColl.Lists.Add(Test Document Library,Test Document Library,$listTemplate) 

    #Update List
    UpdateListEQ $res 19 CreateDocumentLibrary
}
function AddDocument()
{
    $siteColl = Get-SPWeb -Identity <TestSiteUrl>
    
    #add Document
    $spFolder = $siteColl.GetFolder(Test Document Library) 
    $spFileCollection = $spFolder.Files 
    $file = Get-ChildItem 'DSourceScriptsbilelpatchtest Document.rtf'
    $res=$spFileCollection.Add(Test Document Librarytest Document.rtf,$file.OpenRead(),$false) 

    #Update List
    UpdateListEQ $res 20 AddDocument
}

function AddContentTypeToList()
{
    $siteColl = Get-SPWeb -Identity <TestSiteUrl>
    
    #add Content Type To List
    $list=$siteColl.Lists[TestList]
    $list.ContentTypesEnabled=$true
    $list.Update()

    $contenttype=$siteColl.ContentTypes[Test Content Type]
    $res=$list.ContentTypes.Add($contenttype)
    
    #Update List
    UpdateListEQ $res 22 AddContenAddContentTypeToList
}

function AddFieldsToList()
{
    $siteColl = Get-SPWeb -Identity <TestSiteUrl>
    #add Field To List
    $list=$siteColl.Lists[TestList]
    $Single = [Microsoft.SharePoint.SPFieldType]Text
    $DefaultViewFieldOptions = [Microsoft.SharePoint.SPAddFieldOptions]AddFieldToDefaultView
    Function AddField($list,$xml,$bool,$fieldOptions){
    $list.Fields.AddFieldAsXml($xml,$bool,$fieldOptions)
    }
    $SingleLine  =Field Type='Text' DisplayName='test Field' Required = 'FALSE'
    $res=AddField $list $SingleLine  $FALSE $DefaultViewFieldOptions
    #$list.Update()

    #Update List
    UpdateListEQ $res 23 AddFieldsToList
}

function AddItemToList()
{
    $siteColl = Get-SPWeb -Identity <TestSiteUrl>
    #add item To List
    $list=$siteColl.Lists[TestList]
    $item=$list.Items.Add()
    $item[Title]=test2
    $res=$item.Update()
    #Update List
    UpdateListNE $res 24 AddItemToList
}

function DeleteItemFromList()
{
    $siteColl = Get-SPWeb -Identity <TestSiteUrl>
    #Delete item from List
    $list=$siteColl.Lists[TestList]
    $res=$list.GetItemById(16).delete()
    #Update List
    UpdateListNE $res 25 DeleteItemFromList
}

function DeleteItemFromDocumentLibrary()
{
    $siteColl = Get-SPWeb -Identity <TestSiteUrl>
    #Delete Document
    $spFolder = $siteColl.GetFolder(Test Document Library) 
    $spFileCollection = $spFolder.Files 
    $res=$spFileCollection.delete(Test Document Librarytest Document.rtf) 
    #Update List
    UpdateListNE $res 26 DeleteItemFromDocumentLibrary
}

function ActivateFeatures()
{
        enable-SPFeature -Identity PublishingSite -url <TestSiteUrl>
        $res=get-SPFeature -Identity PublishingSite -Site <TestSiteUrl>
    
     #Update List
    UpdateListEQ $res 30 ActivateFeatures


    ##We have to enable the site feature in ordre to be able to add page to the site
    ##We need this action for the next function Addpage
    $web=Get-SPWeb <TestSiteUrl>
    enable-SPFeature -Identity Publishingweb -Url $web.Url
} 

function Addpage()
{
    $siteColl = Get-SPWeb -Identity <TestSiteUrl>
    $pubWeb =[Microsoft.SharePoint.Publishing.PublishingWeb]GetPublishingWeb($siteColl)
    # Create blank web part page
    $pl = $pubWeb.GetAvailablePageLayouts()  Where { $_.Name -eq PageFromDocLayout.aspx } #you may change BlankWebPartPage.aspx to your custom page layout file name
    $newPage = $pubWeb.AddPublishingPage(TestPage.aspx, $pl) #filename need end with .aspx extension
    $newPage.Title=Test Page
    $newPage.Update()
    # Check-in and publish page
    $newPage.CheckIn()
    $res=$newPage.ListItem.File.Publish()
    #Update List
    UpdateListNE $res 27 Addpage
}

function RetrieveUPS()
{
    #Load SharePoint User Profile assemblies
    [System.Reflection.Assembly]LoadWithPartialName(“Microsoft.SharePoint”)
    [System.Reflection.Assembly]LoadWithPartialName(“Microsoft.Office.Server.UserProfiles”)
    #Get Context
    $serviceContext = Get-SPServiceContext -Site “<WebAppUrl_were the test site is created”

    #Instantiate User Profile Manager
    $userProfileConfigManager = New-Object Microsoft.Office.Server.UserProfiles.UserProfileManager($serviceContext);

    #Get All User Profiles
    $profiles = $userProfileConfigManager.GetEnumerator() Where {$_.MultiloginAccounts -like <UserAccount>}

    #Loop through all user profiles and display account name
    foreach($profile in $profiles)
    {
        $res=$profile.get_Item(“AccountName”) 
    }

    #Update List
    UpdateListEQ $res 28 RetrieveUPS
}

function Availability()
{
    # The WEB APPS list to test 
    #please Create the same file in the same location as the script and list the Web Applications Urls to test their availability
    $URLListFile = <Urls.txt> 
    $URLList = Get-Content $URLListFile -ErrorAction SilentlyContinue 
  $result = OK

  Foreach($Uri in $URLList)
  {
    try {
        $req = [system.Net.WebRequest]Create($Uri)
        $req.UseDefaultCredentials = $true
        $res = $req.GetResponse()
        $res1=$res.StatusCode
        } 
    catch [System.Net.WebException]
     {
        $res = $_.Exception.Response    
        $res1=$res.StatusCode     
     }

     if($res1 -notlike OK)
    {
        $result=NotFound
    }
  }
    #Update List
    if($result -like ok)
    {
        UpdateListEQ $res 29 Availability
    }
    else
    {
        UpdateListNE $res 29 Availability
    }
}


createSC
SCAdmin
CreateWeb
CreateSPG
CreateSiteColumn
CreateContentType
CreateList
CreateDocumentLibrary
AddDocument
AddContentTypeToList
AddFieldsToList
AddItemToList
DeleteItemFromList
DeleteItemFromDocumentLibrary
ActivateFeatures
Addpage
RetrieveUPS
Availability


#DeleteSC