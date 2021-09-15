<#
    Title: coe_catme_students.ps1
    Authors: Dean Bunn and Steve Pigg
    Last Edit: 2021-09-14
#>

#Var for Config Settings
$cnfgSettings = $null; 

#Check for Settings File 
if((Test-Path -Path ./config.json) -eq $true)
{
  
    #Import Json Configuration File
    $cnfgSettings =  Get-Content -Raw -Path .\config.json | ConvertFrom-Json;

}
else
{
    #Create Blank Config Object and Export to Json File
    $blnkConfig = new-object PSObject -Property (@{ AD_Domain="DC=my,DC=college,DC=edu"; 
                                                    IAM_Key=""; 
                                                    IAM_URL="https://iamsite.edu";
                                                    Org_Code="";
                                                  });

    $blnkConfig | ConvertTo-Json | Out-File .\config.json;

    #Exit Script
    exit;
}

#Hashtable for Course Instructors
$htInstructors = @{};

#Array for Instructors
$arrInstructors = @();

#Array for Reporting
$arrReport = @();

#Load CATME Students File
$csvCATMEStudents = Import-Csv -Path .\catme_students.csv;

#catme_students.csv
#email,person_id,instructors,first_name,last_name

#Loop Through Students Listings and Fing the Unique Email Addresses for Instructors
foreach($ccs in $csvCATMEStudents)
{
    
    #Check for Empty Entries on Instructor Emails
    if([string]::IsNullOrEmpty($ccs.instructors) -eq $false)
    {

        #Check for Numerous Instructors
        if($ccs.instructors.ToString().Contains(",") -eq $true)
        {

            #Split the Numerous Entries and Check Against Unique Hash Table
            foreach($insrEmlAddr in $ccs.instructors.ToString().Split(","))
            {

                if($htInstructors.ContainsKey($insrEmlAddr.ToString().Trim()) -eq $false)
                {
                    $htInstructors.Add($insrEmlAddr.ToString().Trim(),1);
                }

            }#End of Split Array

        }
        else
        {

            #Check Individual Instructor Entry Against Unique Hash Table
            if($htInstructors.ContainsKey($ccs.instructors.ToString().Trim()) -eq $false)
            {
                $htInstructors.Add($ccs.instructors.ToString().Trim(),1);
            }

        }#End of Null\Empty Check on Instructors Column


    }#End of $ccs.instructors Null\Empty Check
   
}#End of First Loop Through $csvCATMEStudents Looking for Unique Instructors


#Lookup Each Instructor in Active Directory

#Var for ADsPath
$strADSPath = "LDAP://" + $cnfgSettings.AD_Domain;
#Setup Directory Searcher
$ADsPath = [ADSI]$strADSPath;
$ADSearcher = New-Object DirectoryServices.DirectorySearcher($ADsPath);
$ADSearcher.PageSize = 900;
$ADSearcher.SearchScope = "SubTree";
[void]$ADSearcher.PropertiesToLoad.Add("displayName");
[void]$ADSearcher.PropertiesToLoad.Add("sAMAccountName");
[void]$ADSearcher.PropertiesToLoad.Add("userPrincipalName");
[void]$ADSearcher.PropertiesToLoad.Add("extensionAttribute7");

#Lookup Each Unique Instructor
foreach($key in $htInstructors.Keys)
{

    #Base Search Filter on UPN 
    $ADSearcher.filter = "(&(objectClass=user)(userPrincipalName=" + $key + "))"

    #Search for AD User
    $ADSrchResult = $ADSearcher.FindOne();

    #Null\Empty Check on User Search Result
    if($ADSrchResult -ne $null)
    {

        #Custom Object for Instructor
        $cstInstr = new-object PSObject -Property (@{ user_id=""; iam_id=""; display_name=""; email=""; prime_dept=""; coe="No";});

        #Check for IAM ID
        if($ADSrchResult.Properties["extensionAttribute7"].Count -gt 0)
        {

            #Set IAM ID
            $cstInstr.iam_id = $ADSrchResult.Properties["extensionAttribute7"][0].ToString();

            #Set User ID
            $cstInstr.user_id = $ADSrchResult.Properties["sAMAccountName"][0].ToString().ToLower();

            #Set Email
            $cstInstr.email = $ADSrchResult.Properties["userPrincipalName"][0].ToString().ToLower();

            #Check Display Name
            if($ADSrchResult.Properties["displayName"].Count -gt 0)
            {
                #Set Display Name
                $cstInstr.display_name = $ADSrchResult.Properties["displayName"][0].ToString();
            }

            #Add Instructor to Array Listing
            $arrInstructors += $cstInstr;

        }#End of IAM ID Check on AD3 Account
        

    }#End of $ADSrchResult Null Check
    
}#End of Instructor HashTable Foreach

#Pull Instructor IAM Payroll Information from IAM
foreach($cstInstcr in $arrInstructors)
{
    #Var for IAM URL 
    [string]$iamURL = $cnfgSettings.IAM_URL + $cstInstcr.iam_id + "?key=" + $cnfgSettings.IAM_Key + "&v=1.0";
    
    #Pull Instructor Payroll Information from IAM via API Call
    $irmInstcr = Invoke-RestMethod -ContentType "application/json" -Uri $iamURL;

    #Check for Response Data and Results
    if($irmInstcr.responseData -ne $null -and $irmInstcr.responseData.results -ne $null -and $irmInstcr.responseData.results.Count -gt 0)
    {
        
        #Set Primary Dept 
        $cstInstcr.prime_dept = $irmInstcr.responseData.results[0].apptDeptDisplayName;


        foreach($ppsAsgn in $irmInstcr.responseData.results)
        {

            #Check Org 
            if($ppsAsgn.apptBouOrgOId -eq $cnfgSettings.Org_Code)
            {
                $cstInstcr.coe = "Yes";
            }

        }#End of ReponseData.Results Foreach

    }

}#End of $arrInstructors IAM Lookup Foreach 

#Go Back Through CSV File and Parse Report File
foreach($rwCATME in $csvCATMEStudents)
{

    #Custom Object for Report Row
    $cstRptRow = new-object PSObject -Property (@{ email=""; person_id=""; first_name=""; last_name=""; instructors=""; instructor_coe="No"; instructors_primary_dept=""; });
    
    #Pull Over Reporting 
    $cstRptRow.email = $rwCATME.email;
    $cstRptRow.person_id = $rwCATME.person_id;
    $cstRptRow.first_name = $rwCATME.first_name;
    $cstRptRow.last_name = $rwCATME.last_name;

    #Var for Row Instructors Array
    $arrRwInstrtors = @();

    #Parse Instructors Again
    if([string]::IsNullOrEmpty($rwCATME.instructors) -eq $false)
    {
        
        #Load Raw Instructors Information
        $cstRptRow.instructors = $rwCATME.instructors;

        #Check for Numerous Instructors
        if($rwCATME.instructors.ToString().Contains(",") -eq $true)
        {

            #Split Up Multiple Instructor Information By Comma
            foreach($insrEmlAddr in $rwCATME.instructors.ToString().Split(","))
            {
                $arrRwInstrtors += $insrEmlAddr.ToString().ToLower().Trim();
            }
           
        }
        else
        {
            $arrRwInstrtors += $rwCATME.instructors.ToString().ToLower().Trim().Replace(",","");
        }

    }#End of Parsing Instructors Second Time
    

    #Determine If COE and Instructor Department Values
    foreach($rwInstrcr in $arrRwInstrtors)
    {

        foreach($crsInstr in $arrInstructors)
        {
            
            #Match Instructor By Email Value
            if($crsInstr.email.ToLower() -eq $rwInstrcr)
            {

                #Check For COE Status
                if($crsInstr.coe -eq "Yes")
                {
                    $cstRptRow.instructor_coe = "Yes";
                }

                if($cstRptRow.instructors_primary_dept.ToString().Contains($crsInstr.prime_dept) -eq $false)
                {
                    $cstRptRow.instructors_primary_dept += $crsInstr.prime_dept + ",";
                }

            }#End of Email Value Match Check

        }#End of $arrInstructors Foreach

    }#End of $arrRwInstrtors

    #Strip Off Trailing Comma on Instructor Primary Dept
    if([string]::IsNullOrEmpty($cstRptRow.instructors_primary_dept) -eq $false)
    {
        $cstRptRow.instructors_primary_dept = $cstRptRow.instructors_primary_dept.ToString().TrimEnd(',');
    }

    #Add To Report Array
    $arrReport += $cstRptRow;
}


#Export CSV Report File
$rptFileName = "UCD_CATME_Students_Report_" + (Get-Date -Format d).ToString().Replace("/","-") + ".csv";
$arrReport | Select-Object "email","person_id","first_name","last_name","instructors","instructor_coe","instructors_primary_dept" | Export-CSV $rptFileName -NoTypeInformation;