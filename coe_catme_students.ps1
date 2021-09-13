<#
    Title: coe_catme_students.ps1
    Authors: Dean Bunn
    Last Edit: 2021-09-13
#>


#Var for Config Settings
$cnfgSettings = $null; 

#Check for Settings File 
if((Test-Path -Path ./config.json) -eq $true)
{
    

}
else
{
    #Create Blank Config Object and Export to Json File
    $blnkConfig = new-object PSObject -Property (@{ AD_Domain=""; 
                                                    IAM_Key=""; 
                                                    IAM_URL="";
                                                    COE_Dept_Codes=@();
                                                  });

    $blnkConfig | ConvertTo-Json | Out-File .\config.json;

    #Exit Script
    exit;
}


#catme_students.csv
#email,person_id,instructors,first_name,last_name