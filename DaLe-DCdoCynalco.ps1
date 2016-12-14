# Description
# -----------
# Active Directory Management tool
# Created for Task1 Operating Systems Windows 2 @ Vives
# Adapted for the common task Linux/Windows/Networking

# Author : David Lejeune
# Created : 27/09/2016
# School : Vives
# Course : Operating Systems Windows 2
# Class : 3PB-ICT
# Group : 2
#


#------------------------------------------------------------------------------
#Imports

Import-Module ActiveDirectory

#------------------------------------------------------------------------------
#Script Variables
$DC1 = "Vives"
$DC2 = "be"

$Menu = ""
$Menu1 = "Create new OU";
$Menu2 = "New User"
$Menu3 = "Bulk Usermanagement based on CSV"
$Menu4 = "Check User existence"
$Menu5 = "Bulk delete User from CSV"
$Menu6 = "Show all users"
$Menu7 = "Delete a user"
$Menu99 = "Show Description"


#------------------------------------------------------------------------------
#Functions
function Show-Description()
{
  #feeding the narcistic beast
  Write-Host "# Description" -ForegroundColor whGroep1e
  Write-Host "# -----------" -ForegroundColor whGroep1e
  Write-Host "# Active Directory Management tool" -ForegroundColor yellow
  Write-Host "# Adapted for the common task Linux/Windows/Networking"
  Write-Host ""
  Write-Host "# Author : David Lejeune" -ForegroundColor magenta
  Write-Host "# Created : 27/09/2016" -ForegroundColor magenta
  Write-Host "# School : Vives" -ForegroundColor magenta
  Write-Host "# Course : Operating Systems Windows 2" -ForegroundColor magenta
  Write-Host "# Class : 3PB-ICT" -ForegroundColor magenta
  Write-Host "# ----------" -ForegroundColor red
  Write-Host ""
}

function Create-OU()
{
    #create top level OU (needs to worked out further for depth)
    $OUname = Read-Host -Prompt '> OU name ';
    New-ADOrganizationalUnGroep1 $OUname ;
}

function Create-User()
{
    #Create user based on user input
    $UserFirstname = Read-Host -Prompt '> given name ';
    $UserLastname = Read-Host -Prompt '> surname ';
    $Displayname = $UserFirstname + " " + $UserLastname;
    $SAM = Read-Host -Prompt '> SAM account name ';
    $UserpathOU = Read-Host -Prompt '> OU ';
    $UPN = "$($SAM)@Vives.be"
    $pathOU = "ou=$($UserpathOU),ou=VivesAfdelingen,dc=Vives,dc=be"
    New-ADUser -Department:"$($UserpathOU)" -DisplayName:"$($Displayname)" -GivenName:"$($UserFirstname)" -Name:"$($Displayname)" -Path:"OU=$($UserpathOU),OU=VivesAfdelingen,DC=Vives,dc=be" -SamAccountName:"$($SAM)" -Server:"CrispyMcBacon.Vives.be" -Surname:"$($UserLastname)" -Type:"user" -UserPrincipalName:"$($SAM)@Vives.be" -AccountPassword (ConvertTo-SecureString "Password123" -AsPlainText -Force) -Enabled $true
}

function Delete-User()
{
    Show-Users
    #Delete user based on user input
    $SAM = Read-Host -Prompt '> SAM account name ';

    #Check user existence
    if (dsquery user -samid $SAM)
    {
      "Found user"
      remove-aduser -identGroep1y $SAM #-confirm:$false
      Show-Users
      if (dsquery user -samid $SAM){"User unsuccesfully deleted"}
      else {"User succesfully deleted"}
    }
    else
    {
      "Did not find user"
    }
}

function Show-Users()
{
    $sw = [Diagnostics.Stopwatch]::StartNew()
    #log users and show them
    Get-ADUser -SearchBase "OU=VivesAfdelingen,dc=Vives,dc=be" -Filter * -properties * -ResultSetSize 5000 | select identGroep1y ,CN ,SAMAccountName, Department, Description , TGroep1le,UserPrincipalName, DistinguishedName, HomeDirectory, ProfilePath, Office, OfficePhone, Manager    | convertto-html | out-file ADUsers.html
    Get-ADUser -SearchBase "dc=Vives,dc=be" -Filter * -properties * -ResultSetSize 5000 | select DistinguishedName,SAMAccountName, Department | format-table -autosize
}

function Check-UserExistence()
{
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $SAM = Read-Host -Prompt '> Enter SamAccountName ';
    if (dsquery user -samid $SAM){"Found user"}
    else {"Did not find user"}
}

function Bulk-UserDelete()
{
  $sw = [Diagnostics.Stopwatch]::StartNew()

  #import data
  $Users = Import-Csv -Delimiter ";" -Path "Users.csv"

  #header of table
  Write-Host "Get ready for the magic ...`n" -ForegroundColor Gray
  Write-Host "Deleting users`n" -ForegroundColor whGroep1e
  Write-Host "SAM      `tExists?      `t`tAction     `t`t`tOU" -ForegroundColor Yellow
  Write-Host "---      `t-------   `t`t------     `t`t`t--" -ForegroundColor Yellow

  #loop through all users
  foreach ($User in $Users)
  {
      #get csv Variables
      $Displayname = $User.Voornaam + " " + $User.Naam
      $UserFirstname = $User.Naam
      $UserLastname = $User.Voornaam
      $UserAccount = $User.Account
      $SAM = $UserAccount
      $UPN = "$($SAM)@Vives.be"
      $OU = ""
      $DistinguishedName = "CN=" + $Displayname + ","

      #find the department
      $Groep1 = $User.Groep1
      $Groep2 =  $User.Groep2
      $Groep3 = $User.Groep3
      $Groep4 = $User.Groep4
      $UserpathOU = ""

          if ($Groep4 -eq "X")
          {
            $UserpathOU = "Groep4"
            $DistinguishedName = "$($DistinguishedName)OU=$($UserpathOU),"
          }
          if ($Groep3 -eq "X")
          {
            $UserpathOU = "Groep3"
            $DistinguishedName = "$($DistinguishedName)OU=$($UserpathOU),"
          }
          if ($Groep2 -eq "X")
          {
            $UserpathOU = "Groep2"
            $DistinguishedName = "$($DistinguishedName)OU=$($UserpathOU),"
          }
          if ($Groep1 -eq "X")
          {
            $UserpathOU = "Groep1"
            $DistinguishedName = "$($DistinguishedName)OU=$($UserpathOU),"
          }

      $DistinguishedName = "$($DistinguishedName)OU=VivesAfdelingen,DC=Vives,dc=be,"

        $Result = ""
        $Result2 = ""
        if (dsquery user -samid $SAM)
        {
          $Result = "User Found"
          remove-aduser -identGroep1y $SAM -confirm:$false

          #Check after deletion if user exists now
          if (dsquery user -samid $SAM)
          {
            $Result2 =  "Unsuccesfull in deleting user"
          }
          else
          {
            $Result2 =  "User succesfully deleted"
          }
        }
        else
        {
          $Result = "User not found"
          $Result2 =  "No action required"
        }
        Write-Host $SAM"      `t"$Result"`t`t"$Result2"      `t"$UserpathOU  -ForegroundColor Red
  }

  Write-Host ""
  Write-Host " *** Finished bulk deleting users *** " -ForegroundColor blue
}

function Bulk-UserManagement()
{

  # first diable the user not found in csv
  Remove-UnfoundUser
  #main task
  #create users based on csv date
  #import data
  $Users = Import-Csv -Delimiter ";" -Path "Users.csv"

  Write-Host "`nCrunching data like a boss"  -ForegroundColor red
  Write-Host "Get ready for the magic ...`n"  -ForegroundColor red
  Write-Host "Creating users`n" -ForegroundColor whGroep1e
  Write-Host "SAM      `tExists?      `t`tAction     `t`t`tOU`t`t     `tSubgroup"  -ForegroundColor yellow
  Write-Host "---      `t-------   `t`t------     `t`t`t--`t`t     `t--------" -ForegroundColor yellow

  #loop through all users
  foreach ($User in $Users)
  {
      $Displayname = $User.Voornaam + " " + $User.Naam
      $UserFirstname = $User.Naam
      $UserLastname = $User.Voornaam
      $UserAccount = $User.Account
      $SAM = $UserAccount
      $UPN = "$($SAM)@Vives.be"
      $OU = ""
      $DistinguishedName = ""
      $BossName = ""


      $Groep1 = $User.Groep1
      $Groep2 =  $User.Groep2
      $Groep3 = $User.Groep3
      $Groep4 = $User.Groep4

      $UserpathOU = ""

          if ($Groep4 -eq "X")
          {
            $UserpathOU = "Groep4"
            $DistinguishedName = "$($DistinguishedName)OU=$($UserpathOU),"
          }
          if ($Groep3 -eq "X")
          {
            $UserpathOU = "Groep3"
            $DistinguishedName = "$($DistinguishedName)OU=$($UserpathOU),"
          }
          if ($Groep2 -eq "X")
          {
            $UserpathOU = "Groep2"
            $DistinguishedName = "$($DistinguishedName)OU=$($UserpathOU),"
          }
          if ($Groep1 -eq "X")
          {
            $UserpathOU = "Groep1"
            $DistinguishedName = "$($DistinguishedName)OU=$($UserpathOU),"
          }


        $DistinguishedName = "$($DistinguishedName)OU=VivesAfdelingen,DC=Vives,dc=be,"

        $Result = ""
        $Result2 = ""

        if (dsquery user -samid $SAM)
        {
          $Result = "User Found   "

          # if user exists remove them from groups and retarget so an update puts them in the correct path
          $user = "CN=$($UserpathOU),OU=$($UserpathOU),OU=VivesAfdelingen,DC=Vives,dc=be"
          Get-ADPrincipalGroupMembership -IdentGroep1y $user | where {$_.Name -ne "Domain Users"} | % {Remove-ADPrincipalGroupMembership -IdentGroep1y $user -MemberOf $_}
          # retarget to catch updates
          Get-ADUser $SAM| Move-ADObject -TargetPath "OU=$($UserpathOU),OU=VivesAfdelingen,DC=Vives,dc=be"
          # rename department
          Get-ADUser $SAM| Set-ADUser -Department $UserpathOU

          #making sure if a former user (disabled) is found that he is set active again
          $ActiveUser = Get-ADUser -IdentGroep1y $SAM
          if ($ActiveUser.Enabled -eq $False)
          {
              Enable-ADAccount -IdentGroep1y $SAM
              $Result2 =  "Reactivated former user      "
          }
          else{

              $Result2 =  "Update user                  "
          }

        }
        else
        {
          $Result = "User not found"

          #create the user and assign to OU
          New-ADUser -ChangePasswordAtLogon:$true -Department:"$($UserpathOU)" -DisplayName:"$($Displayname)" -GivenName:"$($UserFirstname)" -Name:"$($Displayname)" -Path:"OU=$($UserpathOU),OU=VivesAfdelingen,DC=Vives,dc=be" -SamAccountName:"$($SAM)" -Server:"CrispyMcBacon.Vives.be" -Surname:"$($UserLastname)" -Type:"user" -UserPrincipalName:"$($SAM)@Vives.be" -AccountPassword (ConvertTo-SecureString "Password123" -AsPlainText -Force) -Enabled $true

          #Check after creation if user exists now
          if (dsquery user -samid $SAM)
          {
            $Result2 = "User succesfully created     "
          }
          else
          {
            $Result2 = "Unsuccesfull in creating user"
          }

          ##################################################################
          # BLOCK THAT HELPS DISPLAY ALL THE BEAUTIFUL STUFF but only for new users
           #assign to the correct principal group
           $UserpathOU = ""
           $Boss = "False"
           $countDepartments = 0

               if ($Groep4 -eq "X")
               {
                 $UserpathOU = "Groep4"
                 $DistinguishedName = "$($DistinguishedName)OU=$($UserpathOU),"
                 $countDepartments = $countDepartments + 1
               }
               if ($Groep3 -eq "X")
               {
                 $UserpathOU = "Groep3"
                 $DistinguishedName = "$($DistinguishedName)OU=$($UserpathOU),"
                 $countDepartments = $countDepartments + 1
               }
               if ($Groep2 -eq "X")
               {
                 $UserpathOU = "Groep2"
                 $DistinguishedName = "$($DistinguishedName)OU=$($UserpathOU),"
                 $countDepartments = $countDepartments + 1
               }
               if ($Groep1 -eq "X")
               {
                 $UserpathOU = "Groep1"
                 $DistinguishedName = "$($DistinguishedName)OU=$($UserpathOU),"
                 $countDepartments = $countDepartments + 1
               }
               Set-ADGroup -Add:@{'Member'="CN=$($Displayname),OU=$($UserpathOU),OU=VivesAfdelingen,DC=Vives,dc=be"} -IdentGroep1y:"CN=$($UserpathOU),OU=$($UserpathOU),OU=VivesAfdelingen,DC=Vives,dc=be" -Server:"CrispyMcBacon.Vives.be"


           #assigning to possible (sub) groups
           $UserpathOU = ""
           $Boss = "False"
           $SubOU = ""

               if ($Groep4 -eq "X")
               {
                 $UserpathOU = "Groep4"
                 $DistinguishedName = "$($DistinguishedName)OU=$($UserpathOU),"
               }
               if ($Groep3 -eq "X")
               {
                 $UserpathOU = "Groep3"
                 $DistinguishedName = "$($DistinguishedName)OU=$($UserpathOU),"
               }
               if ($Groep2 -eq "X")
               {
                 $UserpathOU = "Groep2"
                 $DistinguishedName = "$($DistinguishedName)OU=$($UserpathOU),"
               }
               if ($Groep1 -eq "X")
               {
                 $UserpathOU = "Groep1"
                 $DistinguishedName = "$($DistinguishedName)OU=$($UserpathOU),"
               }
               Add-ADPrincipalGroupMembership -IdentGroep1y:"CN=$($Displayname),OU=$($UserpathOU),OU=VivesAfdelingen,DC=Vives,dc=be" -MemberOf:"CN=$($UserpathOU),OU=$($UserpathOU),OU=VivesAfdelingen,DC=Vives,dc=be" -Server:"CrispyMcBacon.Vives.be"



         # END OF BLOCK THAT HELPS DISPLAY ALL THE BEAUTIFUL STUFF for new users
         ##################################################################
        }
        If ($Result -eq "User not found")
        {
          Write-Host "$($SAM)      `t$($Result)`t`t$($Result2)`t$($UserpathOU)     `t`t$($SubOU)" -ForegroundColor cyan
        }
        else
        {
          Write-Host "$($SAM)      `t$($Result)`t`t$($Result2)`t$($UserpathOU)     `t`t$($SubOU)" -ForegroundColor Magenta

        }
  }
  Write-Host ""
  Write-Host " *** Finished creating new users and adding them to the correct OU *** `n"-ForegroundColor blue
  Clear-Groups
  Set-Group

}

function Remove-GroupMembershipDL()
{
  $SearchBase = "OU=$($UserpathOU),OU=VivesAfdelingen,DC=Vives,dc=be"
  $Users = Get-ADUser -filter * -SearchBase $SearchBase -Properties MemberOf
  ForEach($User in $Users){
      if ($User.Enabled -eq $True) {
        $User.MemberOf | Remove-ADGroupMember -Member $User -Confirm:$false
        $Count=$Count+1
      }
      else
      {

      }
  }
  Write-Host "Removed $($Count) user(s) from $($UserpathOU)" -ForegroundColor red
}

function Clear-Groups()
{
  Write-Host "Removing all active users from groups" -ForegroundColor whGroep1e;
  Write-Host "" -ForegroundColor yellow;

  #Choose Organizational UnGroep1
  $Count=0
  $UserpathOU="Groep1"
  Remove-GroupMembershipDL

  #Choose Organizational UnGroep1
  $Count=0
  $UserpathOU="Groep2"
  Remove-GroupMembershipDL

  #Choose Organizational UnGroep1
  $Count=0
  $UserpathOU="Groep3"
  Remove-GroupMembershipDL

  #Choose Organizational UnGroep1
  $Count=0
  $UserpathOU="Groep4"
  Remove-GroupMembershipDL


  #Choose Organizational UnGroep1
  $Count=0
  $UserpathOU="Management"
  Remove-GroupMembershipDL


  Write-Host ""
  Write-Host " *** Finished clearing all users in groups *** `n" -ForegroundColor blue
}


function Set-Group()
{
  #import data
  $Users = Import-Csv -Delimiter ";" -Path "Users.csv"

  Write-Host "Setting Group(s) for users`n" -ForegroundColor whGroep1e
  Write-Host "SAM      `tGroup/OU     `t`tSubgroup" -ForegroundColor yellow
  Write-Host "---      `t--------     `t`t--------" -ForegroundColor yellow
  #loop through all users
  foreach ($User in $Users)
  {
      $Displayname = $User.Voornaam + " " + $User.Naam
      $UserFirstname = $User.Naam
      $UserLastname = $User.Voornaam
      $UserAccount = $User.Account
      $SAM = $UserAccount
      $UPN = "$($SAM)@Vives.be"
      $OU = ""
      $DistinguishedName = ""
      $BossName = ""


      $Groep1 = $User.Groep1
      $Groep2 =  $User.Groep2
      $Groep3 = $User.Groep3
      $Groep4 = $User.Groep4

      $DistinguishedName = "$($DistinguishedName)OU=VivesAfdelingen,DC=POLIFORMA,dc=be,"

        $Result = ""
        $Result2 = ""

        if (dsquery user -samid $SAM)
        {
          $Result = ""

          #assign to the correct principal group
          $UserpathOU = ""
          $Boss = "False"
          $countDepartments = 0


              if ($Groep4 -eq "X")
              {
                $UserpathOU = "Groep4"
                $DistinguishedName = "$($DistinguishedName)OU=$($UserpathOU),"
                $countDepartments = $countDepartments + 1
              }
              if ($Groep3 -eq "X")
              {
                $UserpathOU = "Groep3"
                $DistinguishedName = "$($DistinguishedName)OU=$($UserpathOU),"
                $countDepartments = $countDepartments + 1
              }
              if ($Groep2 -eq "X")
              {
                $UserpathOU = "Groep2"
                $DistinguishedName = "$($DistinguishedName)OU=$($UserpathOU),"
                $countDepartments = $countDepartments + 1
              }
              if ($Groep1 -eq "X")
              {
                $UserpathOU = "Groep1"
                $DistinguishedName = "$($DistinguishedName)OU=$($UserpathOU),"
                $countDepartments = $countDepartments + 1
              }
              Set-ADGroup -Add:@{'Member'="CN=$($Displayname),OU=$($UserpathOU),OU=VivesAfdelingen,DC=Vives,dc=be"} -IdentGroep1y:"CN=$($UserpathOU),OU=$($UserpathOU),OU=VivesAfdelingen,DC=Vives,dc=be" -Server:"CrispyMcBacon.Vives.be"
              $Result = $UserpathOU


          #assigning to possible (sub) groups
          $UserpathOU = ""
          $Boss = "False"
          $SubOU = ""

              if ($Groep4 -eq "X")
              {
                $UserpathOU = "Groep4"
                $DistinguishedName = "$($DistinguishedName)OU=$($UserpathOU),"
              }
              if ($Groep3 -eq "X")
              {
                $UserpathOU = "Groep3"
                $DistinguishedName = "$($DistinguishedName)OU=$($UserpathOU),"
              }
              if ($Groep2 -eq "X")
              {
                $UserpathOU = "Groep2"
                $DistinguishedName = "$($DistinguishedName)OU=$($UserpathOU),"
              }
              if ($Groep1 -eq "X")
              {
                $UserpathOU = "Groep1       "
                $DistinguishedName = "$($DistinguishedName)OU=$($UserpathOU),"
              }
              Add-ADPrincipalGroupMembership -IdentGroep1y:"CN=$($Displayname),OU=$($UserpathOU),OU=VivesAfdelingen,DC=Vives,dc=be" -MemberOf:"CN=$($UserpathOU),OU=$($UserpathOU),OU=VivesAfdelingen,DC=Vives,dc=be" -Server:"CrispyMcBacon.Vives.be"

    }
        Write-Host "$($SAM)      `t$($Result)  `t`t$($Result2)"  -ForegroundColor DarkGreen
  }
  Write-Host ""
  Write-Host " *** Finished adding users to the correct group(s) *** `n" -ForegroundColor blue

}

function Remove-UnfoundUser()
{

    #import data
    $Users = Import-Csv -Delimiter ";" -Path "Users.csv"

    Write-Host "Disabling users not found in CSV`n" -ForegroundColor whGroep1e
    Write-Host "SAM           `t`t`tAction" -ForegroundColor yellow
    Write-Host "---           `t`t`t------" -ForegroundColor yellow


    $usersAD = Get-ADUser -SearchBase "ou=VivesAfdelingen,dc=Vives,dc=be" -filter *
     ForEach($userAD in $usersAD)
        {
          $SAM_AD = $userAD.SAMAccountName
          $UsersCSV = Import-Csv -Delimiter ";" -Path "Users.csv"
          $found=$false
          #loop through all users
              foreach ($User in $UsersCSV)
              {
                  $Displayname = $User.Voornaam + " " + $User.Naam
                  $UserFirstname = $User.Naam
                  $UserLastname = $User.Voornaam
                  $UserAccount = $User.Account
                  $SAM = $UserAccount
                  if ($SAM_AD -eq $SAM)
                  {
                    $found=$true
                  }

              }


              if ($found -eq $true)
              {
                if (dsquery user -samid $SAM_AD)
                {
                  $ActiveUser = Get-ADUser -IdentGroep1y $SAM_AD
                  if ($ActiveUser.Enabled -eq $True)
                  {
                    $Result  = "User remains active"
                    Write-host $SAM_AD"     `t`t`t"$Result -ForegroundColor whGroep1e
                  }
                  else
                  {
                    $Result  = "Old user reactivated"
                    Write-host $SAM_AD"     `t`t`t"$Result -ForegroundColor darkgray
                  }

                }
                else
                {
                  $Result  = "New user detected"
                  Write-host $SAM_AD"     `t`t`t"$Result -ForegroundColor magenta
                }
              }
              else
              {
                $ActiveUser = Get-ADUser -IdentGroep1y $SAM_AD
                if ($ActiveUser.Enabled -eq $True)
                {
                    Disable-ADaccount -identGroep1y $SAM_AD
                    $Result  = "User has been disabled"
                    Write-host $SAM_AD"     `t`t`t"$Result -ForegroundColor  Red
                }
                else{
                    Disable-ADaccount -identGroep1y $SAM_AD
                    $Result  = "User remains disabled"
                    Write-host $SAM_AD"     `t`t`t"$Result -ForegroundColor  Green

                }

              }


        }



}

function Set-Manager()
{
  #import data
  $Users = Import-Csv -Delimiter ";" -Path "Users.csv"

  Write-Host "Setting manager for users in OU's`n" -ForegroundColor whGroep1e
  Write-Host "SAM      `tManager of   `t`t`tAction" -ForegroundColor yellow
  Write-Host "---      `t----------   `t`t`t------" -ForegroundColor yellow

  $manManager = ""
  $manGroep1 = ""
  $manGroep2 = ""
  $manGroep4 = ""
  $manGroep3 = ""

  #loop through all users
  foreach ($User in $Users)
  {
      $Displayname = $User.Voornaam + " " + $User.Naam
      $UserFirstname = $User.Naam
      $UserLastname = $User.Voornaam
      $UserAccount = $User.Account
      $SAM = $UserAccount
      $UPN = "$($SAM)@Vives.be"
      $OU = ""
      $DistinguishedName = ""
      $BossName = ""

      $Manager = $User.Manager
      $Groep1 = $User.Groep1
      $Groep2 =  $User.Groep2
      $Groep3 = $User.Groep3
      $Groep4 = $User.Groep4

      $UserpathOU = ""
      $DistinguishedName = "$($DistinguishedName)OU=VivesAfdelingen,DC=POLIFORMA,dc=be,"

        $Result = ""
        $Result2 = ""
        $countDepartments = 0

          $Boss = "False"
          $SubOU = ""

          if ($Manager -eq "X")
          {
            $UserpathOU = "Management"
            $DistinguishedName = "$($DistinguishedName)OU=$($UserpathOU),"
            $Boss = "True"
            $countDepartments = $countDepartments + 1

            if ($Groep4 -eq "X")
            {
              $SubOU = "Groep4"
              $DistinguishedName = "OU=$($UserpathOU),"
              $countDepartments = $countDepartments + 1
              #Set-ADGroup -IdentGroep1y:"CN=Manager,OU=Management,OU=VivesAfdelingen,DC=Vives,dc=be" -ManagedBy:"CN=Bert Laplasse,OU=Management,OU=VivesAfdelingen,DC=Vives,dc=be" -Server:"CrispyMcBacon.Vives.be"
              $manGroep4 = $Displayname
              Get-ADUser -SearchBase "OU=$($SubOU),OU=VivesAfdelingen,dc=Vives,dc=be" -Filter * -properties * -ResultSetSize 5000 | Set-ADUser  -Manager "$($SAM)" #-IdentGroep1y:"$($_.SAMAccountName)" -Manager  #-Server:"CrispyMcBacon.Vives.be"
              $Result = "    `tSet as manager for all users in OU"
            }

            if ($Groep3 -eq "X")
            {
              $SubOU = "Groep3"
              $DistinguishedName = "OU=$($UserpathOU),"
              $countDepartments = $countDepartments + 1
              $manGroep3 = $Displayname
              Get-ADUser -SearchBase "OU=$($SubOU),OU=VivesAfdelingen,dc=Vives,dc=be" -Filter * -properties * -ResultSetSize 5000 | Set-ADUser  -Manager "$($SAM)" #-IdentGroep1y:"$($_.SAMAccountName)" -Manager  #-Server:"CrispyMcBacon.Vives.be"
              $Result = "Set as manager for all users in OU"
                $SubOU = "Groep3`t"
            }
            if ($Groep2 -eq "X")
            {
              $SubOU = "Groep2"
              $DistinguishedName = "OU=$($UserpathOU),"
              $countDepartments = $countDepartments + 1
              $manGroep2 = $Displayname
              Get-ADUser -SearchBase "OU=$($SubOU),OU=VivesAfdelingen,dc=Vives,dc=be" -Filter * -properties * -ResultSetSize 5000 | Set-ADUser  -Manager "$($SAM)" #-IdentGroep1y:"$($_.SAMAccountName)" -Manager  #-Server:"CrispyMcBacon.Vives.be"
              $Result = "Set as manager for all users in OU"
                $SubOU = "Groep2`t"
            }
            if ($Groep1 -eq "X")
            {
              $SubOU = "Groep1"
              $DistinguishedName = "OU=$($UserpathOU),"
              $countDepartments = $countDepartments + 1
              $manGroep1 = $Displayname
              Get-ADUser -SearchBase "OU=$($SubOU),OU=VivesAfdelingen,dc=Vives,dc=be" -Filter * -properties * -ResultSetSize 5000 | Set-ADUser  -Manager "$($SAM)" #-IdentGroep1y:"$($_.SAMAccountName)" -Manager  #-Server:"CrispyMcBacon.Vives.be"
              $Result = "Set as manager for all users in OU"
                $SubOU = "Groep1         `t"
            }
          }
          else
          {
          }

          if ($Manager -eq "X")
          {
            if ($countDepartments -eq 1)
              {
                $SubOU = "Management"
                $manManager = $Displayname
                #Set-ADUser -IdentGroep1y:"CN=Linda Hombroeckx,OU=Groep3,OU=VivesAfdelingen,DC=Vives,dc=be" -Replace:'manager'="CN=Bert Laplasse,OU=Management,OU=VivesAfdelingen,DC=Vives,dc=be" -Server:"CrispyMcBacon.Vives.be"
                #Set-ADUser -IdentGroep1y:"CN=Linda Hombroeckx,OU=Groep3,OU=VivesAfdelingen,DC=Vives,dc=be" -Manager:$null -Server:"CrispyMcBacon.Vives.be"
                #Get-ADUser -SearchBase "OU=$($UserpathOU),dc=Vives,dc=be" -Filter * -ResultSetSize 5000 | Select Name,SamAccountName
                #Write-Host "Setting manager $($Displayname) for users in $($SubOU) "
                #Get-ADUser -SearchBase "OU=$($UserpathOU),OU=VivesAfdelingen,dc=Vives,dc=be" -Filter * -properties * -ResultSetSize 5000 | select SAMAccountName # Set-ADUser -IdentGroep1y:"CN=Linda Hombroeckx,OU=Groep3,OU=VivesAfdelingen,DC=Vives,dc=be" -Manager:$null -Server:"CrispyMcBacon.Vives.be"
                Get-ADUser -SearchBase "OU=$($SubOU),OU=VivesAfdelingen,dc=Vives,dc=be" -Filter * -properties * -ResultSetSize 5000 | Set-ADUser  -Manager "$($SAM)" #-IdentGroep1y:"$($_.SAMAccountName)" -Manager  #-Server:"CrispyMcBacon.Vives.be"
                $Result = "Set as manager for all users in OU"
                $SubOU = "Management`t"
              }
          }

          #show only if is a boss
          if ($SubOU -eq "")
          {}
            else{
              Write-Host "$($SAM)      `t$($SubOU)`t`t$($Result)"  -ForegroundColor Magenta
            }
  }
  Write-Host ""
  Write-Host " *** Finished setting managers for all users *** " -ForegroundColor Blue
}


function Log-Action()
{
  $Date = Get-Date
  $Entry = $Date.ToString() + "," + $env:username.ToString() + "," + $Menu.ToString() + ","+ $time_elapsed + ","
  Add-Content script_logbook.csv $Entry
}

function Show-Header()
{
  #making this script sexy
    Clear
    Write-Host '      ____              __        ' -ForegroundColor Yellow
    Write-Host '     / __ \   ____ _   / /      ___ ' -ForegroundColor Yellow
    Write-Host '    / / / /  / __ `/  / /      / _ \' -ForegroundColor Yellow
    Write-Host '   / /_/ /  / /_/ /  / /___   /  __/' -ForegroundColor Yellow
    Write-Host '  /_____/   \__,_/  /_____/   \___/ ' -ForegroundColor Yellow
    Write-Host ''
    Write-Host '    +-+-+-+-+-+-+-+-+-+-+ +-+-+-+' -ForegroundColor Blue
    Write-Host '    |P|o|w|e|r|s|h|e|l|l| |C|L|I|' -ForegroundColor Blue
    Write-Host '    +-+-+-+-+-+-+-+-+-+-+ +-+-+-+' -ForegroundColor Blue
    Write-Host ''
    Write-Host '  >> Author : David Lejeune' -ForegroundColor Red
    Write-Host "  >> Created : 27/09/2016" -ForegroundColor Red
    Write-Host ''
    Write-Host ' #####################################'  -ForegroundColor DarkGreen
    Write-Host ' #    ACTIVE DIRECTORY MANAGEMENT    #' -ForegroundColor DarkGreen
    Write-Host ' #####################################' -ForegroundColor DarkGreen
    Write-Host ''
}

function Show-Menu()
{
  #making this script sexy
    Write-Host " Menu :" -ForegroundColor Magenta;
    Write-Host "";
    Write-Host '    1. '$Menu1  -ForegroundColor Gray;
    Write-Host '    2. '$Menu2 -ForegroundColor Gray;
    Write-Host '    3. '$Menu3 -ForegroundColor Magenta;
    Write-Host '    4. '$Menu4 -ForegroundColor Gray;
    Write-Host '    5. '$Menu5 -ForegroundColor Red;
    Write-Host '    6. '$Menu6 -ForegroundColor DarkMagenta;
    Write-Host '    7. '$Menu7 -ForegroundColor DarkGreen;
    Write-Host '   ';
    Write-Host '    99.'$Menu99 -ForegroundColor DarkGRay;
    Write-Host "";
}

#------------------------------------------------------------------------------
#Script

$Host.UI.RawUI.BackgroundColor = ($bckgrnd = 'black')
$Host.UI.RawUI.ForegroundColor = 'WhGroep1e'
$Host.PrivateData.ErrorForegroundColor = 'Red'
$Host.PrivateData.ErrorBackgroundColor = $bckgrnd
$Host.PrivateData.WarningForegroundColor = 'Magenta'
$Host.PrivateData.WarningBackgroundColor = $bckgrnd
$Host.PrivateData.DebugForegroundColor = 'Yellow'
$Host.PrivateData.DebugBackgroundColor = $bckgrnd
$Host.PrivateData.VerboseForegroundColor = 'Green'
$Host.PrivateData.VerboseBackgroundColor = $bckgrnd
$Host.PrivateData.ProgressForegroundColor = 'Cyan'
$Host.PrivateData.ProgressBackgroundColor = $bckgrnd


Show-Header;
Show-Menu;

#Select action
$Menu = Read-Host -Prompt 'Select an option ';
$sw = [Diagnostics.Stopwatch]::StartNew()
swGroep1ch ($Menu)
    {
        1
          {
              Write-Host "`nYou have selected $(($Menu1).ToUpper())`n" -ForegroundColor DarkGreen;
              $Menu = $Menu1;
              Create-OU;
          }

        2
          {
              Write-Host "`nYou have selected $(($Menu2).ToUpper())`n" -ForegroundColor DarkGreen;
              $Menu = $Menu2;
              Create-User;
          }

        3
          {
              Write-Host "`nYou have selected $(($Menu3).ToUpper())`n" -ForegroundColor DarkGreen;
              $Menu = $Menu3;
              Bulk-UserManagement;
          }

        4
          {
              Write-Host "`nYou have selected $(($Menu4).ToUpper())`n" -ForegroundColor DarkGreen;
              $Menu = $Menu4;
              Check-UserExistence;
          }
        5
          {
              Write-Host "`nYou have selected $(($Menu5).ToUpper())`n" -ForegroundColor DarkGreen;
              $Menu = $Menu5;
              Bulk-UserDelete;
          }
        6
          {
              Write-Host "`nYou have selected $(($Menu6).ToUpper())`n" -ForegroundColor DarkGreen;
              $Menu = $Menu6;
              Show-Users;
          }
        7
          {
              Write-Host "`nYou have selected $(($Menu7).ToUpper())`n" -ForegroundColor DarkGreen;
              $Menu = $Menu7;
              Delete-User;
          }
        99
          {
              Write-Host "`nYou have selected $(($Menu99).ToUpper())`n" -ForegroundColor DarkGreen;
              $Menu = $Menu99;
              Show-Description;
          }

        default {
          Write-Host "The choice could not be determined." -ForegroundColor Red
        }
    }

    $sw.Stop()
    $time_elapsed = $sw.Elapsed.TotalSeconds
    Write-Host " *** Task completed in "$time_elapsed" seconds. ***" -ForegroundColor Yellow
    Log-Action
#Clear-Host
