<#
    .SYNOPSIS 
    Configures Entra ID Authentication Methods 
    .DESCRIPTION
    This Script will set the Entra ID Authentication methods for a Tenant with the Specified groups for the Methods      
    .EXAMPLE
    Set-CLCAuthenticationMethods
    .NOTES
        AUTHOR: Andres-Miguel Sichel
        LASTEDIT: 28/06/2024
        KEYWORDS: EntraID, Authentications
    .Link
        https://github.com/MuedenJoe/andisitblog
        https://learn.microsoft.com/en-us/graph/api/resources/authenticationmethods-overview?view=graph-rest-1.0
#>

#MSGraph Authentication Scope and Connect
Connect-MgGraph -Scopes UserAuthenticationMethod.ReadWrite.All,Policy.ReadWrite.AuthenticationMethod,Group.ReadWrite.All

#Query for Group IDs

<#
M365-AAD-AP_MSAuthenticator-Users
M365-AAD-AP_FIDO-Users
M365-AAD-AP_TAP-Users
M365-AAD-AP_SMS-Users
M365-AAD-AP_HardwareOAUTH-Users
M365-AAD-AP_3rdPartyOAUTH-Users
M365-AAD-AP_Voicecall-Users
M365-AAD-AP_EmailOTP-Users
M365-AAD-AP_CBA-Users
#>

#Microsoft Authenticator Group ID
$MSAuthenticatorGroupID = (Get-MgGroup -Filter "DisplayName eq 'M365-AAD-AP_MSAuthenticator-Users'").ID

#Fido Group ID
$FidoGroupID = (Get-MgGroup -Filter "DisplayName eq 'M365-AAD-AP_FIDO-Users'").ID

#TAP Group ID
$TAPGroupID = (Get-MgGroup -Filter "DisplayName eq 'M365-AAD-AP_TAP-Users'").ID

#SMS Group ID
$SMSGroupID = (Get-MgGroup -Filter "DisplayName eq 'M365-AAD-AP_SMS-Users'").ID

#Hardware Token Group ID - not supported yet
$HardwareTokenGroupID = (Get-MgGroup -Filter "DisplayName eq 'M365-AAD-AP_HardwareOAUTH-Users'").ID

#3rd Party Token
$3rdPartyTokenGroupID = (Get-MgGroup -Filter "DisplayName eq 'M365-AAD-AP_3rdPartyOAUTH-Users'").ID

#Voicecall Users
$VoicecallGroupID = (Get-MgGroup -Filter "DisplayName eq 'M365-AAD-AP_Voicecall-Users'").ID

#Email OTP
$EmailOTPGroupID = (Get-MgGroup -Filter "DisplayName eq 'M365-AAD-AP_EmailOTP-Users'").ID

#Certificate Based
$CBAGroupID = (Get-MgGroup -Filter "DisplayName eq 'M365-AAD-AP_CBA-Users'").ID


#----------------------------------------------------------------------------------------------------------------------#
# Setup Microsoft Authenticator
function Set-CLCEntraIDAPMSAuthenticator {

$authenticationMethodConfigurationId = "MicrosoftAuthenticator"

$featureSettings = @{
    displayAppInformationRequiredState = @{
        state = "enabled"
        includeTarget = @{
            targetType = "group"
            id = "all_users"
        }
        excludeTarget = @{
            targetType = "group"
            id = "00000000-0000-0000-0000-000000000000"
        }
    }
    displayLocationInformationRequiredState = @{
        state = "enabled"
        includeTarget = @{
            targetType = "group"
            id = "bd72009e-587b-46c8-b4bf-82f56a74a65f"
        }
        excludeTarget = @{
            targetType = "group"
            id = "00000000-0000-0000-0000-000000000000"
        }
    }
    

}


$params = @{
    "@odata.type" = "#microsoft.graph.microsoftAuthenticatorAuthenticationMethodConfiguration"
    id = $authenticationMethodConfigurationId
    state = "enabled"
    includeTargets  = @(
        @{
            id         = "$MSAuthenticatorGroupID"
            targetType = "group"
        }
    
    )
    #for exclusion of specified group comment in
    <#
    excludeTargets  = @(
        @{
            id         = "$FidoGroupID"
            targetType = "group"
        }
    )
    #>
    featureSettings = $featureSettings
    
}

Update-MgPolicyAuthenticationMethodPolicyAuthenticationMethodConfiguration `
    -AuthenticationMethodConfigurationId $authenticationMethodConfigurationId `
    -AdditionalProperties $params

}

#----------------------------------------------------------------------------------------------------------------------#
# Setup Fido Authentications
function Set-CLCEntraIDAPFido2 {

    $authenticationMethodConfigurationId = "Fido2"
    
    $params = @{
        "@odata.type"                    = "#microsoft.graph.fido2AuthenticationMethodConfiguration"
        id                               = "Fido2"
        state                            = "enabled"
        includeTargets                   = @(
            @{
                id         = "$FidoGroupID"
                targetType = "group"
            }
        )
        excludeTargets                   = @(
        )
        isSelfServiceRegistrationAllowed = $true
        isAttestationEnforced            = $false
        keyRestrictions                  = @{
            isEnforced      = $true
            enforcementType = "Allow"
            aaGuids         = @(
                "90a3ccdf-635c-4729-a248-9b709135078f", #iOS Microsoft Authenticator
                "257fa02a-18f3-4e34-8174-95d454c2e9ad", #iOS Microsoft Authenticator
                "de1e552d-db1d-4423-a619-566b625cdc84", #Android Microsoft Authenticator
                "b6879edc-2a86-4bde-9c62-c1cac4a8f8e5", #Android Microsoft Authenticator
                "a25342c0-3cdc-4414-8e46-f4807fca511c", #Yubikey 5c NFC Firmware 5.7
                "2fc0579f-8113-47ea-b116-bb5a8db9202a"  #Yubikey 5C NFC Firmware 5.2 & 5.4
            )
        }
    }
    
    Update-MgPolicyAuthenticationMethodPolicyAuthenticationMethodConfiguration `
        -AuthenticationMethodConfigurationId $authenticationMethodConfigurationId `
        -AdditionalProperties $params
    
    }

#----------------------------------------------------------------------------------------------------------------------#
# Setup Temporary Access Pass
function Set-CLCEntraIDAPTAP {

        $authenticationMethodConfigurationId = "TemporaryAccessPass"
        
        $params = @{
            "@odata.type"                    = "#microsoft.graph.temporaryAccessPassAuthenticationMethodConfiguration"
            id                               = "TemporaryAccessPass"
            state                            = "enabled"
            includeTargets                   = @(
                @{
                    id         = "$TAPGroupID"
                    targetType = "group"
                }
            )
            excludeTargets                   = @(
            )
            idefaultLifetimeInMinutes = 120
            defaultLength = 10
            minimumLifetimeInMinutes = 60
            maximumLifetimeInMinutes = 480
            isUsableOnce = $false
        }
        
        Update-MgPolicyAuthenticationMethodPolicyAuthenticationMethodConfiguration `
            -AuthenticationMethodConfigurationId $authenticationMethodConfigurationId `
            -AdditionalProperties $params
        
        }

#----------------------------------------------------------------------------------------------------------------------#
# Setup SMS Authentication Method
function Set-CLCEntraIDAPSMS {

            $authenticationMethodConfigurationId = "SMS"
            
            $params = @{
                "@odata.type"                    = "#microsoft.graph.smsAuthenticationMethodConfiguration"
                id                               = "SMS"
                state                            = "enabled"
                includeTargets                   = @(
                    @{
                        id         = "$SMSGroupID"
                        targetType = "group"
                    }
                )
                excludeTargets                   = @(
                )
                #idefaultLifetimeInMinutes = 120
                #defaultLength = 10
                #minimumLifetimeInMinutes = 60
                #maximumLifetimeInMinutes = 480
                #isUsableOnce = $false
            }
            
            Update-MgPolicyAuthenticationMethodPolicyAuthenticationMethodConfiguration `
                -AuthenticationMethodConfigurationId $authenticationMethodConfigurationId `
                -AdditionalProperties $params
            
            }
        
#----------------------------------------------------------------------------------------------------------------------#
# Setup Software Token Authentication Method
function Set-CLCEntraIDAPSoftwareToken {

    $authenticationMethodConfigurationId = "softwareOath"
    
    $params = @{
        "@odata.type"                    = "#microsoft.graph.softwareOathAuthenticationMethodConfiguration"
        id                               = "softwareOath"
        state                            = "enabled"
        includeTargets                   = @(
            @{
                id         = "$3rdPartyTokenGroupID"
                targetType = "group"
            }
        )
        excludeTargets                   = @(
        )
        #idefaultLifetimeInMinutes = 120
        #defaultLength = 10
        #minimumLifetimeInMinutes = 60
        #maximumLifetimeInMinutes = 480
        #isUsableOnce = $false
    }
    
    Update-MgPolicyAuthenticationMethodPolicyAuthenticationMethodConfiguration `
        -AuthenticationMethodConfigurationId $authenticationMethodConfigurationId `
        -AdditionalProperties $params
    
    }

#----------------------------------------------------------------------------------------------------------------------#
# Setup voice call Authentication Method
function Set-CLCEntraIDAPVoiceCall {

    $authenticationMethodConfigurationId = "Voice"
    
    $params = @{
        "@odata.type"                    = "#microsoft.graph.voiceAuthenticationMethodConfiguration"
        id                               = "Voice"
        state                            = "enabled"
        includeTargets                   = @(
            @{
                id         = "$VoicecallGroupID"
                targetType = "group"
            }
        )
        excludeTargets                   = @(
        )
        isOfficePhoneAllowed = $true
        #defaultLength = 10
        #minimumLifetimeInMinutes = 60
        #maximumLifetimeInMinutes = 480
        #isUsableOnce = $false
    }
    
    Update-MgPolicyAuthenticationMethodPolicyAuthenticationMethodConfiguration `
        -AuthenticationMethodConfigurationId $authenticationMethodConfigurationId `
        -AdditionalProperties $params
    
    }

    #----------------------------------------------------------------------------------------------------------------------#
# Setup Email Authentication Method
function Set-CLCEntraIDAPEmail {

    $authenticationMethodConfigurationId = "Email"
    
    $params = @{
        "@odata.type"                    = "#microsoft.graph.emailAuthenticationMethodConfiguration"
        id                               = "Email"
        state                            = "enabled"
        includeTargets                   = @(
            @{
                id         = "$EmailOTPGroupID"
                targetType = "group"
            }
        )
        excludeTargets                   = @(
        )
        allowExternalIdToUseEmailOtp = $true
        #defaultLength = 10
        #minimumLifetimeInMinutes = 60
        #maximumLifetimeInMinutes = 480
        #isUsableOnce = $false
    }
    
    Update-MgPolicyAuthenticationMethodPolicyAuthenticationMethodConfiguration `
        -AuthenticationMethodConfigurationId $authenticationMethodConfigurationId `
        -AdditionalProperties $params
    
    }


Set-CLCEntraIDAPMSAuthenticator
Set-CLCEntraIDAPFido2
Set-CLCEntraIDAPTAP
Set-CLCEntraIDAPSMS
Set-CLCEntraIDAPSoftwareToken
Set-CLCEntraIDAPVoiceCall
Set-CLCEntraIDAPEmail

#region disconnect
try{Disconnect-MgGraph -ErrorAction SilentlyContinue}catch{}
#endregion