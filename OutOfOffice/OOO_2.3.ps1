<#
 Exchange Server / O365 Out of Office Tool 2.3                             
 © Andres Sichel // asichel.de // blog@asichel.de                   
 Hilft bei der Konfiguration von Abwesenheitsnachrichten                 
 für Benutzer mit Mailboxen auf einem Exchange Server                    
 2010 / 2013  / 2016 / O365                                                         
 
 Changelog 2.3:
 - Sprachauswahl der GUI (EN oder DE) - Language selection (EN or DE)
 - Code aufgeräumt
 - Liste mit allen Mailboxen mit aktiven Out Of Office

 Changelog 2.2:                                                          
 - Benutzerliste wird geldaen (Nur Mailboxen von 2007 & höher werden berücksichtig) - Fehlerbeseitigung und ProgressBar                                         
 - Neues Icon                                                            
 - Variablen mit Start-OU für Benutzerlisten  (Settings.ini)
 - Weiterleitung kann konfiguriert werden
 - Login nach Office 365 implementiert (Settings.ini)                       
 - Start der Anwendung optimiert, nur noch .Net4 oder höher                                       
 - Code aufgeräum / optimiert                                            
 - Zugehörige Dateien:                                                    
   - GvS.Controls.HtmlTextbox.dll                                        
   - icon.ico                                                            
   - settings.ini                                                        
   - exit.png                                                            
   - OOO_2.2.exe
   - aktiv.png
   - inaktiv.png                                          

 
 Changelog 2.1:                                                          
 - Benutzerliste wird geldaen (Nur Mailboxen von 2007 & höher werden berücksichtig) - Fehlerbeseitigung und ProgressBar                                         
 - Benutzerliste kann auf Knopfdruck neu erzeugt werden um evtl. neue Mailboxen vorzuschlagen                                               
 - Tool kann in die Infosymbolleiste minimiert werden                    
 - Neues Icon                                                            
 - Variablen mit Servernamen und anmeldung  (Settings.ini)                       
 - Start der Anwendung optimiert - Je nach Anzahl der Benutzer kann der Start dauern!                                          
 - Code aufgeräum / optimiert                                            
 - Zugehörige Dateien:                                                    
   - GvS.Controls.HtmlTextbox.dll                                        
   - icon.ico                                                            
   - settings.ini                                                        
   - exit.png                                                            
   - OOO_2.2.exe                                          

#>

# Ausführungspfad wichtig für die HTML DLL GvS.controls muss im Pfad vom Script liegen
$scriptRoot = [System.AppDomain]::CurrentDomain.BaseDirectory.TrimEnd('\')
if ($scriptRoot -eq $PSHOME.TrimEnd('\'))
{
    $scriptRoot = $PSScriptRoot
}

#########################################################################################################################################
# Systembibliotheken laden

[Reflection.Assembly]::LoadFile("$scriptRoot\GvS.Controls.HtmlTextbox.dll") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][reflection.assembly]::Load("mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
[void][reflection.assembly]::Load("System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
[void][reflection.assembly]::Load("System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
[void][reflection.assembly]::Load("System.Data, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
[void][reflection.assembly]::Load("System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
[void][reflection.assembly]::Load("System.Xml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
[void][reflection.assembly]::Load("System.DirectoryServices, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
[void][reflection.assembly]::Load("System.Core, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
[void][reflection.assembly]::Load("System.ServiceProcess, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")

#########################################################################################################################################
#INIfile laden
function Get-IniContent ($scriptRoot)
{
    $ini = @{}
    switch -regex -file $scriptRoot
    {
        "^\[(.+)\]" # Section
        {
            $section = $matches[1]
            $ini[$section] = @{}
            $CommentCount = 0
        }
        "^(;.*)$" # Comment
        {
            $value = $matches[1]
            $CommentCount = $CommentCount + 1
            $name = "Comment" + $CommentCount
            $ini[$section][$name] = $value
        } 
        "(.+?)\s*=(.*)" # Key
        {
            $name,$value = $matches[1..2]
            $ini[$section][$name] = $value
        }
    }
    return $ini
}

$globalsettingsfile = "$scriptRoot\settings.ini"
$inifile = get-inicontent "$globalsettingsfile"

#########################################################################################################################################
#Werte aus der Settings.INI einlesen und weitere globale Variablen

# Yes & No und der Haken "Anmledung mit aktuellem benutzer" ist bereits beim start des Programms gesetzt - Wird bei O365 ingoniert
# Steht hier Yes kann der Wert im programm nicht geändert werden
$LoggedOnUser = $inifile.LogonSettings.LoggedOnUser

# Ist der Servername gesetzt muss dieser nicht immer angegeben werden, 
# ist hier ein Wert vergeben kann dieser im Programm nicht geändert werden 
$SetServername = $inifile.LogonSettings.SetServername

# Wenn Yes müssen beide obrigen werte gesetzt sein da es sonst zum fehler kommt. 
# Andernfalls wird die Anmeldung ausgeführt sobald das Programm startet 
$LogOnatStart = $inifile.LogonSettings.LogOnatStart 

#Anmeldung an Office 365, OU muss leer sein anedere werte werden ignoriert
$LogOntoO365 = $inifile.LogonSettings.O365 

#Display language - Anzeigesprache konfigurieren
$DisplayLang = $inifile.LogonSettings.DisplayLang

# Gibt an aus welcher OU die Benutzer eingelesen werden, wenn der Wert "0" ist werden alle Benutzer geladen (OU muss eindeutig sein!)
# nicht gültig für Office 365, da muss der wert leer sein
$OU = $inifile.LogonSettings.OrganizationalUnit

# Globale Variable für die versionsnummer
$VersionNumber ="2.3"
#########################################################################################################################################

#Ladebalkenfenster für ersten Start anzeigen (Benutzer werden geladen)
$Icon = [system.drawing.icon]::ExtractAssociatedIcon("$scriptRoot\icon.ico")
$Font = New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Regular)

$Fenster0 = New-Object System.Windows.Forms.Form
$Fenster0.Size = New-Object System.Drawing.Size(890,710)
$Fenster0.Height = 100
$Fenster0.Width = 500
$Fenster0.Text = "Out of Office Tool Version $VersionNumber $MsgLoadUser"
$Fenster0.StartPosition = "CenterScreen"
$Fenster0.BackColor = "#B2E4FF"
$Fenster0.Icon = $Icon
$Fenster0.Font = $Font
$Fenster0.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle

#########################################################################################################################################

# Hauptfensterfenster erstellen und konfigurieren

$Fenster1 = New-Object System.Windows.Forms.Form
$Fenster1.Size = New-Object System.Drawing.Size(890,710)
$Fenster1.MaximumSize = "890,710"
$Fenster1.MinimumSize = "890,710"
$Fenster1.Text = "Exchange Server / O365 Out of Office Tool Version $VersionNumber © Andres Sichel"
$Fenster1.StartPosition = "CenterScreen"
$Fenster1.BackColor = "#B2E4FF"
$Fenster1.Icon = $Icon
$Fenster1.Font = $Font
$Fenster1.WindowState = "Normal"
$Fenster1.add_SizeChanged({Minimieren})
$Fenster1.KeyPreview = $True #Enter drücken zulassen (nur für Benutzerabfragen)
$Fenster1.Add_KeyDown({if ($_.KeyCode -eq "Enter"){Abfrage}})
$Fenster1.Add_KeyDown({if ($_.KeyCode -eq "Escape"){Abmelden}}) # Mit Escape wird Abmelden ausgeführt und das Fenster wird geschlossen

#########################################################################################################################################
# Minimieren funktion einbauen

#Kontextmenü für Icon erstellen
$Kontext = New-Object System.Windows.Forms.ContextMenu
$KontextExit = New-Object System.Windows.Forms.MenuItem
$KontextInfo = New-Object System.Windows.Forms.MenuItem

$KontextExit.Text = "Exit"
$KontextExit.Index = 1
$KontextExit.add_Click({$NotifyIcon.Visible = $False,$Fenster1.close()})

$KontextInfo.Text = "Info"
$KontextInfo.Index = 2
$KontextInfo.add_Click({ShowInfo})

# NotifyIcon erstellen, erscheint wenn Fenster minimiert worden ist
$NotifyIcon = New-Object System.Windows.Forms.NotifyIcon
$NotifyIcon.Icon = $Icon
$NotifyIcon.Visible = $false
$NotifyIcon.add_Click({Iconclick})
$NotifyIcon.Text = "Out of Office Tool $VersionNumber"
$NotifyIcon.BalloonTipTitle = "Out of Office Tool $VersionNumber"
$NotifyIcon.BalloonTipIcon = "Info"
$NotifyIcon.ContextMenu = $Kontext
$NotifyIcon.ContextMenu.MenuItems.AddRange($KontextInfo)
$NotifyIcon.ContextMenu.MenuItems.AddRange($KontextExit)

function ShowInfo
{
$Fenster1.Hide()
$Fenster2 = New-Object System.Windows.Forms.Form
$Fenster2.Size = New-Object System.Drawing.Size(300,300)
$Fenster2.MaximumSize = "300,300"
$Fenster2.MinimumSize = "300,300"
$Fenster2.Text = "Out of Office Tool $VersionNumber Info"
$Fenster2.StartPosition = "CenterScreen"
$Fenster2.BackColor = "#B2E4FF"
$Fenster2.Icon = $Icon
$Fenster2.Font = $Font
$Fenster2.WindowState = "Normal"


#Text1
$Text3 = New-Object System.Windows.Forms.Label
$Text3.Location = New-Object System.Drawing.Size(20,20)
$Text3.Size = New-Object System.Drawing.Size(240,150)
$Text3.Text = "Out of Office Tool

Erstellt von Andres Sichel

blog@asichel.de

Version $VersionNumber"

$Text3.TextAlign = "MiddleCenter"
$Fenster2.Controls.Add($Text3)

$FontB = New-Object System.Drawing.Font("Arial",14,[System.Drawing.FontStyle]::Regular)
$CloseInfo = New-Object System.Windows.Forms.Button
$CloseInfo.Location = New-Object System.Drawing.Size(90,200)
$CloseInfo.Size = New-Object System.Drawing.Size(100,30)
$CloseInfo.Text = "OK"
$CloseInfo.BackColor = "#0064AF"
$CloseInfo.font = $FontB
$CloseInfo.foreColor = "White"
$CloseInfo.Add_Click({$Fenster2.Close()})
$Fenster2.Controls.Add($CloseInfo)

$Fenster2.Add_Shown({$Fenster2.Activate()})
[void] $Fenster2.ShowDialog()
}

function Minimieren
{
    if($Fenster1.WindowState -eq [Windows.Forms.FormWindowState]"Minimized")
    {
      $Fenster1.Hide()
      $NotifyIcon.Visible = $True      
      $NotifyIcon.ShowBalloonTip(5000)
    }       

}

function IconClick
{
    if($Fenster1.Visible)
    {
        $Fenster1.Hide()
    }
    else
    {        
        $Fenster1.Show()        
        $Fenster1.WindowState = [Windows.Forms.FormWindowState]"Normal"
    }
}

#########################################################################################################################################
#Globale Objekte erstellen

$progressBar1 = New-Object System.Windows.Forms.ProgressBar
$TextProgressBar = New-Object system.Windows.Forms.Label
$GroupBox1 = New-Object System.Windows.Forms.GroupBox
$GroupBox2 = New-Object System.Windows.Forms.GroupBox
$GroupBox3 = New-Object System.Windows.Forms.GroupBox
$GroupBox4 = New-Object System.Windows.Forms.GroupBox
$GroupBox5 = New-Object System.Windows.Forms.GroupBox
$GroupBox6 = New-Object System.Windows.Forms.GroupBox
$GroupBox7 = New-Object System.Windows.Forms.GroupBox
$UserName = New-Object System.Windows.Forms.Combobox
$Abfrage = New-Object System.Windows.Forms.Button
$ReportGrid = New-Object System.Windows.Forms.Button
$TextExtern = New-Object System.Windows.Forms.Label
$TextZeitStart = New-Object System.Windows.Forms.Label
$TextZeitEnde = New-Object System.Windows.Forms.Label
$About1 = New-Object System.Windows.Forms.Label
$ServerName = New-Object System.Windows.Forms.Textbox
$Login = New-Object System.Windows.Forms.Button
$WindowsUser = New-Object System.Windows.Forms.CheckBox
$Speichern = New-Object System.Windows.Forms.Button
$Abmelden = New-Object System.Windows.Forms.Button
$BenutzerLaden = New-Object System.Windows.Forms.Button
$HTMLForm1 = New-Object GvS.Controls.HtmlTextbox
$HTMLForm2 = New-Object GvS.Controls.HtmlTextbox
$RadioDisabled = New-Object System.Windows.Forms.RadioButton
$RadioDisabled.Checked = $true
$RadioEnabled = New-Object System.Windows.Forms.RadioButton
$Extern = New-Object System.Windows.Forms.CheckBox
$ExternKontakte = New-Object System.Windows.Forms.RadioButton
$ExternAlle = New-Object System.Windows.Forms.RadioButton
$Zeit = New-Object System.Windows.Forms.CheckBox
$ZeitStart = New-Object System.Windows.Forms.DateTimePicker
$DatumStart = New-Object System.Windows.Forms.DateTimePicker
$ZeitEnde = New-Object System.Windows.Forms.DateTimePicker
$DatumEnde = New-Object System.Windows.Forms.DateTimePicker
$Ausgabe1 = New-Object System.Windows.Forms.Label
$ExitBild = New-Object System.Windows.Forms.PictureBox
$StatusBild = New-Object System.Windows.Forms.PictureBox
$ToolTip = New-Object System.Windows.Forms.ToolTip
$Weiterleitung = New-Object System.Windows.Forms.CheckBox
$WeiterleitungMail = New-Object System.Windows.Forms.ComboBox
$WeiterleitungMailKopie = New-Object System.Windows.Forms.CheckBox
$TextWeiterleitung = New-Object System.Windows.Forms.Label

#########################################################################################################################################
#GUI text / Display language

if ($DisplayLang -eq "de")
{
<# statische texte #>
$NotifyIcon.BalloonTipText = "Out of Office Tool wurde minnimiert, Klick zum öffnen"
$GroupBox1.Text = "Zusatzoption bei aktiven Out of Office"
$GroupBox2.Text = "Verbindungs / Benutzeroptionen"
$GroupBox3.Text = "Out of Office Status abfragen / setzen"
$GroupBox4.Text = "Informationen"
$GroupBox5.Text = "Nachricht für Interne Empfänger"
$GroupBox6.Text = "Erweiterte Einstellungen"
$GroupBox7.Text = "Report"
$TextProgressBar.text="Programmstart wird vorbereitet!"
$TextExtern.Text = "Nachricht für Externe Empfänger:"
$TextZeitStart.Text = "Startzeit:"
$TextZeitEnde.Text = "Endzeit:"
$TextWeiterleitung.Text = "Keine Informationen über Mailweiterleitung abgefragt"
$TextProgressBar.Text = "Start wird vorbereitet"
$Login.Text = "Anmelden"
$WindowsUser.Text = "Anmeldung mit aktuellem Benutzer"
$Abfrage.Text = "Abfragen"
$ReportGrid.Text = "Liste aller Benutzer mit OoO"
$Speichern.Text = "Speichern / aktualisieren"
$BenutzerLaden.Text = "Benutzerliste erneut laden?"
$RadioDisabled.Text = "Out of Office deaktivieren / ist deaktiviert"
$RadioEnabled.Text = "Out of Office aktivieren / ist aktiviert"
$Extern.Text = "Auch an Absender außerhalb der Organisation senden?"
$ExternKontakte.Text = "Antworten nur an Absender in meiner Kontaktliste senden"
$ExternAlle.Text = "Antworten an alle externen Absender senden"
$Zeit.Text = "Antworten nur in diesem Zeitraum senden:"
$Weiterleitung.Text = "E-Mail weiterleitung an anderen Empfänger?"
$WeiterleitungMail.Text = "E-Mail zur Weiterleitung"
$WeiterleitungMailKopie.Text = "Nachrichten an Empfänger und alternativen Empfänger übermitteln?"

<# dynamische Texte in Funktionen #>
$MsgLogonFail = "Anmeldung nicht erfolgreich"
$MsgLogonSuccess ="Erfolgreich angemeldet, los gehts!`
                       Benutzerliste wird geladen!"
$MsgLoadUser = "Benutzer werden geladen"
$MsgNoUser = "Es wurde kein Benutzer eingetragen oder der Benutzer existiert nicht."
$MsgOOOdisabled = "Abwesenheitsassistent ist deaktiviert"
$MsgOOOEnabledAll = "Abwesenheitsassistent ist mit Nachrichten an alle Absender aktiv"
$MsgOOOEnabledContact = "Abwesenheitsassistent ist mit Nachrichten an bekannte Absender aktiv"
$MsgOOOEnabledInt = "Abwesenheitsassistent ist mit Nachrichten an interne Absender aktiv"
$MsgOOOEnabledTmAll = "Abwesenheitsassistent ist Zeitgesteuert mit Nachrichten an alle Absender aktiv"
$MsgOOOEnabledTmContact = "Abwesenheitsassistent ist Zeitgesteuert mit Nachrichten an bekannt Absender aktiv"
$MsgOOOEnabledTmInt = "Abwesenheitsassistent ist Zeitgesteuert mit Nachrichten an interne Absender aktiv"
$MsgFwDisabled =  "Keine E-Mail weiterleitung eingerichtet!"
$MsgFwEnableCopy = "Mails werden an angebene Adresse in Kopie weitergeleitet"
$MsgFwEnableNoCopy = "Mails werden an angebene Adresse weitergeleitet"
$MsgOOOSetDisable = "Die Abwesenheitsnotiz fuer den Benutzer $Mailbox1 wurde deaktiviert"
$MsgNoText = "Bitte einen Text hinterlegen"
$MsgOOOsetAll = "Die Abwesenheistnotiz fuer das Konto $Mailbox1 wurde gespeichert / aktualisiert. Es wird eine Externe Nachricht an ALLE Absender verschickt."
$MsgOOOSetContact = "Die Abwesenheistnotiz fuer das Konto $Mailbox1 wurde gespeichert / aktualisiert. Es wird eine Externe Nachricht nur an ALLE Absender aus der Kontaktliste verschickt"
$MsgOOOSetInt = "Die Abwesenheistnotiz fuer das Konto $Mailbox1 wurde gespeichert / aktualisiert. Es wird KEINE Externe Nachricht"
$MsgOOOSetTmAll = "Die Abwesenheistnotiz fuer das Konto $Mailbox1 wurde gespeichert / aktualisiert. Nachricht an alle externen Absender! Zusätzlich ist eine Zeitspanne konfiguriert! Von $StartZeit bis $EndZeit." 
$MsgOOOSetTmContact = "Die Abwesenheistnotiz fuer das Konto $Mailbox1 wurde gespeichert / aktualisiert. Nachricht nur an Bekannte Absender! Zusätzlich ist eine Zeitspanne konfiguriert! Von $StartZeit bis $EndZeit."
$MsgOOOSetTmInt = "Die Abwesenheistnotiz fuer das Konto $Mailbox1 wurde gespeichert / aktualisiert. Keine Nachricht an Externe Absender. Zusätzlich ist eine Zeitspanne konfiguriert! Von $StartZeit bis $EndZeit."
$MsgFwSet = "E-Mail weiterleitung eingerichtet!" 
$MsgFwSetCopy = "E-Mail Kopie-Weiterleitung eingerichtet!"  
$MsgFwSetNo = "Keine E-Mail weiterleitung eingerichtet!" 
$MsgLogoff = "Erfolgreich abgemeldet"
$MsUserLoadSuccess = "Benutzerliste wurde erneut geladen"
$MsgAutoLogonDisabeld = "Automatische Anmeldung wurde deaktiviert"
$MsgAutoLogonFailedText = "Bitte Variablen in Settings.ini setzen"
$MsgAutoLogonFailedTitle = "Automatischer Login fehlgeschlagen"
$MsgTextGridTitle = "Alle Benutzer mit Out of Office"


}
if ($DisplayLang -eq "en")
{
<# statische texte #>
$NotifyIcon.BalloonTipText = "Out of Office Tool was minimized, click to open"
$GroupBox1.Text = "Automatic replies settings"
$GroupBox2.Text = "Connection settings"
$GroupBox3.Text = "Out of Office enable / enabled"
$GroupBox4.Text = "Additional informations"
$GroupBox5.Text = "Message for internal reciepients"
$GroupBox6.Text = "Advanced settings"
$GroupBox7.Text = "Report"
$TextProgressBar.text="Programm starts..."
$TextExtern.Text = "Message for external recipients"
$TextZeitStart.Text = "Start time"
$TextZeitEnde.Text = "End time"
$TextWeiterleitung.Text = "No informations about mail forwarding"
$TextProgressBar.Text = "Start in progress"
$Login.Text = "Log on"
$WindowsUser.Text = "Log on with Windows User"
$Abfrage.Text = "Query"
$ReportGrid.Text = "Report all users with OoO"
$Speichern.Text = "Save / update"
$BenutzerLaden.Text = "Rerun User query"
$RadioDisabled.Text = "Out of Office disabled"
$RadioEnabled.Text = "Out of Office enabled"
$Extern.Text = "Send message to external reciepients?"
$ExternKontakte.Text = "Send replies only to senders in my contacts list"
$ExternAlle.Text = "Send replies to all senders"
$Zeit.Text = "Send replies only during this time period"
$Weiterleitung.Text = "Forward mails to another recipient?"
$WeiterleitungMail.Text = "Mail forwarding"
$WeiterleitungMailKopie.Text = "Deliver message to both forwarding address and mailbox"

<# dynamische Texte in Funktionen #>
$MsgLogonFail = "Logon not successfull"
$MsgLogonSuccess ="Logon successfull, lets go!`
                       Userlist loaded"
$MsgLoadUser = "Userlist loading"
$MsgNoUser = "No username set or user does not exist"
$MsgOOOdisabled = "Out of Office is disabled"
$MsgOOOEnabledAll = "Out of Office for all external senders enabled"
$MsgOOOEnabledContact = "Out of Office for all senders in contact list enabled"
$MsgOOOEnabledInt = "Out of Office only for internal senders enabled"
$MsgOOOEnabledTmAll = "Out of Office for all external senders with time peroid enabled"
$MsgOOOEnabledTmContact = "Out of Office for all senders in contact list with time period enabled"
$MsgOOOEnabledTmInt = "Out of Office only for internal senders with time period enabled"
$MsgFwDisabled =  "No mailforwarding settings"
$MsgFwEnableCopy = "Copy forwarding to adress enabled"
$MsgFwEnableNoCopy = "Forwarding to address enabled"
$MsgOOOSetDisable = "Out of Office for user $Mailbox1 disabled"
$MsgNoText = "Please insert Out of Office Message"
$MsgOOOsetAll = "Out of Office settings for $Mailbox1 saved / refreshed. Replay send to all external senders."
$MsgOOOSetContact = "Out of Office settings for $Mailbox1 saved / refreshed. Reply send to all senders in contact list"
$MsgOOOSetInt = "Out of Office settings for $Mailbox1 saved / refreshed. Reply send only to internal senders"
$MsgOOOSetTmAll = "Out of Office settings for $Mailbox1 saved / refreshed. Reply send to all external senders! Time period configured from $StartZeit till $EndZeit." 
$MsgOOOSetTmContact = "Out of Office settings for $Mailbox1 saved / refreshed. Reply send to all senders in contact list! Time period configured from $StartZeit till $EndZeit."
$MsgOOOSetTmInt = "Out of Office settings for $Mailbox1 saved / refreshed. Reply send only to internal senders. Time period configured from $StartZeit till $EndZeit."
$MsgFwSet = "E-Mail forwarding configured" 
$MsgFwSetCopy = "E-Mail copy forwarding configured"  
$MsgFwSetNo = "No E-Mail forwarding configured" 
$MsgLogoff = "Logoff successfull"
$MsUserLoadSuccess = "user list refreshed succsessfull"
$MsgAutoLogonDisabeld = "Automatic logon disabeld"
$MsgAutoLogonFailedText = "Please set variables in settings.ini file"
$MsgAutoLogonFailedTitle = "Automatic logon failure"
$MsgTextGridTitle = "All Users with Out of Office"

}

#########################################################################################################################################
#Gruppenboxen: Namen vergeben und Layout erstellen

# Gruppenbox für Zusatzoptionen

$GroupBox1.Location = New-Object System.Drawing.Size(400,100)
$GroupBox1.size = New-Object System.Drawing.Size(465,190) 
$GroupBox1.Font = $Font
$Fenster1.Controls.Add($GroupBox1)

# Gruppenbox fuer Server Benutzer und usw

$GroupBox2.Location = New-Object System.Drawing.Size(5,50)
$GroupBox2.size = New-Object System.Drawing.Size(380,190) 
$GroupBox2.Font = $Font
$Fenster1.Controls.Add($GroupBox2)

# Gruppenbox fuer Status Out of Office

$GroupBox3.Location = New-Object System.Drawing.Size(400,10)
$GroupBox3.size = New-Object System.Drawing.Size(370,80) 
$GroupBox3.Font = $Font
$Fenster1.Controls.Add($GroupBox3)

# Gruppenbox Ergebnisausgabe

$GroupBox4.Location = New-Object System.Drawing.Size(5,480)
$GroupBox4.size = New-Object System.Drawing.Size(380,180) 
$GroupBox4.Font = $Font
$Fenster1.Controls.Add($GroupBox4)

# Gruppenbox HTML Felder

$GroupBox5.Location = New-Object System.Drawing.Size(400,300)
$GroupBox5.size = New-Object System.Drawing.Size(465,360) 
$GroupBox5.Font = $Font
$Fenster1.Controls.Add($GroupBox5)

# Gruppenbox Erweiterte Optionen

$GroupBox6.Location = New-Object System.Drawing.Size(5,300)
$GroupBox6.size = New-Object System.Drawing.Size(380,170) 
$GroupBox6.Font = $Font
$Fenster1.Controls.Add($GroupBox6)

#Gruppenbox für Report

$GroupBox7.Location = New-Object System.Drawing.Size(5,245)
$GroupBox7.size = New-Object System.Drawing.Size(380,50) 
$GroupBox7.Font = $Font
$Fenster1.Controls.Add($GroupBox7)

#########################################################################################################################################
#Form für Ladebalken erstellen

$progressBar1 = New-Object System.Windows.Forms.ProgressBar
$progressBar1.Name = 'progressBar1'
$progressBar1.Value = 0
$progressBar1.Style="Continuous"

$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 500 - 40
$System_Drawing_Size.Height = 20
$progressBar1.Size = $System_Drawing_Size
$progressBar1.Left = 5
$progressBar1.Top = 40
$Fenster0.Controls.Add($progressBar1)

$Fenster0.Show()| out-null
$Fenster0.Refresh()

#########################################################################################################################################
# Funktionen erstellen

function Anmelden  
            {
            if ($LogOntoO365 -eq "yes")
            {
            #Anmeldung an Office 365 iniitieren
            
            $Fenster1.Cursor = [System.Windows.Forms.Cursors]::WaitCursor       
                     #Am Exchange Server anmelden
                     #Zertifikatsprüfungen abschalten
                     $SkipCertificate = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck -Verbose

                     #Servernamen auslesen
                     $Server = "outlook.office365.com"
                     $ServerName.Text = "outlook.office365.com"
                              
                        #Anmeldedaten erfragen
                        $cred = Get-Credential

                        #PS-Session starten --> Bei connection URI den korrekten Exchange 2010 oder 2013 Servernamen eingeben, vorzugsweise CAS-Server
                        $Session = New-PSSession `
                                         -ConfigurationName Microsoft.Exchange `
                                         -ConnectionUri https://$Server/PowerShell-liveid/ `
                                         -Authentication basic `
                                         -SessionOption $SkipCertificate  `
                                         -credential $cred `
                                         -AllowRedirection `
                                         -Name Exchange
                                  

                        #Importieren der erzeugten Session in die aktuelle Session, lokal gleichlautende Befehle werden ersetzen
                        Import-PSSession -Session $Session -AllowClobber -Verbose
                        $Fenster1.Cursor = [System.Windows.Forms.Cursors]::Arrow
                        if (!$Session.State -eq "open")
                            {
                            $Ausgabe1.text = $MsgLogonFail
                            }
                        else
                            {
                            $Ausgabe1.text = $MsgLogonSuccess

                            $Fenster0.Show()
                            if ($OU -notlike $null)
                            {
                            $UserArray = Get-User -OrganizationalUnit $OU -RecipientTypeDetails UserMailbox -SortBy Name -ResultSize Unlimited
                            }
                            else
                            {
                            $UserArray = Get-User -RecipientTypeDetails UserMailbox -SortBy Name -ResultSize Unlimited
                            }
                            $i = 0
                            foreach($user in $UserArray)
                            {
                            $i++
                            [int]$pct = ($i/$UserArray.count)*100
                            $progressbar1.Value = $pct
                            $TextProgressBar.text="$MsgLoadUser : $($user.name)"
                            $Fenster0.Refresh()
                            $UserName.items.add($user.Name)
                            $WeiterleitungMail.items.add($user.Name)
                            }
                            $BenutzerLaden.Enabled = $true
                            }
                            $Fenster0.Hide()
                            $UserName.Enabled = $true
                            $RadioDisabled.Enabled = $true
                            $RadioEnabled.Enabled = $true
                            $Weiterleitung.Enabled = $true
                            $ReportGrid.Enabled = $true 
            }
            else
            {
            if (!$WindowsUser.Checked) #Wenn anmeldung nicht automatsich erfolgt
                {
                     $Fenster1.Cursor = [System.Windows.Forms.Cursors]::WaitCursor       
                     #Am Exchange Server anmelden
                     #Zertifikatsprüfungen abschalten
                     $SkipCertificate = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck -Verbose

                     #Servernamen auslesen
                     $Server = $ServerName.Text;
                              
                        #Anmeldedaten erfragen
                        $cred = Get-Credential

                        #PS-Session starten --> Bei connection URI den korrekten Exchange 2010 oder 2013 Servernamen eingeben, vorzugsweise CAS-Server
                        $Session = New-PSSession `
                                         -ConfigurationName Microsoft.Exchange `
                                         -ConnectionUri http://$Server/PowerShell/ `
                                         -Authentication Kerberos `
                                         -SessionOption $SkipCertificate  `
                                         -credential $cred `
                                         -Name Exchange
                                  

                        #Importieren der erzeugten Session in die aktuelle Session, lokal gleichlautende Befehle werden ersetzen
                        Import-PSSession -Session $Session -AllowClobber -Verbose
                        $Fenster1.Cursor = [System.Windows.Forms.Cursors]::Arrow
                        if (!$Session.State -eq "open")
                            {
                            $Ausgabe1.text = $MsgLogonFail
                            }
                        else
                            {
                            $Ausgabe1.text =$MsgLogonSuccess
                            $Fenster0.Show()
                            if ($OU -notlike $null)
                            {
                            $UserArray = Get-User -OrganizationalUnit $OU -RecipientTypeDetails UserMailbox -SortBy Name -ResultSize Unlimited
                            }
                            else
                            {
                            $UserArray = Get-User -RecipientTypeDetails UserMailbox -SortBy Name -ResultSize Unlimited
                            }
                            $i = 0
                            foreach($user in $UserArray)
                            {
                            $i++
                            [int]$pct = ($i/$UserArray.count)*100
                            $progressbar1.Value = $pct
                            $TextProgressBar.text="$MsgLoadUser : $($user.name)"
                            $Fenster0.Refresh()
                            $UserName.items.add($user.Name)
                            $WeiterleitungMail.items.add($user.Name)
                            }
                            $BenutzerLaden.Enabled = $true
                            }
                            $Fenster0.Hide()
                            $UserName.Enabled = $true
                            $RadioDisabled.Enabled = $true
                            $RadioEnabled.Enabled = $true
                            $Weiterleitung.Enabled = $true
                            $ReportGrid.Enabled = $true
                 }
                 else
                 {
                    $Fenster1.Cursor = [System.Windows.Forms.Cursors]::WaitCursor       
                     #Am Exchange Server anmelden
                     #Zertifikatsprüfungen abschalten
                     $SkipCertificate = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck -Verbose

                     #Servernamen auslesen
                     $Server = $ServerName.Text;
                              
                        #Anmeldedaten erfragen
                        #PS-Session starten --> Bei connection URI den korrekten Exchange 2010 oder 2013 Servernamen eingeben, vorzugsweise CAS-Server
                        $Session = New-PSSession `
                                         -ConfigurationName Microsoft.Exchange `
                                         -ConnectionUri http://$Server/PowerShell/ `
                                         -Authentication Kerberos `
                                         -SessionOption $SkipCertificate  `
                                         -Name Exchange
                                  

                        #Importieren der erzeugten Session in die aktuelle Session, lokal gleichlautende Befehle werden ersetzen
                        Import-PSSession -Session $Session -AllowClobber -Verbose
                        $Fenster1.Cursor = [System.Windows.Forms.Cursors]::Arrow
                        if (!$Session.State -eq "open")
                            {
                            $Ausgabe1.text = $MsgLogonFail
                            }
                        else
                          {
                           $Ausgabe1.text = $MsgLogonSuccess
                           $Fenster0.Show() 
                          if ($OU -notlike $null)
                            {
                            $UserArray = Get-User -OrganizationalUnit $OU -RecipientTypeDetails UserMailbox -SortBy Name -ResultSize Unlimited
                            }
                            else
                            {
                            $UserArray = Get-User -RecipientTypeDetails UserMailbox -SortBy Name -ResultSize Unlimited
                            }
                            $i = 0
                            foreach($user in $UserArray)
                            {
                            $i++
                            [int]$pct = ($i/$UserArray.count)*100
                            $progressbar1.Value = $pct
                            $TextProgressBar.text="$MsgLoadUser : $($user.name)"
                            $Fenster0.Refresh()
                            $UserName.items.add($user.Name)
                            $WeiterleitungMail.items.add($user.Name)
                            }
                            $BenutzerLaden.Enabled = $true
                            }
                            $Fenster0.Hide()
                            $UserName.Enabled = $true
                            $RadioDisabled.Enabled = $true
                            $RadioEnabled.Enabled = $true
                            $Weiterleitung.Enabled = $true
                            $ReportGrid.Enabled = $true
                 }           
                 }
                 }#Beende Funktion Anmelden
                  
function Abfrage
    {
            #Felder leeren
            $HTMLForm1.Text = $null
            $HTMLForm2.Text = $null
            $WeiterleitungMail.Text = $null
            $Ausgabe1.Text = $null
            
            $Mailbox = $username.text
                 if (($Username.text.Length -eq 0) -or (!(Get-Recipient $Username.text -ErrorAction SilentlyContinue))) 
                        { 
                        $Ausgabe1.text = $MsgNoUser
                        }
                else {                
                        $Info = Get-MailboxAutoReplyConfiguration -Identity $Mailbox | select AutoReplyState,StartTime,Endtime,ExternalAudience,ExternalMessage,InternalMessage
                        $info2 = Get-Mailbox $Mailbox | select DeliverToMailboxAndForward,ForwardingAddress,ForwardingSmtpAddress
                        #$Ausgabe1.text =  $Info | Out-String
                        
                        <# Hier werden nun die verschieden Stati abgefrgat, 
                         Buttons und Boxen aktiviert / deaktiviert
                         und die Entsprechenden Daten vom Postfach hineingeschrieben #>
                        if (($Info.AutoReplyState -match "Disabled") -and ($Info.ExternalAudience -match "All") )
                        {
                          $Radiodisabled.Checked = $true
                          $RadioDisabled.Enabled = $true
                          $RadioEnabled.Enabled = $true
                          $Extern.Checked = $true
                          $Extern.Enabled = $false
                          $ExternKontakte.Enabled = $false
                          $ExternAlle.Enabled = $false
                          $ExternAlle.Checked = $true
                          $Zeit.Enabled = $false
                          $Zeit.Checked = $false
                          $ZeitStart.Enabled = $false
                          $DatumStart.Enabled = $false
                          $ZeitEnde.Enabled = $false
                          $DatumEnde.Enabled = $false
                          $HTMLForm1.Text = $Info.InternalMessage
                          $HTMLForm2.Text = $Info.ExternalMessage
                          $Ausgabe1.Text = $MsgOOOdisabled
                          $StatusBild.ImageLocation = "$scriptRoot\inaktiv.png"
                          $HTMLForm1.Enabled = $false
                          $HTMLForm2.Enabled = $false
                                                    
                        }
                        if (($Info.AutoReplyState -match "Disabled") -and ($Info.ExternalAudience -match "Known"))
                        {
                          $Radiodisabled.Checked = $true
                          $RadioDisabled.Enabled = $true
                          $RadioEnabled.Enabled = $true
                          $Extern.Checked = $true
                          $Extern.Enabled = $false
                          $ExternKontakte.Enabled = $false
                          $ExternAlle.Enabled = $false
                          $ExternKontakte.Checked = $true
                          $Zeit.Enabled = $false
                          $Zeit.Checked = $false
                          $ZeitStart.Enabled = $false
                          $DatumStart.Enabled = $false
                          $ZeitEnde.Enabled = $false
                          $DatumEnde.Enabled = $false
                          $HTMLForm1.Text = $Info.InternalMessage
                          $HTMLForm2.Text = $Info.ExternalMessage
                          $Ausgabe1.Text = $MsgOOOdisabled
                          $StatusBild.ImageLocation = "$scriptRoot\inaktiv.png"
                          $HTMLForm1.Enabled = $false
                          $HTMLForm2.Enabled = $false  
                                                
                        }
                        if (($Info.AutoReplyState -match "Disabled") -and ($Info.ExternalAudience -match "None"))
                        {
                          $Radiodisabled.Checked = $true
                          $RadioDisabled.Enabled = $true
                          $RadioEnabled.Enabled = $true
                          $Extern.Checked = $false
                          $Extern.Enabled = $false
                          $ExternKontakte.Enabled = $false
                          $ExternAlle.Enabled = $false
                          $ExternKontakte.Checked = $false
                          $Zeit.Enabled = $false
                          $Zeit.Checked = $false
                          $ZeitStart.Enabled = $false
                          $DatumStart.Enabled = $false
                          $ZeitEnde.Enabled = $false
                          $DatumEnde.Enabled = $false
                          $HTMLForm1.Text = $Info.InternalMessage
                          $HTMLForm2.Text = $Info.ExternalMessage
                          $Ausgabe1.Text = $MsgOOOdisabled
                          $StatusBild.ImageLocation = "$scriptRoot\inaktiv.png"
                          $HTMLForm1.Enabled = $false
                          $HTMLForm2.Enabled = $false
                                                  
                        }
                        if (($Info.AutoReplyState -match "Enabled") -and ($Info.ExternalAudience -match "All") )
                        {
                            $RadioEnabled.Checked = $true
                            $RadioEnabled.Enabled = $true
                            $Radiodisabled.Enabled = $true
                            $HTMLForm1.Text = $Info.InternalMessage
                            $HTMLForm2.Text = $Info.ExternalMessage
                            $Extern.Checked = $true
                            $Extern.Enabled = $true
                            $ExternAlle.Enabled = $true
                            $ExternAlle.Checked = $true
                            $ExternKontakte.Enabled =$true
                            $Zeit.Checked = $false
                            $Zeit.Enabled = $true
                            $Ausgabe1.Text = $MsgOOOEnabledAll
                            $StatusBild.ImageLocation = "$scriptRoot\aktiv.png"
                            $HTMLForm1.Enabled = $true
                            $HTMLForm2.Enabled = $true
                            
                        }
                        if (($Info.AutoReplyState -match "Enabled") -and ($Info.ExternalAudience -match "Known"))
                        {
                            $RadioEnabled.Checked = $true
                            $RadioEnabled.Enabled = $true
                            $Radiodisabled.Enabled = $true
                            $HTMLForm1.Text = $Info.InternalMessage
                            $HTMLForm2.Text = $Info.ExternalMessage
                            $Extern.Checked = $true
                            $Extern.Enabled = $true
                            $ExternAlle.Enabled = $true
                            $ExternKontakte.Enabled =$true
                            $ExternKontakte.Checked = $true
                            $Zeit.Checked = $false
                            $Zeit.Enabled = $true
                            $Ausgabe1.Text = $MsgOOOEnabledContact
                            $StatusBild.ImageLocation = "$scriptRoot\aktiv.png"
                            $HTMLForm1.Enabled = $true
                            $HTMLForm2.Enabled = $true
                            
                        }
                        if (($Info.AutoReplyState -match "Enabled") -and ($Info.ExternalAudience -match "None"))
                        {
                            $RadioEnabled.Checked = $true
                            $RadioEnabled.Enabled = $true
                            $Radiodisabled.Enabled = $true
                            $HTMLForm1.Text = $Info.InternalMessage
                            $HTMLForm2.Text = $Info.ExternalMessage
                            $Extern.Checked = $false
                            $Extern.Enabled = $true
                            $ExternAlle.Checked = $false
                            $ExternKontakte.Checked = $false                         
                            $Zeit.Checked = $false
                            $Zeit.Enabled = $true
                            $Ausgabe1.Text = $MsgOOOEnabledInt
                            $StatusBild.ImageLocation = "$scriptRoot\aktiv.png"
                            $HTMLForm1.Enabled = $true
                            $HTMLForm2.Enabled = $true
                        }
                        if (($info.AutoReplyState -match "Scheduled") -and ($Info.ExternalAudience -match "All"))
                        {
                            $RadioEnabled.Checked = $true
                            $RadioEnabled.Enabled = $true
                            $Radiodisabled.Enabled = $true
                            $HTMLForm1.Text = $Info.InternalMessage
                            $HTMLForm2.Text = $Info.ExternalMessage
                            $Extern.Checked = $true
                            $Extern.Enabled = $true
                            $ExternAlle.Enabled = $true
                            $ExternAlle.Checked = $true
                            $ExternKontakte.Enabled =$true
                            $Zeit.Checked = $true
                            $Zeit.Enabled = $true
                            $ZeitStart.Text = [System.Timespan]::Parse($Info.StartTime.ToString("HH:mm"))
                            $ZeitEnde.Text = [System.Timespan]::Parse($Info.EndTime.ToString("HH:mm"))
                            $DatumStart.Value = $Info.StartTime
                            $DatumEnde.Value = $Info.EndTime
                            $Ausgabe1.Text = $MsgOOOEnabledTmAll
                            $StatusBild.ImageLocation = "$scriptRoot\aktiv.png"
                            $HTMLForm1.Enabled = $true
                            $HTMLForm2.Enabled = $true
                            
                        }
                        if (($Info.AutoReplyState -match "Scheduled") -and ($Info.ExternalAudience -match "Known"))
                        {
                            $RadioEnabled.Checked = $true
                            $RadioEnabled.Enabled = $true
                            $Radiodisabled.Enabled = $true
                            $HTMLForm1.Text = $Info.InternalMessage
                            $HTMLForm2.Text = $Info.ExternalMessage
                            $Extern.Checked = $true
                            $Extern.Enabled = $true
                            $ExternAlle.Enabled = $true
                            $ExternKontakte.Enabled =$true
                            $ExternKontakte.Checked = $true
                            $Zeit.Checked = $true
                            $Zeit.Enabled = $true
                            $ZeitStart.Text = [System.Timespan]::Parse($Info.StartTime.ToString("HH:mm"))
                            $ZeitEnde.Text = [System.Timespan]::Parse($Info.EndTime.ToString("HH:mm"))
                            $DatumStart.Value = $Info.StartTime
                            $DatumEnde.Value = $Info.EndTime
                            $Ausgabe1.Text = $MsgOOOEnabledTmContact
                            $StatusBild.ImageLocation = "$scriptRoot\aktiv.png"
                            $HTMLForm1.Enabled = $true
                            $HTMLForm2.Enabled = $true
                        }
                        if (($Info.AutoReplyState -match "Scheduled") -and ($Info.ExternalAudience -match "None"))
                        {
                            $RadioEnabled.Checked = $true
                            $RadioEnabled.Enabled = $true
                            $Radiodisabled.Enabled = $true
                            $HTMLForm1.Text = $Info.InternalMessage
                            $HTMLForm2.Text = $Info.ExternalMessage
                            $Extern.Checked = $false
                            $Extern.Enabled = $true
                            $ExternAlle.Checked = $false
                            $ExternKontakte.Checked = $false                         
                            $Zeit.Checked = $True
                            $Zeit.Enabled = $true
                            $ZeitStart.Text = [System.Timespan]::Parse($Info.StartTime.ToString("HH:mm"))
                            $ZeitEnde.Text = [System.Timespan]::Parse($Info.EndTime.ToString("HH:mm"))
                            $DatumStart.Value = $Info.StartTime
                            $DatumEnde.Value = $Info.EndTime
                            $Ausgabe1.Text = $MsgOOOEnabledTmInt
                            $StatusBild.ImageLocation = "$scriptRoot\aktiv.png"
                            $HTMLForm1.Enabled = $true
                            $HTMLForm2.Enabled = $true
                            
                        }
                        if ($info2.ForwardingAddress -eq $null)
                        {
                            $WeiterleitungMail.Enabled = $false
                            $Weiterleitung.Checked = $false
                            $Weiterleitung.Enabled = $true
                            $WeiterleitungMailKopie.Enabled = $false
                            $WeiterleitungMailKopie.Checked = $false
                            $TextWeiterleitung.Text = $MsgFwDisabled 
                        }
                                               }
                          if ( ($info2.ForwardingAddress -ne $null) -and ($info2.DeliverToMailboxAndForward -eq $true)) 
                        {
                            $info3 = get-mailbox $info2.ForwardingAddress
                            $Weiterleitung.Enabled = $true
                            $Weiterleitung.Checked = $true
                            $WeiterleitungMail.Enabled = $true
                            $WeiterleitungMail.Text = $info3
                            $WeiterleitungMailKopie.Checked = $true
                            $TextWeiterleitung.Text = $MsgFwEnableNoCopy 
                        }        

                        if ( ($info2.ForwardingAddress -ne $null) -and ($info2.DeliverToMailboxAndForward -eq $false)) 
                        {
                            $info3 = get-mailbox $info2.ForwardingAddress
                            $Weiterleitung.Enabled = $true
                            $Weiterleitung.Checked = $true
                            $WeiterleitungMail.Enabled = $true
                            $WeiterleitungMail.Text = $info3
                            $WeiterleitungMailKopie.Checked = $false
                            $TextWeiterleitung.Text = $MsgFwEnableNoCopy 
                        }        
                      } # Beendet die Funktion Abfrage
    
function Speichern {
                
            $Mailbox1 = $username.text
                 if (($Username.text.Length -eq 0) -or (!(Get-Recipient $Username.text -ErrorAction SilentlyContinue))) 
                        { 
                        $Ausgabe1.text = $MsgNoUser
                        }
                else {  
                        <# Hir werden jetzt alle Variationen abgerfgat die eingetragen werden können,
                           mit Zeit, ohne Externe Absender usw. Jeweils auch die Abfrage ob Text hinterlegt worden ist #>
                        if ($RadioDisabled.Checked)
                        {
                            Set-MailboxAutoReplyConfiguration $Mailbox1 -AutoReplyState disabled
                            $Ausgabe1.Text = $MsgOOOSetDisable
                            
                        }
                        if (($RadioEnabled.Checked) -and ($Extern.checked) -and ($ExternAlle.Checked) -and (!$ExternKontakte.Checked) -and (!$Zeit.Checked)) #Aktiv, alle externen
                        {             
                            if (($HTMLForm1.Text.Length -eq 0) -or ($HTMLForm2.Text.Length -eq 0))
                            {
                                $Ausgabe1.Text = $MsgNoText
                            }
                            else
                            {
                            $Intern = $HTMLForm1.Text
                            $Extern = $HTMLForm2.Text
                            Set-MailboxAutoReplyConfiguration $Mailbox1 -AutoReplyState enabled -InternalMessage "$intern" -ExternalMessage "$extern" -ExternalAudience All
                            $Ausgabe1.text = $MsgOOOsetAll
                            
                            }
                        }
                        if (($RadioEnabled.Checked) -and ($Extern.checked) -and (!$ExternAlle.Checked) -and ($ExternKontakte.Checked) -and (!$Zeit.Checked)) #aktiv, extern nur kontakte
                        {             
                            if (($HTMLForm1.Text.Length -eq 0) -or ($HTMLForm2.Text.Length -eq 0))
                            {
                                $Ausgabe1.Text = $MsgNoText
                            }
                            else
                            {
                            $Intern = $HTMLForm1.Text
                            $Extern = $HTMLForm2.Text
                            Set-MailboxAutoReplyConfiguration $Mailbox1 -AutoReplyState enabled -InternalMessage "$intern" -ExternalMessage "$extern" -ExternalAudience Known
                            $Ausgabe1.text = $MsgOOOSetContact
                            
                            }
                        }
                        if (($RadioEnabled.Checked) -and (!$Extern.checked) -and (!$ExternAlle.Checked) -and (!$ExternKontakte.Checked) -and (!$Zeit.Checked)) #aktiv, keine Externen nachrichten keine Zeitspanne
                        {             
                            if (($HTMLForm1.Text.Length -eq 0) -or ($HTMLForm2.Text.Length -eq 0))
                            {
                                $Ausgabe1.Text = $MsgSet
                            }
                            else
                            {
                            $Intern = $HTMLForm1.Text
                            $Extern = $HTMLForm2.Text
                            Set-MailboxAutoReplyConfiguration $Mailbox1 -AutoReplyState enabled -InternalMessage "$intern" -ExternalMessage "$extern" -ExternalAudience none
                            $Ausgabe1.text = $MsgOOOSetInt
                            
                            }
                        }
                        if (($RadioEnabled.Checked) -and ($Extern.checked) -and ($ExternAlle.Checked) -and (!$ExternKontakte.Checked) -and ($Zeit.Checked)) #aktiv, extern alle und Zeitspanne
                         {             
                            if (($HTMLForm1.Text.Length -eq 0) -or ($HTMLForm2.Text.Length -eq 0))
                            {
                                $Ausgabe1.Text = $MsgNoText
                            }
                            else
                            {
                            $Gebiet = New-Object system.globalization.cultureinfo 'en-us'
                            $Datum_Start = [System.DateTime]::Parse($DatumStart.Value.ToString("MM.dd.yyyy"),$Gebiet)
                            $Zeit_Start = [System.Timespan]::Parse($ZeitStart.Value.ToString("HH:mm"))
                            $StartZeit = $Datum_Start.Add($Zeit_Start)
                            $StartZeit = Get-Date $StartZeit -format 'MM.dd.yyyy HH:mm'            
                            
                            $Datum_Ende = [System.DateTime]::Parse($DatumEnde.Value.ToString("MM.dd.yyyy"),$Gebiet)
                            $Zeit_Ende = [System.Timespan]::Parse($ZeitEnde.Value.ToString("HH:mm"))
                            $EndZeit = $Datum_Ende.Add($Zeit_Ende)
                            $EndZeit = Get-Date $EndZeit -format 'MM.dd.yyyy HH:mm'
                            
                            $Intern = $HTMLForm1.Text
                            $Extern = $HTMLForm2.Text
                            Set-MailboxAutoReplyConfiguration $Mailbox1 -AutoReplyState scheduled -InternalMessage "$intern" -ExternalMessage "$extern" -StartTime $StartZeit -EndTime $EndZeit -ExternalAudience all
                            $Ausgabe1.text = $MsgOOOSetTmAll
                            
                            }
                        }
                        if (($RadioEnabled.Checked) -and ($Extern.checked) -and (!$ExternAlle.Checked) -and ($ExternKontakte.Checked) -and ($Zeit.Checked)) #aktiv, extern nur kontakte und Zeitspanne
                         {             
                            if (($HTMLForm1.Text.Length -eq 0) -or ($HTMLForm2.Text.Length -eq 0))
                            {
                                $Ausgabe1.Text = $MsgNoText
                            }
                            else
                            {
                            $Gebiet = New-Object system.globalization.cultureinfo 'en-us'
                            $Datum_Start = [System.DateTime]::Parse($DatumStart.Value.ToString("MM.dd.yyyy"),$Gebiet)
                            $Zeit_Start = [System.Timespan]::Parse($ZeitStart.Value.ToString("HH:mm"))
                            $StartZeit = $Datum_Start.Add($Zeit_Start)
                            $StartZeit = Get-Date $StartZeit -format 'MM.dd.yyyy HH:mm'
                            
       
                            $Datum_Ende = [System.DateTime]::Parse($DatumEnde.Value.ToString("MM.dd.yyyy"),$Gebiet)
                            $Zeit_Ende = [System.Timespan]::Parse($ZeitEnde.Value.ToString("HH:mm"))
                            $EndZeit = $Datum_Ende.Add($Zeit_Ende)
                            $EndZeit = Get-Date $EndZeit -format 'MM.dd.yyyy HH:mm'
                            
                            $Intern = $HTMLForm1.Text
                            $Extern = $HTMLForm2.Text
                            Set-MailboxAutoReplyConfiguration $Mailbox1 -AutoReplyState scheduled -InternalMessage "$intern" -ExternalMessage "$extern" -StartTime $StartZeit -EndTime $EndZeit -ExternalAudience Known
                            $Ausgabe1.text = $MsgOOOSetTmContact
                            
                            }
                        }
                        if (($RadioEnabled.Checked) -and (!$Extern.checked) -and (!$ExternAlle.Checked) -and (!$ExternKontakte.Checked) -and ($Zeit.Checked)) #aktiv, keine Externen und Zeitspanne
                         {             
                            if (($HTMLForm1.Text.Length -eq 0) -or ($HTMLForm2.Text.Length -eq 0))
                            {
                                $Ausgabe1.Text = $MsgNoText
                            }
                            else
                            {
                            $Gebiet = New-Object system.globalization.cultureinfo 'en-us'
                            $Datum_Start = [System.DateTime]::Parse($DatumStart.Value.ToString("MM.dd.yyyy"),$Gebiet)
                            $Zeit_Start = [System.Timespan]::Parse($ZeitStart.Value.ToString("HH:mm"))
                            $StartZeit = $Datum_Start.Add($Zeit_Start)
                            $StartZeit = Get-Date $StartZeit -format 'MM.dd.yyyy HH:mm'
                            
       
                            $Datum_Ende = [System.DateTime]::Parse($DatumEnde.Value.ToString("MM.dd.yyyy"),$Gebiet)
                            $Zeit_Ende = [System.Timespan]::Parse($ZeitEnde.Value.ToString("HH:mm"))
                            $EndZeit = $Datum_Ende.Add($Zeit_Ende)
                            $EndZeit = Get-Date $EndZeit -format 'MM.dd.yyyy HH:mm'
                            
                            $Intern = $HTMLForm1.Text
                            $Extern = $HTMLForm2.Text
                            Set-MailboxAutoReplyConfiguration $Mailbox1 -AutoReplyState scheduled -InternalMessage "$intern" -ExternalMessage "$extern" -StartTime $StartZeit -EndTime $EndZeit -ExternalAudience none
                            $Ausgabe1.text = $MsgOOOSetTmInt
                            
                            }
                        }
                        if (($Weiterleitung.Checked) -and (!$WeiterleitungMailKopie.Checked))
                        { 
                            $ForwardAddress = $WeiterleitungMail.Text
                            Set-Mailbox $Mailbox1 -ForwardingAddress $ForwardAddress -DeliverToMailboxAndForward $false
                            $TextWeiterleitung.Text = $MsgFwSet 
                        }
                        if (($Weiterleitung.Checked) -and ($WeiterleitungMailKopie.Checked))
                        { 
                            $ForwardAddress = $WeiterleitungMail.Text
                            Set-Mailbox $Mailbox1 -ForwardingAddress $ForwardAddress -DeliverToMailboxAndForward $true
                            $TextWeiterleitung.Text = $MsgFwSetCopy
                        }
                        if (!$Weiterleitung.Checked)
                        {
                            Set-Mailbox $Mailbox1 -ForwardingAddress $null
                            $TextWeiterleitung.Text = $MsgFwSetNo 
                        }
                        }
                      } # Beendet die Funktion Speichern


function Abmelden {
                
                Remove-PSSession -Name Exchange               
                $Ausgabe1.text = $MsgLogoff
                $Fenster1.Close()
                
                
         } # Beendet die Funktion Abmelden                

<# Aktionen bei veränderten Radio Buttons für aktiviertes /deaktiviertes Out of Office
   Jenachdem welcher Status abgerfragt wird, bzw welcher Status gesetzt wird müssen die Buttons ein und ausgeschaltet werden #>

function Button_Status
{
        if ($RadioEnabled.Checked)
        {
                          
            $RadioDisabled.Enabled = $true
            $RadioEnabled.Enabled = $true                          
            $Extern.Enabled = $true
            $ExternKontakte.Enabled = $false
            $ExternAlle.Enabled = $False      
            $Zeit.Enabled = $true
            $ZeitStart.Enabled = $false
            $ZeitEnde.Enabled = $false
            $DatumStart.Enabled = $false
            $DatumEnde.Enabled = $false
            $HTMLForm1.Enabled = $true
            $HTMLForm2.Enabled = $true
         }
         else
         {
            $RadioDisabled.Enabled = $true
            $RadioEnabled.Enabled = $true                          
            $Extern.Enabled = $false
            $Extern.Checked = $false
            $ExternKontakte.Enabled = $false
            $ExternAlle.Enabled = $False                          
            $Zeit.Enabled = $false
            $Zeit.Checked = $false
            $ZeitStart.Enabled = $false
            $ZeitEnde.Enabled = $false
            $DatumStart.Enabled = $false
            $DatumEnde.Enabled = $false
            $HTMLForm1.Enabled = $false
            $HTMLForm2.Enabled = $false

         }
        if ($Extern.Checked)
        {
            $ExternKontakte.Enabled = $true
            $ExternAlle.Enabled = $true
        }
        else
        {
            $ExternKontakte.Enabled = $false
            $ExternKontakte.Checked = $false
            $ExternAlle.Enabled = $false
            $ExternAlle.Checked = $false
        }
        if ($Zeit.Checked)
        {
            $ZeitStart.Enabled = $true
            $ZeitEnde.Enabled = $true
            $DatumStart.Enabled = $true
            $DatumEnde.Enabled = $true
        }
        if ($Weiterleitung.Checked)
        {
            $WeiterleitungMail.Enabled = $true
            $WeiterleitungMailKopie.Enabled = $true
        }
        else
        {
            $WeiterleitungMail.Enabled = $false
            $WeiterleitungMailKopie.Enabled = $false   
        }
                                       
}

if ($RadioDisabled.Checked) {Button_Status}

$handler_UserArray = {
                    $Fenster1.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
                    $Fenster0.Show() | Out-Null
                    $UserName.Items.Clear();
                    if ($OU -notlike $null)
                            {
                            $UserArray = Get-User -OrganizationalUnit $OU -RecipientTypeDetails UserMailbox -SortBy Name -ResultSize Unlimited
                            }
                            else
                            {
                            $UserArray = Get-User -RecipientTypeDetails UserMailbox -SortBy Name -ResultSize Unlimited
                            }
                                            $i = 0
                                            foreach($user in $UserArray)
                                            {
                                            $i++
                                            [int]$pct = ($i/$UserArray.count)*100
                                            $progressbar1.Value = $pct
                                            $TextProgressBar.text="$MsgLoadUser $($user.name)"
                                            $Fenster0.Refresh()
                                            $UserName.items.add($user.Name)
                                            $WeiterleitungMail.items.add($user.Name)
                                            }
                    
                    $Fenster0.Close()
                    $Fenster1.Cursor = [System.Windows.Forms.Cursors]::Arrow
                    $Ausgabe1.Text = $MsUserLoadSuccess
                    }

$handler_Report = {
                    $Fenster1.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
                    if ($OU -notlike $null)
                            {
                            $QueryMBXReport = (Get-Mailbox -OrganizationalUnit $OU -ResultSize Unlimited -RecipientTypeDetails UserMailbox).Name
                            $QueryMBXReport | Get-MailboxAutoReplyConfiguration | ? {($_.AutoReplyState -eq "Enabled") -or ($_.AutoReplyState -eq "Scheduled") } |select Identity,AutoReplyState,ExternalAudience,StartTime,EndTime | Out-GridView -Title $MsgTextGridTitle
                            }
                            else
                            {
                            $QueryMBXReport = (Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox).Name
                            $QueryMBXReport | Get-MailboxAutoReplyConfiguration | ? {($_.AutoReplyState -eq "Enabled") -or ($_.AutoReplyState -eq "Scheduled") } |select Identity,AutoReplyState,ExternalAudience,StartTime,EndTime | Out-GridView -Title $MsgTextGridTitle
                           }
                    $Fenster1.Cursor = [System.Windows.Forms.Cursors]::Arrow
                    $Ausgabe1.Text = $MsUserLoadSuccess
                    }


#########################################################################################################################################
# Sämtliche Texte hinzufügen
      
    $TextExtern.Location = New-Object System.Drawing.Size(10,185)
    $TextExtern.Size = New-Object System.Drawing.Size(450,20)
    $GroupBox5.Controls.Add($TextExtern)

    # Text fuer die Start und Endzeit einbauen
    
    $TextZeitStart.Location = New-Object System.Drawing.Size(25,133)
    $TextZeitStart.Size = New-Object System.Drawing.Size(70,20)
    $GroupBox1.Controls.Add($TextZeitStart)
        
    $TextZeitEnde.Location = New-Object System.Drawing.Size(25,163)
    $TextZeitEnde.Size = New-Object System.Drawing.Size(60,20)
    $GroupBox1.Controls.Add($TextZeitEnde)

    #Text für Ergebnisausgabe E-Mail weiterleitung

    $TextWeiterleitung.Location = New-Object System.Drawing.Size(10,130)
    $TextWeiterleitung.Size = New-Object System.Drawing.Size(350,35)
    $GroupBox6.Controls.Add($TextWeiterleitung)
    
    # Info-Text Anzeigen lassen
    
    $About1.Location = New-Object System.Drawing.Size(5,10)
    $About1.Size =New-Object System.Drawing.Size(400,35)
    $About1.Text = "Exchange Server / O365 Out of Office Tool Version $VersionNumber
    © Andres Sichel // asichel.de // blog@asichel.de"
    $Fenster1.Controls.Add($About1)

   
    $TextProgressBar.Left=5
    $TextProgressBar.Top= 10
    $TextProgressBar.Width= 500 - 20
    $TextProgressBar.Height=15
    $Fenster0.controls.add($TextProgressBar)

#########################################################################################################################################
# Verbindungsoptionen
# Eingabefeld für Servernamen bauen

If ($SetServername -gt 0)
{ 
$SetServer = $SetServername
$ServerName.ReadOnly = $true
} 
else { $SetServer = "'Servername im FQDN'" }

$ServerName.Location = New-Object System.Drawing.Size(10,40)
$ServerName.Size = New-Object System.Drawing.Size(200,20)
$ServerName.Text = "$SetServer"
$ServerName.TabIndex = 1
$GroupBox2.Controls.Add($ServerName)

# Anmeldebutton bauen

$Login.Location = New-Object System.Drawing.Size(250,40)
$Login.Size = New-Object System.Drawing.Size(80,22)
$Login.TabIndex = 3
$Login.BackColor = "white"
$Login.Add_Click({Anmelden})
$GroupBox2.Controls.Add($Login)

# Abfrage aktuelle Benutzer
If ($LoggedOnUser -eq "Yes")
{ $WindowsUser.Checked = $true
  $WindowsUser.Enabled = $false  
}
else { $WindowsUser.Checked = $false }

$WindowsUser.Location = New-Object System.Drawing.Size(10,70)
$WindowsUser.Size = New-Object System.Drawing.Size(250,20)
$WindowsUser.TabIndex = 2
$GroupBox2.Controls.Add($WindowsUser)

# Eingabefeld für Usernamen bauen
$UserName.Location = New-Object System.Drawing.Size(10,100)
$UserName.Size = New-Object System.Drawing.Size(200,20)
$UserName.AutoCompleteMode = 3
$UserName.AutoCompleteSource = 256
$UserName.DataBindings.DefaultDataSourceUpdateMode = 0
$UserName.FormattingEnabled = $True
$UserName.Enabled = $false
$UserName.TabIndex = 4
$GroupBox2.Controls.Add($UserName)

# Abfragebutton bauen
$Abfrage.Location = New-Object System.Drawing.Size(250,100)
$Abfrage.Size = New-Object System.Drawing.Size(80,22)
$Abfrage.TabIndex = 5
$Abfrage.BackColor = "white"
$Abfrage.Add_Click({Abfrage})
$GroupBox2.Controls.Add($Abfrage)

# Speicherbutton bauen
$Speichern.Location = New-Object System.Drawing.Size(25,140)
$Speichern.Size = New-Object System.Drawing.Size(170,22)
$Speichern.TabIndex = 20
$Speichern.BackColor = "white"
$Speichern.Add_Click({Speichern})
$GroupBox2.Controls.Add($Speichern)


# Benutzer neu laden bauen
$BenutzerLaden.Location = New-Object System.Drawing.Size(230,140)
$BenutzerLaden.Size = New-Object System.Drawing.Size(120,40)
$BenutzerLaden.TabIndex = 7
$BenutzerLaden.Enabled = $false
$BenutzerLaden.BackColor = "white"
$BenutzerLaden.Add_Click($handler_UserArray)
$GroupBox2.Controls.Add($BenutzerLaden)

#Reportbutton bauen
$ReportGrid.Location = New-Object System.Drawing.Size(5,20)
$ReportGrid.Size = New-Object System.Drawing.Size(200,20)
$ReportGrid.TabIndex = 7
$ReportGrid.Enabled = $false
$ReportGrid.BackColor = "white"
$ReportGrid.Add_Click($handler_Report)
$GroupBox7.Controls.Add($ReportGrid)

# Abmeldebutton bauen
$ToolTip.SetToolTip($ExitBild, "Abmelden & Exit")
$ExitBild.Location = New-Object System.Drawing.Size(790,22)
$ExitBild.ImageLocation = "$scriptRoot\exit.png"
$ExitBild.SizeMode = "AutoSize"
$ExitBild.Text = "Exit"
$ExitBild.TabIndex = 21
$ExitBild.add_mouseclick{Abmelden}
$Fenster1.Controls.Add($ExitBild)


#########################################################################################################################################
#HTML Felder

# Interne Nachricht
$HTMLForm1.Location = New-Object System.Drawing.Size(10,25)
$HTMLForm1.Size = New-Object System.Drawing.Size(450,150)
$HTMLForm1.ToolbarStyle = 2
$HTMLForm1.TabIndex = 18
$HTMLForm1.Enabled = $false
$GroupBox5.Controls.Add($HTMLForm1)

# Externe Nachricht
$HTMLForm2.Location = New-Object System.Drawing.Size(10,205)
$HTMLForm2.Size = New-Object System.Drawing.Size(450,150)
$HTMLForm2.ToolbarStyle = 2
$HTMLForm2.TabIndex = 19
$HTMLForm2.Enabled = $false
$GroupBox5.Controls.Add($HTMLForm2)


#########################################################################################################################################
# Optionsfelder erstellen
# Auswahlbuttons erstellen
$RadioDisabled.Location = New-Object System.Drawing.Size(10,25)
$RadioDisabled.Size = New-Object System.Drawing.Size(350,20)
$RadioDisabled.Enabled = $false
$RadioDisabled.TabIndex = 8
$RadioDisabled.add_Click({Button_Status})
$GroupBox3.Controls.Add($RadioDisabled)


$RadioEnabled.Location = New-Object System.Drawing.Size(10,50)
$RadioEnabled.Size = New-Object System.Drawing.Size(350,20)
$RadioEnabled.Enabled = $false
$RadioEnabled.TabIndex = 9
$RadioEnabled.add_Click({Button_Status})
$GroupBox3.Controls.Add($RadioEnabled)

# Checkbox auch für Externe Empfänger
$Extern.Location = New-Object System.Drawing.Size(10,25)
$Extern.Size = New-Object System.Drawing.Size(370,20)
$Extern.TabIndex = 10
$Extern.add_Click({Button_Status})
$GroupBox1.Controls.Add($Extern) 

# RadioButton Extern nur Kontaktliste
$ExternKontakte.Location = New-Object System.Drawing.Size(25,45)
$ExternKontakte.Size = New-Object System.Drawing.Size(380,20)
$ExternKontakte.TabIndex = 11
$GroupBox1.Controls.Add($ExternKontakte)

# RadioButton Extern alle
$ExternAlle.Location = New-Object System.Drawing.Size(25,70)
$ExternAlle.Size = New-Object System.Drawing.Size(380,20)
$ExternAlle.TabIndex = 12
$GroupBox1.Controls.Add($ExternAlle)

# Checkbox Zeitintervall
$Zeit.Location = New-Object System.Drawing.Size(10,100)
$Zeit.Size = New-Object System.Drawing.Size(370,20)
$Zeit.add_Click({Button_Status})
$Zeit.TabIndex = 13
$GroupBox1.Controls.Add($Zeit)

# StartzeitFeld einbauen
$ZeitStart.Location =  New-Object System.Drawing.Size(100,130)
$ZeitStart.Size = New-Object System.Drawing.Size(70,20)
$ZeitStart.CustomFormat = "HH:mm"
$ZeitStart.Format = 'Custom'
$ZeitStart.Name = "Startzeit"
$TextZeitStart.TabIndex = 14
$ZeitStart.ShowUpDown = $true
$GroupBox1.Controls.Add($ZeitStart)

# StartdatumFeld einbauen
$DatumStart.Location =  New-Object System.Drawing.Size(180,130)
$DatumStart.Size = New-Object System.Drawing.Size(130,20)
$DatumStart.CustomFormat = " ddd dd.MM.yyyy"
$DatumStart.Format = 'Custom'
$DatumStart.Name = "Startdatum"
$DatumStart.TabIndex = 15
$GroupBox1.Controls.Add($DatumStart)

# EndzeitFeld einbauen
$ZeitEnde.Location =  New-Object System.Drawing.Size(100,160)
$ZeitEnde.Size = New-Object System.Drawing.Size(70,20)
$ZeitEnde.CustomFormat = "HH:mm"
$ZeitEnde.Format = 'Custom'
$ZeitEnde.Name = "Endzeit"
$ZeitEnde.TabIndex = 16
$ZeitEnde.ShowUpDown = $true
$GroupBox1.Controls.Add($ZeitEnde)

# EnddatumFeld einbauen
$DatumEnde.Location =  New-Object System.Drawing.Size(180,160)
$DatumEnde.Size = New-Object System.Drawing.Size(130,20)
$DatumEnde.CustomFormat = " ddd dd.MM.yyyy"
$DatumEnde.Format = 'Custom'
$DatumEnde.Name = "Enddatum"
$DatumEnde.TabIndex = 17
$GroupBox1.Controls.Add($DatumEnde)

# Ausgabetextfeld bauen
$Ausgabe1.Location = New-Object System.Drawing.Size(5,25)
$Ausgabe1.Size = New-Object System.Drawing.Size(200,150)
$Ausgabe1.Font = $Font 
$GroupBox4.Controls.Add($Ausgabe1)

# Ststusbild anzeigen lassen
$ToolTip.SetToolTip($StatusBild, "Out of Office Status")
$StatusBild.Location = New-Object System.Drawing.Size(225,25)
$StatusBild.SizeMode = "AutoSize"
$StatusBild.Text = "Out of Office Status"
$StatusBild.TabIndex = 21
$GroupBox4.Controls.Add($StatusBild)

#########################################################################################################################################
# Erweiterte optionen

# Checkbox Zeitintervall
$Weiterleitung.Location = New-Object System.Drawing.Size(10,25)
$Weiterleitung.Size = New-Object System.Drawing.Size(320,20)
$Weiterleitung.Enabled = $false


$Weiterleitung.add_Click({Button_Status})
$Weiterleitung.TabIndex = 13
$GroupBox6.Controls.Add($Weiterleitung)

$WeiterleitungMail.Location = New-Object System.Drawing.Size(25,45)
$WeiterleitungMail.Size = New-Object System.Drawing.Size(300,20)
$WeiterleitungMail.AutoCompleteMode = 3
$WeiterleitungMail.AutoCompleteSource = 256
$WeiterleitungMail.DataBindings.DefaultDataSourceUpdateMode = 0
$WeiterleitungMail.FormattingEnabled = $True
$WeiterleitungMail.TabIndex = 1
$WeiterleitungMail.Enabled = $false
$GroupBox6.Controls.Add($WeiterleitungMail)

$WeiterleitungMailKopie.Location = New-Object System.Drawing.Size(25,70)
$WeiterleitungMailKopie.Size = New-Object System.Drawing.Size(340,40)
$WeiterleitungMailKopie.Enabled = $false
$WeiterleitungMailKopie.add_Click({Button_Status})
$WeiterleitungMailKopie.TabIndex = 13
$GroupBox6.Controls.Add($WeiterleitungMailKopie)

#########################################################################################################################################
# Aktion für automatsiche Anmledung ausführen

<# Dies ist keine Funktion. Ist der Wert zur Automatischen anmeldung auf "Yes" festgelegt,
   so wird unmittelbar nach Programmstart die Anmeldung ausgeführt!#>

    if (($LogOnatStart -eq "Yes") -and ($SetServername -gt 0) -and ($LoggedOnUser -eq "Yes"))
        {             
         Anmelden
        }
        
    
    if ($LogOnatStart -eq "no")
        {
        $Ausgabe1.Text = $MsgAutoLogonDisabeld
        $Fenster0.hide()
        }
    if (($LogOnatStart -eq "Yes") -and ($SetServername -lt 0) -and ($LoggedOnUser -eq "No"))
        {
        $AutoLogonFail = [Windows.Forms.MessageBox]
        $AutoLogonFail::Show($MsgAutoLogonFailedText , $MsgAutoLogonFailedTitle, "OK", "Error")    
        }
    if (($LogOnatStart -eq "Yes") -and ($SetServername -lt 0) -and ($LoggedOnUser -eq "Yes"))
        {
        $AutoLogonFail = [Windows.Forms.MessageBox]
        $AutoLogonFail::Show($MsgAutoLogonFailedText , $MsgAutoLogonFailedTitle, "OK", "Error")        
        }

#########################################################################################################################################
# Fenster anzeigen lassen
[System.Windows.Forms.Application]::Run($Fenster1)
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::DoEvents()
$Fenster1.Add_Shown({$Fenster1.Activate()})
$notifyIcon.Dispose()