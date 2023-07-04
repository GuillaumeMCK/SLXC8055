# Variable globale
$recherche = "apManager	crash"

# Importation de Selenium dans powershell https://www.selenium.dev/documentation/en/webdriver/
Import-Module "$($PSScriptRoot)\WebDriver.dll"

function chromeDriverUpdater {

    $url_LastRelease = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE_"
    $url_Download = "https://chromedriver.storage.googleapis.com/"
    $chromeDriverDir = "$PSScriptRoot\chromedriver.exe"

    # Récuperation de la version de chrome installé
    Try { $ChromeVer = (Get-Item (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe' -ErrorAction Stop).'(Default)').VersionInfo.FileVersion }
    Catch { Write-Host "Google Chrome n'a pas été trouver dans le registre." -ForegroundColor Red -ErrorAction Stop }
    
    # Récupération de la version de chromeDriver
    if (Test-Path $chromeDriverDir){
        $ChromeDriverVer = (& $chromeDriverDir -v).Split(" ")[1]
    }
    else { $ChromeDriverVer = '' }

    # Obtention de la version global de chrome installé
    $globalChromeVer = $ChromeVer.Split(".")[0..2] -join(".")
    
    # Récuperation de la dernier version de ChromeDriver compatible avec chrome installé
    try{ $ChromeDriverlastestVer = Invoke-RestMethod -Uri $url_LastRelease$globalChromeVer }
    catch{Write-Host "Une erreur est survenue lors de la récupération de la dernière version de chrome driver. Essayer de relancer le script dans un nouveau terminal si nécessaire." -ForegroundColor Red}
    if ($ChromeDriverVer -notmatch $ChromeDriverlastestVer){
        
        $url_Download = $url_Download + $ChromeDriverlastestVer + "/chromedriver_win32.zip"
        $DownloadFileDir = $PSScriptRoot+"\chromedriver_win32.zip"
        # Telechargement
        Invoke-WebRequest -Uri $url_Download -OutFile $DownloadFileDir 
        # Extraction
        Expand-Archive $DownloadFileDir -DestinationPath "$PSScriptRoot" -Force
        # Supression de l'archive
        Remove-Item -Path $DownloadFileDir -Force

        $ChromeDriverVer = (& $chromeDriverDir -v).Split(" ")[1]
        return (Write-Host "ChromeDriver a été mise a jour vers $ChromeDriverVer" -ForegroundColor Gray)
    }
    else { return (Write-Host "ChromeDriver est a jour." -ForegroundColor Gray) }
}

function checkUrlStatus($url) {

    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

    # Creation de la requete
    $Request = [System.Net.WebRequest]::Create($url)
    
    try{
        # Réponse du site.
        $Response = $Request.GetResponse()
    
        # On reçois le status HTTP sous forme d'entier.
        $Status = [int]$Response.StatusCode
        
        If ($Status -eq 200) { return $true }
        Else { return $false }
    }
    catch{ }
    
    # Fermeture de la requete
    If ($Response -eq $null) { } 
    Else { $Response.Close() }
    
}

function main ([string]$InputPrinterName, [string]$research, [bool]$inputList){
    
    Write-Host 'Début du script.' -ForegroundColor Gray

    ########## SETUP ##########

    chromeDriverUpdater

    $nameOutPutFile = "$PSScriptRoot\output$(Get-Date -Format "_yyyyMMdd_HHmms").txt"

    Write-Host 'Lancement du Web Driver.' -ForegroundColor Gray

    # Definition des options pour le chromeDriver
    $downloadDir = "$($PSScriptRoot)\.tmp"
    $service = [OpenQA.Selenium.Chrome.ChromeDriverService]::CreateDefaultService()
    $service.HideCommandPromptWindow = $true  
    $ChromeOptions = New-Object OpenQA.Selenium.Chrome.ChromeOptions
    $ChromeOptions.AddArgument("--headless")
    $ChromeOptions.AddArgument("--disable-extensions")
    $ChromeOptions.AddArgument("--no-sandbox")
    $ChromeOptions.AddArgument("--disable-infobars")
    $ChromeOptions.AddUserProfilePreference("safebrowsing.enabled", "true")
    $ChromeOptions.AddUserProfilePreference("download.default_directory", "$downloadDir")
    $ChromeOptions.AddUserProfilePreference("download.prompt_for_download", "false")
    $ChromeOptions.AddUserProfilePreference("download.directory_upgrade", "true")
    $ChromeOptions.AddUserProfilePreference("prompt_for_download", "true")
    $ChromeOptions.AcceptInsecureCertificates = $True
    # Création de l'objet ChromeDriver 
    try {
        $ChromeDriver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($service, $ChromeOptions)   
        Write-Host 'ChromeDriver a été démarré avec succès.' -ForegroundColor Green
    }
    catch {
        Stop-Process -Name chromedriver -Force | Out-Null # Kill lost ChromeDriver process  
        Write-Error "Une erreur est survenue au lancement de chromeDriver." -ErrorAction Stop
    }
    # Creation du fichier temporaire pour manipuler les fichiers telechargés
    If(!(test-path $downloadDir)){
        New-Item -ItemType Directory -Force -Path $downloadDir| Out-Null
    }
    else{ 
        Remove-Item -Path $downloadDir -Recurse:$true
        New-Item -ItemType Directory -Force -Path $downloadDir| Out-Null 
    }

    # Obtention des identifiants
    $cred = Get-Credential -Message "Veuillez saisir l'identifiant administrateur des imprimantes."

    # Tableau dynamique contenant les informations recuperé par le script pour chaque imprimante
    $output = New-Object System.Collections.Generic.List[System.Object] 

    if ($inputList){
        # recuperation de la liste contenant les noms des imprimantes
        try{ $list = Get-content $InputPrinterName }
        catch{ Write-Host "La liste contenant les noms des imprimantes n'a pas été trouvé" -ForegroundColor Red}
    }
    else{ $list = @($InputPrinterName) }
   
    ########## BOT ##########
        
    foreach ($printerName in $list){
        # Variables
        $url_LoginPage = "https://$printerName/properties/authentication/login.php"
        $url_Info = "https://$printerName/properties/description.php"
        $url_LogPage = "https://$printerName/properties/security/auditlog.php"
        $url_LogDownload = "https://$printerName/properties/security/auditLogDownload.php"

        Write-Host "$printerName -> " -ForegroundColor Gray -NoNewline

        $output.Add("______________| $printerName |____________________________")

        # Test de connexion a la page web
        if (checkStatusOfUrl $url_LoginPage){

            # Connexion a l'interface Web
            $ChromeDriver.Navigate().GoToUrl($url_LoginPage)       
           
            # Remplissage du formulaire de connexion
            $ChromeDriver.FindElementByXPath('//*[@id="frmwebUsername"]').SendKeys($cred.UserName)
            $ChromeDriver.FindElementByXPath('//*[@id="frmwebPassword"]').SendKeys($cred.GetNetworkCredential().Password) 
            $ChromeDriver.FindElementByXPath('//*[@id="loginBtn"]').Click()
       
            if ($ChromeDriver.Url -match "autoConfiguration.php"){
                
                # Récuperation du numero de série
                $ChromeDriver.Navigate().GoToUrl($url_Info)
                $serialNbr = $ChromeDriver.FindElementByXPath('//*[@id="widSerNum"]/div[1]/div').Text
                
                # Changement d'url pour avoir la page audit log
                $ChromeDriver.Navigate().GoToUrl($url_LogPage)
                
                # Generation du fichier log
                $ChromeDriver.FindElementByXPath('//*[@id="saveFile"]').Click()
                
                # Telechargement du fichier
                if ($ChromeDriver.Url -match "auditLogReady.php"){
                    
                    $ChromeDriver.Navigate().GoToUrl($url_LogDownload)
                    # Attente du telechargement
                    Start-Sleep -s 2
                    
                    # Decompression
                    $ZipName = Get-ChildItem -Path $downloadDir -Name -Filter *$serialNbr*
                    Expand-Archive "$downloadDir\$ZipName" -DestinationPath "$downloadDir" -Force
                    
                    # Lecture du fichier log
                    $LogFile = Get-Content -Path "$downloadDir\auditfile.txt"
             
                    for ($line = $LogFile.Length; $line -ge 0 ; $line--){
                        if ($LogFile[$line] -match $research){
                            $output.Add($LogFile[$line-1])                      
                            $output.Add($LogFile[$line])
                            $output.Add($LogFile[$line+1])
                            $output.Add("...")
                            Write-Host $LogFile[$line].Split("	")[1..2]
                            break
                            
                        }
                        elseif ($line -eq 0) { 
                            $output.Add("Aucune trace de '$($research)'.")
                            Write-Host "Aucune trace de '$($research)'"
                       } 
                    }
                }
            }
            else{
                $output.Add("Identifiant incorrecte.") 
                Write-Host 'Identifiant incorrecte.' -ForegroundColor Red 
            }
        }
        else{
            $output.Add("$printerName est inacessible.")
            Write-Host 'Inacessible.' -ForegroundColor Red 
        }
    }

    ########## CLEANUP ##########

    Write-Host 'Nettoyage.' -ForegroundColor Gray
    Remove-Item -Path $downloadDir -Recurse:$true
    try{
        $ChromeDriver.Close();
        $ChromeDriver.Quit();
    }
    catch { Stop-Process -Name chromedriver -Force }

    Write-Host 'Fin du script.' -ForegroundColor Gray

    # Recap
    $output | Out-File $nameOutPutFile -Force
    notepad.exe $nameOutPutFile
    exit
}

# Entry Point
if ($args[0] -eq $null){
    $name = Read-Host "Imprimantes a tester"
    main $name $recherche $false
}
else{
    main $args[0] $recherche $true
}
