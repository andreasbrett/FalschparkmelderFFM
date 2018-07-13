Import-Module PSSQLite
Add-Type -AssemblyName System.Web
[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
[System.Windows.Forms.Application]::EnableVisualStyles() | Out-Null



# ------------------------------------------------------------------------------------------------------------
#  > MAIL SETTINGS
# ------------------------------------------------------------------------------------------------------------
    $Global:mail_receiver = "owi.datenerfassung.amt32@stadt-frankfurt.de"
    $Global:mail_subject = "Anzeige: Verkehrsordnungswidrigkeit"




# ------------------------------------------------------------------------------------------------------------
#  > DATABASE SETTINGS
# ------------------------------------------------------------------------------------------------------------
    $Global:db = "$($PSScriptRoot)\$((Get-Item $PSCommandPath).BaseName).db"
    $Global:dbconn = $null




# ------------------------------------------------------------------------------------------------------------
#  > DATABASE FUNCTIONS
# ------------------------------------------------------------------------------------------------------------
    Function exportGPX($incidents) {
        # show save file dialog
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.InitialDirectory = "."
        $saveFileDialog.Filter = "GPX (*.gpx)| *.gpx"
        $saveFileDialog.ShowDialog() | Out-Null
        $sFilePath = $saveFileDialog.FileName
    
        # GPX Header
        '<?xml version="1.0" encoding="UTF-8" standalone="no" ?>' | Out-File -Encoding utf8 $sFilePath
        '<gpx version="1.1" creator="FalschParkMelder">' | Out-File -Encoding utf8 $sFilePath -Append

        # iterate over incidents
        ForEach($incident in $incidents) {
            '
            <wpt lat="' + $incident.location_lat + '" lon="' + $incident.location_lon + '">
                <time>' + ($incident.timestamp).ToString("o") + '</time>
                <name>' + $incident.car_model + ' (' + $incident.car_color + ')</name>
                <desc>' + ($incident.timestamp).ToString("g") + ' | ' + $incident.location + '</desc>
            </wpt>
            ' | Out-File -Encoding utf8 $sFilePath -Append
        }

        # GPX Footer
        '</gpx>' | Out-File -Encoding utf8 $sFilePath -Append
    }
        


    Function geoDecode($sLocation) {
        $sLocation = [System.Web.HttpUtility]::UrlEncode($sLocation)
        $url = "https://geocode.xyz/$($sLocation)?json=1"
        $result = Invoke-WebRequest $url | ConvertFrom-Json

        If ($result.error) {
            $null
        } Else {
            $result | Select longt, latt
        }
    }



    Function saveBlobToFile($blob, $sFilePath) {
        [System.IO.File]::WriteAllBytes($sFilePath, $blob)
    }



    Function getPhotoDateTaken($sFilePath) {
        # fetch date taken from EXIF metadata
        $oImg = New-Object System.Drawing.Bitmap($sFilePath)
        $ExifDate = $oImg.GetPropertyItem(36867)
        $oImg.Dispose()

        # convert to DateTime object
        $DateTaken = (New-Object System.Text.UTF8Encoding).GetString($ExifDate.Value)
        [DateTime]::ParseExact($DateTaken, "yyyy:MM:dd HH:mm:ss`0", $null)
    }



    Function openDbConnection() {
        $Global:dbconn = New-SQLiteConnection @Verbose -DataSource $db 
    }



    Function closeDbConnection() {
        $Global:dbconn.Close()
    }



    Function createTablesIfNotExisting() {
        $qryCreateTable_users = "
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
                forename NVARCHAR(250) NOT NULL,
                lastname NVARCHAR(250) NOT NULL,
                street NVARCHAR(250) NOT NULL,
                zip NVARCHAR(250) NOT NULL,
                city NVARCHAR(250) NOT NULL,
                mail_address NVARCHAR(250) NOT NULL,
                mail_server NVARCHAR(250) NOT NULL,
                mail_server_port NVARCHAR(250) NOT NULL,
                mail_user NVARCHAR(250) NOT NULL,
                mail_password NVARCHAR(1024) NOT NULL
            );
        "
        Invoke-SqliteQuery -Connection $Global:dbconn -Query $qryCreateTable_users

        $qryCreateTable_incidents = "
            CREATE TABLE IF NOT EXISTS incidents (
                id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
                user_id INTEGER NOT NULL,
                timestamp DATE NOT NULL,
                location NVARCHAR(250) NOT NULL,
                location_lat NUMERIC,
                location_lon NUMERIC,
                car_license_tag NVARCHAR(250) NOT NULL,
                car_model NVARCHAR(250) NOT NULL,
                car_color NVARCHAR(250) NOT NULL,
                photo BLOB NOT NULL,
                timestamp_sent DATE
            );
        "
        Invoke-SqliteQuery -Connection $Global:dbconn -Query $qryCreateTable_incidents
    }



    Function addUser($sForename, $sLastname, $sStreet, $sZip, $sCity, $sMailAddress, $sMailServer, $sMailServerPort, $sMailUser, $sMailPassword) {
        $qryAddUser = "
            INSERT INTO
                users(forename, lastname, street, zip, city, mail_address, mail_server, mail_server_port, mail_user, mail_password)
                VALUES(@forename, @lastname, @street, @zip, @city, @mail_address, @mail_server, @mail_server_port, @mail_user, @mail_password)
            ;
        "

        # add entry to database
        Invoke-SqliteQuery -SQLiteConnection $Global:dbconn -Query $qryAddUser -SqlParameters @{
            forename = $sForename
            lastname = $sLastname
            street = $sStreet
            zip = $sZip
            city = $sCity
            mail_address = $sMailAddress
            mail_server = $sMailServer
            mail_server_port = $sMailServerPort
            mail_user = $sMailUser
            mail_password = ($sMailPassword | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString)
        }
    }



    Function removeUser($id, $stopIfUnsentIncidents) {
        # stop if user has unsent incidents
        If ($stopIfUnsentIncidents) {
            If (fetchIncidentsByUserUnsent -userId $id) { return $true }
        }

        # remove user from database
        $qryRemoveUser = "DELETE FROM users WHERE id = @id"
        Invoke-SqliteQuery -SQLiteConnection $Global:dbconn -Query $qryRemoveUser -SqlParameters @{id=$id}
    }



    Function editUser($id, $sForename, $sLastname, $sStreet, $sZip, $sCity, $sMailAddress, $sMailServer, $sMailServerPort) {
        $qryEditUser = "
            UPDATE users
            SET
                forename = @forename,
                lastname = @lastname,
                street = @street,
                zip = @zip,
                city = @city,
                mail_address = @mail_address,
                mail_server = @mail_server,
                mail_server_port = @mail_server_port
            WHERE
                id = @id
            ;
        "

        # update entry in database
        Invoke-SqliteQuery -SQLiteConnection $Global:dbconn -Query $qryEditUser -SqlParameters @{
            id = $id
            forename = $sForename
            lastname = $sLastname
            street = $sStreet
            zip = $sZip
            city = $sCity
            mail_address = $sMailAddress
            mail_server = $sMailServer
            mail_server_port = $sMailServerPort
        }
    }



    Function editUserCredentials($id, $sMailUser, $sMailPassword) {
        $qryUpdateCredentials = "
            UPDATE users
            SET
                mail_user = @mail_user,
                mail_password = @mail_password
            WHERE
                id = @id
            ;
        "

        # update entry in database
        Invoke-SqliteQuery -SQLiteConnection $Global:dbconn -Query $qryUpdateCredentials -SqlParameters @{
            id = $id
            mail_user = $sMailUser
            mail_password = ($sMailPassword | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString)
        }
    }



    Function addIncident($userId, $sPathToPhoto, $sLocation, $sCarLicenseTag, $sCarModel, $sCarColor) {
        $qryAddIncident = "
            INSERT INTO
                incidents(timestamp, user_id, location, location_lat, location_lon, car_license_tag, car_model, car_color, photo)
                VALUES(@timestamp, @user_id, @location, @location_lat, @location_lon, @car_license_tag, @car_model, @car_color, @photo)
            ;
        "

        # fetch exif timestamp from photo
        $timestamp = getPhotoDateTaken($sPathToPhoto)

        # try and fetch GPS coordinates for location
        $gps = geoDecode -sLocation "$($sLocation) Frankfurt Main Germany"
        If ($gps) {
            $location_lon = $gps.longt
            $location_lat = $gps.latt
        } Else {
            $location_lon = $null
            $location_lat = $null
        }

        # load photo into blob
        $photoBlob = [System.IO.File]::ReadAllBytes($sPathToPhoto)

        # add entry to database
        Invoke-SqliteQuery -SQLiteConnection $Global:dbconn -Query $qryAddIncident -SqlParameters @{
            timestamp = $timestamp
            user_id = $userId
            location = $sLocation
            location_lat = $location_lat
            location_lon = $location_lon
            car_license_tag = $sCarLicenseTag
            car_model = $sCarModel
            car_color = $sCarColor
            photo = $photoBlob
        }
    }



    Function sendIncident($id) {
    
        # fetch incident and user
        $incident = fetchIncident -id $id
        $user = fetchUserById -id $incident.user_id

        # temporarily extract photo
        $tmpPhoto = New-TemporaryFile
        saveBlobToFile -blob $incident.photo -sFilePath $tmpPhoto.FullName
        $rnd = Get-Random -Minimum 10000 -Maximum 99999
        Rename-Item -Path $tmpPhoto -NewName "falschparker_$($rnd).jpg"
        $tmpPhoto = "$($tmpPhoto.Directory.FullName)\falschparker_$($rnd).jpg"

        # build body
        $body = "
            Tattag:  $(($incident.timestamp).ToString("dd.MM.yyyy"))
            Tatzeit: $(($incident.timestamp).ToString("HH:mm:ss"))
            Tatort:  $($incident.location)

            KfZ-Kennzeichen:     $($incident.car_license_tag)
            KfZ-Marke und Farbe: $($incident.car_model) ($($incident.car_color))

            genauer Tatvorwurf:
            Halten/Parken auf einem Radweg/Gehweg mit Behinderung von Radfahrern/Fußgängern, die ausweichen mussten

            vollständiger Name und Anschrift des Anzeigenden:
            $($user.forename) $($user.lastname)
            $($user.street)
            $($user.zip) $($user.city)

            Ein Beweisfoto aus dem Kennzeichen und Tatvorwurf erkennbar hervorgehen, befindet sich im Anhang
        "

        try {
			# send mail
			$utf8 = New-Object System.Text.utf8encoding
			$cred = New-Object System.Management.Automation.Pscredential -Argumentlist ($user.mail_user), ($user.mail_password | ConvertTo-SecureString)
            Send-MailMessage -ErrorAction Stop -Encoding $utf8 -From $user.mail_address -To $Global:mail_receiver -Subject $Global:mail_subject -Body $body -SmtpServer $user.mail_server -Port $user.mail_server_port -UseSsl -Credential $cred -Attachments $tmpPhoto
		}
		catch {
			Write-Error $_.Exception.Message
			return $false
		}
		
		# delete temporary photo
		Remove-Item $tmpPhoto

		# update timestamp-sent
		$qryMarkIncident = "UPDATE incidents SET timestamp_sent = @timestamp WHERE id = @id;"
		Invoke-SqliteQuery -SQLiteConnection $Global:dbconn -Query $qryMarkIncident -SqlParameters @{
			id = $incident.id
			timestamp = (Get-Date)
		}
		
		return $true
	
    }



    Function removeIncident($id) {
        $qryRemoveIncident = "DELETE FROM incidents WHERE id = @id"
        Invoke-SqliteQuery -SQLiteConnection $Global:dbconn -Query $qryRemoveIncident -SqlParameters @{id=$id}
    }



    Function fetchUsers() {
        $qryFetchUsers = "SELECT * FROM users;"
        Invoke-SqliteQuery -SQLiteConnection $Global:dbconn -Query $qryFetchUsers
    }



    Function fetchUser($sMailAddress) {
        $qryFetchUser = "SELECT * FROM users WHERE mail_address = @mail_address;"
        Invoke-SqliteQuery -SQLiteConnection $Global:dbconn -Query $qryFetchUser -SqlParameters @{mail_address=$sMailAddress}
    }



    Function fetchUserById($id) {
        $qryFetchUserById = "SELECT * FROM users WHERE id = @id;"
        Invoke-SqliteQuery -SQLiteConnection $Global:dbconn -Query $qryFetchUserById -SqlParameters @{id=$id}
    }



    Function fetchIncidentsByUser($userId) {
        $qryFetchIncidentsByUser = "SELECT * FROM incidents WHERE user_id = @user_id;"
        Invoke-SqliteQuery -SQLiteConnection $Global:dbconn -Query $qryFetchIncidentsByUser -SqlParameters @{user_id=$userId}
    }



    Function fetchIncidentsByUserSent($userId) {
        $qryFetchIncidentsByUser = "SELECT * FROM incidents WHERE user_id = @user_id AND timestamp_sent IS NOT NULL;"
        Invoke-SqliteQuery -SQLiteConnection $Global:dbconn -Query $qryFetchIncidentsByUser -SqlParameters @{user_id=$userId}
    }



    Function fetchIncidentsByUserUnsent($userId) {
        $qryFetchIncidentsByUser = "SELECT * FROM incidents WHERE user_id = @user_id AND timestamp_sent IS NULL;"
        Invoke-SqliteQuery -SQLiteConnection $Global:dbconn -Query $qryFetchIncidentsByUser -SqlParameters @{user_id=$userId}
    }



    Function fetchIncident($id) {
        $qryFetchIncident = "SELECT * FROM incidents WHERE id = @id;"
        Invoke-SqliteQuery -SQLiteConnection $Global:dbconn -Query $qryFetchIncident -SqlParameters @{id=$id}
    }




# ------------------------------------------------------------------------------------------------------------
#  > GUI FUNCTIONS
# ------------------------------------------------------------------------------------------------------------
    Function makeLabel($form, $text, $posX, $posY, $width, $height) {
        $obj = New-Object System.Windows.Forms.Label
        $obj.Location = New-Object System.Drawing.Size($posX, $posY)
        $obj.Size = New-Object System.Drawing.Size($width, $height)
        $obj.Text = $text
        $obj.AutoSize = $true
        $form.Controls.Add($obj)
    }



    Function makeCombo($form, $posX, $posY, $width, $height, $height2) {
        $obj = New-Object System.Windows.Forms.Combobox
        $obj.Location = New-Object System.Drawing.Size($posX, $posY)
        $obj.Size = New-Object System.Drawing.Size($width, $height)
        $obj.Height = $height2
        $form.Controls.Add($obj)
        return $obj
    }



    Function makeButton($form, $text, $posX, $posY, $width, $height) {
        $obj = New-Object System.Windows.Forms.Button
        $obj.Location = New-Object System.Drawing.Size($posX, $posY)
        $obj.Size = New-Object System.Drawing.Size($width, $height)
        $obj.Text = $text
        $obj.AutoSize = $true
        $form.Controls.Add($obj)
        return $obj
    }



    Function makeListView($form, $posX, $posY, $width, $height) {
        $obj = New-Object System.Windows.Forms.ListView
        $obj.Location = New-Object System.Drawing.Size($posX, $posY)
        $obj.Size = New-Object System.Drawing.Size($width, $height)
        $obj.MultiSelect = $false
        $obj.View = "Details"
        $obj.FullRowSelect = $true
        $obj.MultiSelect = $true
        $obj.GridLines = $true
        $form.Controls.Add($obj)
        return $obj
    }



    Function makeInput($form, $posX, $posY, $width, $height) {
        $obj = New-Object System.Windows.Forms.TextBox
        $obj.Location = New-Object System.Drawing.Size($posX, $posY)
        $obj.Size = New-Object System.Drawing.Size($width, $height)
        $form.Controls.Add($obj)
        return $obj
    }



    Function makePictureBox($form, $image, $posX, $posY, $width, $height) {
        $obj = New-Object System.Windows.Forms.PictureBox
        $obj.Width = $width
        $obj.Height = $height
        $obj.Location = New-Object System.Drawing.Size($posX, $posY)
        $obj.Image = $image
        $obj.SizeMode = "Zoom"
        $form.Controls.Add($obj)
        return $obj
    }



    Function makeForm($title, $width, $height, $backColor) {
        $obj = New-Object System.Windows.Forms.Form
        $obj.Text = $title
        $obj.Size = New-Object System.Drawing.Size($width, $height)
        $obj.AutoSize = $true
        $obj.Backcolor = $backColor
        $obj.Padding = 10
        $obj.StartPosition = "CenterScreen"
        $obj.FormBorderStyle = "FixedSingle"
        $obj.MaximizeBox = $false
        $obj.AutoScaleDimensions = New-Object System.Drawing.SizeF @([double] 8, [double] 17)
        $obj.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
        $obj.PerformAutoScale()
        return $obj
    }
    
    
    
    Function showPhoto($sPhotoPath, $width) {
        # load image
        $fs = New-Object System.IO.FileStream($sPhotoPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
        $img = [System.Drawing.Image]::FromStream($fs)
        $fs.Close()

        # calculate height of PictureBox
        $height = ($img.Size.Height*$width/$img.Size.Width)

        # form
        $objForm = makeForm -title "" -width ($width+35) -height ($height+90) -backColor "#aaaaaa"

        # PictureBox
        $objPictureBox = makePictureBox -form $objForm -image $img -posX 15 -posY 12 -width $width -height $height

        # BUTTON: save photo
        $objButtonSavePhoto = makeButton -form $objForm -text "Speichern" -posX 15 -posY ($height+20) -width 95 -height 30
        $objButtonSavePhoto.Add_Click({
            # show save file dialog
            $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
            $saveFileDialog.InitialDirectory = "."
            $saveFileDialog.Filter = "Fotos (*.jpg)| *.jpg"
            $saveFileDialog.ShowDialog() | Out-Null
            $sFilePath = $saveFileDialog.FileName

            # copy temporary image
            Copy-Item -Path $sPhotoPath -Destination (Split-Path $sFilePath -Parent)
            Rename-Item -Path "$(Split-Path $sFilePath -Parent)\$((Split-Path $sPhotoPath -Leaf))" -NewName (Split-Path $sFilePath -Leaf)
        })
 
        # finally show dialog
        [void] $objForm.ShowDialog()
    }



    Function sortListView($listview, $column) {
        $arToBeSorted = New-Object System.Collections.ArrayList

        Write-Warning $listview.Items[0].SubItems[$column].Text

        # fill array
        ForEach($item in $listview.Items) {
            $obj = New-Object PSObject
            $obj | Add-Member -type "NoteProperty" -Name "Text" -Value $item.SubItems[$column].Text
            $obj | Add-Member -type "NoteProperty" -Name "ListItem" -Value $item
            $arToBeSorted.Add($obj)
        }
    
        # sort array
        $arToBeSorted = ($arToBeSorted | Sort-Object -Property Text)
        $arToBeSorted
    
        # re-build listview
        $listview.BeginUpdate()
        $listview.Items.Clear()
        ForEach($item in $arToBeSorted) { $listview.Items.Add($item.ListItem) }
        $listview.EndUpdate()
    }



    Function guiAddUser() {
        # form
        $objForm = makeForm -title "Benutzer hinzufügen" -width 280 -height 445 -backColor "#f6f6f6"

        # forename
        makeLabel -form $objForm -text "Vorname" -posX 15 -posY 12 -width 120 -height 15
        $objInputForename = makeInput -form $objForm -posX 15 -posY 30 -width 120 -height 18

        # lastname
        makeLabel -form $objForm -text "Nachname" -posX 140 -posY 12 -width 120 -height 15
        $objInputLastname = makeInput -form $objForm -posX 140 -posY 30 -width 120 -height 18

        # street
        makeLabel -form $objForm -text "Straße + Hausnr." -posX 15 -posY 60 -width 245 -height 15
        $objInputStreet = makeInput -form $objForm -posX 15 -posY 78 -width 245 -height 18

        # zip
        makeLabel -form $objForm -text "PLZ" -posX 15 -posY 108 -width 50 -height 15
        $objInputZip = makeInput -form $objForm -posX 15 -posY 126 -width 50 -height 18

        # city
        makeLabel -form $objForm -text "Stadt" -posX 80 -posY 108 -width 120 -height 15
        $objInputCity = makeInput -form $objForm -posX 80 -posY 126 -width 180 -height 18

        # mail_address
        makeLabel -form $objForm -text "Email-Adresse" -posX 15 -posY 165 -width 245 -height 15
        $objInputMailAddress = makeInput -form $objForm -posX 15 -posY 183 -width 245 -height 18

        # mail_server
        makeLabel -form $objForm -text "Mail Server" -posX 15 -posY 213 -width 195 -height 15
        $objInputMailServer = makeInput -form $objForm -posX 15 -posY 231 -width 195 -height 18

        # mail_server_port
        makeLabel -form $objForm -text "Port" -posX 220 -posY 213 -width 40 -height 15
        $objInputMailServerPort = makeInput -form $objForm -posX 220 -posY 231 -width 40 -height 18

        # mail_user
        makeLabel -form $objForm -text "Benutzername" -posX 15 -posY 261 -width 245 -height 15
        $objInputMailUser = makeInput -form $objForm -posX 15 -posY 279 -width 245 -height 18

        # mail_password
        makeLabel -form $objForm -text "Passwort" -posX 15 -posY 309 -width 245 -height 15
        $objInputMailPassword = makeInput -form $objForm -posX 15 -posY 327 -width 245 -height 18

        # BUTTON: add user
        $objButtonSave = makeButton -form $objForm -text "Benutzer anlegen" -posX 75 -posY 370 -width 125 -height 22
        $objButtonSave.Add_Click({
            addUser -sForename $objInputForename.Text -sLastname $objInputLastname.Text -sStreet $objInputStreet.Text -sZip $objInputZip.Text -sCity $objInputCity.Text -sMailAddress $objInputMailAddress.Text -sMailServer $objInputMailServer.Text -sMailServerPort $objInputMailServerPort.Text -sMailUser $objInputMailUser.Text -sMailPassword $objInputMailPassword.Text
            $objForm.Close()
        })

        # finally show dialog
        [void] $objForm.ShowDialog()
    }


    
    Function guiEditUser($user) {
        # form
        $objForm = makeForm -title "Benutzer bearbeiten" -width 280 -height 485 -backColor "#f6f6f6"

        # forename
        makeLabel -form $objForm -text "Vorname" -posX 15 -posY 12 -width 120 -height 15
        $objInputForename = makeInput -form $objForm -posX 15 -posY 30 -width 120 -height 18
        $objInputForename.Text = $user.forename

        # lastname
        makeLabel -form $objForm -text "Nachname" -posX 140 -posY 12 -width 120 -height 15
        $objInputLastname = makeInput -form $objForm -posX 140 -posY 30 -width 120 -height 18
        $objInputLastname.Text = $user.lastname

        # street
        makeLabel -form $objForm -text "Straße + Hausnr." -posX 15 -posY 60 -width 245 -height 15
        $objInputStreet = makeInput -form $objForm -posX 15 -posY 78 -width 245 -height 18
        $objInputStreet.Text = $user.street

        # zip
        makeLabel -form $objForm -text "PLZ" -posX 15 -posY 108 -width 50 -height 15
        $objInputZip = makeInput -form $objForm -posX 15 -posY 126 -width 50 -height 18
        $objInputZip.Text = $user.zip

        # city
        makeLabel -form $objForm -text "Stadt" -posX 80 -posY 108 -width 120 -height 15
        $objInputCity = makeInput -form $objForm -posX 80 -posY 126 -width 180 -height 18
        $objInputCity.Text = $user.city

        # mail_address
        makeLabel -form $objForm -text "Email-Adresse" -posX 15 -posY 165 -width 245 -height 15
        $objInputMailAddress = makeInput -form $objForm -posX 15 -posY 183 -width 245 -height 18
        $objInputMailAddress.Text = $user.mail_address

        # mail_server
        makeLabel -form $objForm -text "Mail Server" -posX 15 -posY 213 -width 195 -height 15
        $objInputMailServer = makeInput -form $objForm -posX 15 -posY 231 -width 195 -height 18
        $objInputMailServer.Text = $user.mail_server

        # mail_server_port
        makeLabel -form $objForm -text "Port" -posX 220 -posY 213 -width 40 -height 15
        $objInputMailServerPort = makeInput -form $objForm -posX 220 -posY 231 -width 40 -height 18
        $objInputMailServerPort.Text = $user.mail_server_port

        # BUTTON: save edits
        $objButtonSave = makeButton -form $objForm -text "Speichern" -posX 75 -posY 270 -width 125 -height 22
        $objButtonSave.Add_Click({
            editUser -id $user.id -sForename $objInputForename.Text -sLastname $objInputLastname.Text -sStreet $objInputStreet.Text -sZip $objInputZip.Text -sCity $objInputCity.Text -sMailAddress $objInputMailAddress.Text -sMailServer $objInputMailServer.Text -sMailServerPort $objInputMailServerPort.Text
            $objForm.Close()
        })


        # mail_user
        makeLabel -form $objForm -text "Benutzername" -posX 15 -posY 320 -width 245 -height 15
        $objInputMailUser = makeInput -form $objForm -posX 15 -posY 338 -width 245 -height 18
        $objInputMailUser.Text = $user.mail_user

        # mail_password
        makeLabel -form $objForm -text "Passwort" -posX 15 -posY 368 -width 245 -height 15
        $objInputMailPassword = makeInput -form $objForm -posX 15 -posY 386 -width 245 -height 18

        # BUTTON: save credentials
        $objButtonSaveCredentials = makeButton -form $objForm -text "Speichern" -posX 75 -posY 425 -width 125 -height 22
        $objButtonSaveCredentials.Add_Click({
            editUserCredentials -id $user.id -sMailUser $objInputMailUser.Text -sMailPassword $objInputMailPassword.Text
            $objForm.Close()
        })

        # finally show dialog
        [void] $objForm.ShowDialog()
    }



    Function guiAddIncident($objComboUser) {
        # form
        $objForm = makeForm -title "Vorfall hinzufügen" -width 390 -height 245 -backColor "#f6f6f6"

        # location
        makeLabel -form $objForm -text "Ort" -posX 15 -posY 12 -width 235 -height 15
        $objInputLocation = makeInput -form $objForm -posX 15 -posY 30 -width 235 -height 18

        # path to photo
        makeLabel -form $objForm -text "Foto" -posX 15 -posY 60 -width 235 -height 15
        $objInputPhoto = makeInput -form $objForm -posX 15 -posY 78 -width 235 -height 18

        # BUTTON: load photo
        $objButtonLoadPhoto = makeButton -form $objForm -text "Durchsuchen..." -posX 265 -posY 77 -width 95 -height 22
        $objButtonLoadPhoto.Add_Click({
            $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $openFileDialog.InitialDirectory = "."
            $openFileDialog.Filter = "Fotos (*.jpg)| *.jpg"
            $openFileDialog.Multiselect = $false
            $openFileDialog.ShowDialog() | Out-Null
            $objInputPhoto.Text = $openFileDialog.FileName
        })

        # car license tag
        makeLabel -form $objForm -text "Kennzeichen" -posX 15 -posY 108 -width 75 -height 15
        $objInputLicenseTag = makeInput -form $objForm -posX 15 -posY 126 -width 75 -height 18
        
        # car model
        makeLabel -form $objForm -text "Modell" -posX 100 -posY 108 -width 175 -height 15
        $objInputCarModel = makeInput -form $objForm -posX 100 -posY 126 -width 175 -height 18

        # car color
        makeLabel -form $objForm -text "Farbe" -posX 285 -posY 108 -width 75 -height 15
        $objInputCarColor = makeInput -form $objForm -posX 285 -posY 126 -width 75 -height 18

        # BUTTON: save
        $objButtonLoadPhoto = makeButton -form $objForm -text "Speichern" -posX 135 -posY 165 -width 110 -height 35
        $objButtonLoadPhoto.Add_Click({
            # set mouse cursor to waiting icon
            $objForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

            # fetch user
            $user = fetchUser -sMailAddress $objComboUser.SelectedItem

            # add incident to database
            addIncident -userId $user.id -sPathToPhoto $objInputPhoto.Text -sLocation $objInputLocation.Text -sCarLicenseTag $objInputLicenseTag.Text -sCarModel $objInputCarModel.Text -sCarColor $objInputCarColor.Text

            # reset mouse cursor
            $objForm.Cursor = [System.Windows.Forms.Cursors]::Default

            # close this form
            $objForm.Close()
        })

        # finally show dialog
        [void] $objForm.ShowDialog()
    }



    Function guiMain_pullDataFromDb($objComboUser, $objListViewUnsent, $objListViewSent) {
        # fetch user and incidents from DB
        $user = fetchUser -sMailAddress $objComboUser.SelectedItem
        $incidentsUnsent = fetchIncidentsByUserUnsent -userId $user.id
        $incidentsSent = fetchIncidentsByUserSent -userId $user.id

        # add unsent incidents
        $objListViewUnsent.BeginUpdate()
        $objListViewUnsent.Items.Clear()
        ForEach($incident in $incidentsUnsent) {
            $item = New-Object System.Windows.Forms.ListViewItem
            $item.Name = $incident.id
            $item.Text = $incident.id
            $item.SubItems.Add(($incident.timestamp).ToString()) | Out-Null
            $item.SubItems.Add($incident.location) | Out-Null
            $item.SubItems.Add("$($incident.location_lat), $($incident.location_lon)") | Out-Null
            $item.SubItems.Add("$($incident.car_model) ($($incident.car_color))") | Out-Null
            $item.SubItems.Add($incident.car_license_tag) | Out-Null
            $objListViewUnsent.Items.Add($item) | Out-Null
        }
        $objListViewUnsent.EndUpdate()

        # add sent incidents
        $objListViewSent.BeginUpdate()
        $objListViewSent.Items.Clear()
        ForEach($incident in $incidentsSent) {
            $item = New-Object System.Windows.Forms.ListViewItem
            $item.Name = $incident.id
            $item.Text = $incident.id
            $item.SubItems.Add(($incident.timestamp).ToString()) | Out-Null
            $item.SubItems.Add($incident.location) | Out-Null
            $item.SubItems.Add("$($incident.location_lat), $($incident.location_lon)") | Out-Null
            $item.SubItems.Add("$($incident.car_model) ($($incident.car_color))") | Out-Null
            $item.SubItems.Add($incident.car_license_tag) | Out-Null
            $item.SubItems.Add(($incident.timestamp_sent).ToString()) | Out-Null
            $objListViewSent.Items.Add($item) | Out-Null
        }
        $objListViewSent.EndUpdate()
    }



    Function guiMain() {
        # form
        $objForm = makeForm -title "FalschparkmelderFFM" -width 875 -height 610 -backColor "#f6f6f6"

        # ---------------------------------------------
        #  USER AREA
        # ---------------------------------------------

        # User Selection
        makeLabel -form $objForm -text "Benutzer:" -posX 15 -posY 12 -width 55 -height 15
        $objComboUser = makeCombo -form $objForm -posX 70 -posY 9 -width 195 -height 15 -height2 70
        $objComboUser.Add_SelectedIndexChanged({
            guiMain_pullDataFromDb -objComboUser $objComboUser -objListViewUnsent $objListViewUnsent -objListViewSent $objListViewSent
        })

        
        # BUTTON: User Edit
        $objButtonUserEdit = makeButton -form $objForm -text "Bearbeiten" -posX 275 -posY 8 -width 75 -height 23
        $objButtonUserEdit.Add_Click({
            # add user to DB
            guiEditUser -user (fetchUser -sMailAddress $objComboUser.SelectedItem)

            # fetch users from DB
            $users = fetchUsers

            # fill combo box
            $objComboUser.Items.Clear()
            ForEach($user in $users) { $objComboUser.Items.Add($user.mail_address) }
            $objComboUser.SelectedIndex = 0
        })

        # BUTTON: User Deletion
        $objButtonUserDelete = makeButton -form $objForm -text "Löschen" -posX 355 -posY 8 -width 75 -height 23
        $objButtonUserDelete.Add_Click({
            # fetch user
            $user = fetchUser -sMailAddress $objComboUser.SelectedItem

            # delete user from DB
            $deleted = $true
            If (removeUser -id $user.id -stopIfUnsentIncidents $true) {
                
                # ask if user should be deleted although there are unsent incidents
                If ([System.Windows.Forms.MessageBox]::Show("Dieser Benutzer hat noch ungesendete Vorfälle. Wirklich löschen?","Ungesendete Vorfälle", [System.Windows.Forms.MessageBoxButtons]::YesNo) -eq "Yes") {
                    removeUser -id $user.id -stopIfUnsentIncidents $false
                } Else {
                    $deleted = $false
                }
            }

            If ($deleted) {
                # remove entry from list
                $objComboUser.Items.Remove($objComboUser.SelectedItem)

                # refresh list views
                guiMain_pullDataFromDb -objComboUser $objComboUser -objListViewUnsent $objListViewUnsent -objListViewSent $objListViewSent
            }
        })

        # BUTTON: User Add
        $objButtonUserAdd = makeButton -form $objForm -text "Hinzufügen" -posX 435 -posY 8 -width 75 -height 23
        $objButtonUserAdd.Add_Click({
            # add user to DB
            guiAddUser

            # fetch users from DB
            $users = fetchUsers

            # fill combo box
            $objComboUser.Items.Clear()
            ForEach($user in $users) { $objComboUser.Items.Add($user.mail_address) }
            $objComboUser.SelectedIndex = 0
        })

        
        # fetch user and incidents from DB
        $user = fetchUser -sMailAddress $objComboUser.SelectedItem
        $incidentsUnsent = fetchIncidentsByUserUnsent -userId $user.id
        $incidentsUnsent
        $incidentsSent = fetchIncidentsByUserSent -userId $user.id


    
        # ---------------------------------------------
        #  INCIDENT AREA
        # ---------------------------------------------

        # ---------------------------------------------
        # Unsent Incidents
        # ---------------------------------------------
        makeLabel -form $objForm -text "Ungesendete Vorfälle" -posX 15 -posY 50 -width 125 -height 15
        $objListViewUnsent = makeListView -form $objForm -posX 15 -posY 70 -width 750 -height 150
        $objListViewUnsent.Add_ColumnClick({sortListView -listview $objListViewUnsent -column $_.Column})
        $objListViewUnsent.Add_DoubleClick({
            # fetch chosen incident
            $incident_id = $objListViewUnsent.SelectedItems[0].Name
            
            # fetch incident
            $incident = fetchIncident -id $incident_id

            # save photo temporarily
            $tmpFile = New-TemporaryFile
            saveBlobToFile -blob $incident.photo -sFilePath $tmpFile.FullName

            # show temporary photo
            showPhoto -width 1024 -sPhotoPath $tmpFile.FullName

            # delete temporary photo
            Remove-Item $tmpFile.FullName
        })

        # add columns
        ("ID", "Datum", "Ort", "Koordinaten", "Fahrzeug", "Kennzeichen") | % {
            $column = New-Object System.Windows.Forms.ColumnHeader
            $column.Width = -2
            $column.Text = $_
            $objListViewUnsent.Columns.Add($column)
        }

        # BUTTON: Incident Add
        $objButtonIncidentAdd = makeButton -form $objForm -text "Hinzufügen" -posX 785 -posY 75 -width 85 -height 28
        $objButtonIncidentAdd.Add_Click({
            # add incident
            guiAddIncident -objComboUser $objComboUser

            # reload list views
            guiMain_pullDataFromDb -objComboUser $objComboUser -objListViewUnsent $objListViewUnsent -objListViewSent $objListViewSent
        })

        # BUTTON: Incident Deletion
        $objButtonIncidentDelete = makeButton -form $objForm -text "Löschen" -posX 785 -posY 110 -width 85 -height 28
        $objButtonIncidentDelete.Add_Click({
            # delete incident
            removeIncident -id $objListViewUnsent.SelectedItems[0].Name

            # reload list views
            guiMain_pullDataFromDb -objComboUser $objComboUser -objListViewUnsent $objListViewUnsent -objListViewSent $objListViewSent
        })

        # BUTTON: Incident Send
        $objButtonIncidentSend = makeButton -form $objForm -text "Senden" -posX 785 -posY 145 -width 85 -height 28
        $objButtonIncidentSend.Add_Click({
            # set mouse cursor to waiting icon
            $objForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

            # iterate over selected item(s)
            ForEach($item in $objListViewUnsent.SelectedItems) {
                # send incidents
                $result = sendIncident -id $item.Name
            
                # move list item from unsent to sent
                If ($result) {
					$objListViewUnsent.Items.Remove($item)
					$item.SubItems.Add("vor Kurzem")
					$objListViewSent.Items.Add($item)
				} Else {
					[System.Windows.Forms.MessageBox]::Show("Email konnte nicht gesendet werden!","Mail Error", [System.Windows.Forms.MessageBoxButtons]::Ok)
				}
            }

            # reset mouse cursor
            $objForm.Cursor = [System.Windows.Forms.Cursors]::Default

        })


        # ---------------------------------------------
        # Sent Incidents
        # ---------------------------------------------
        makeLabel -form $objForm -text "Gesendete Vorfälle" -posX 15 -posY 235 -width 125 -height 15
        $objListViewSent = makeListView -form $objForm -posX 15 -posY 255 -width 750 -height 300
        $objListViewSent.Add_ColumnClick({sortListView -listview $objListViewSent -column $_.Column})
        $objListViewSent.Add_DoubleClick({
            # fetch chosen incident
            $incident_id = $objListViewSent.SelectedItems[0].Name
            
            # fetch incident
            $incident = fetchIncident -id $incident_id

            # save photo temporarily
            $tmpFile = New-TemporaryFile
            saveBlobToFile -blob $incident.photo -sFilePath $tmpFile.FullName

            # show temporary photo
            showPhoto -width 1024 -sPhotoPath $tmpFile.FullName

            # delete temporary photo
            Remove-Item $tmpFile.FullName
        })

        # add columns
        ("ID", "Datum", "Ort", "Koordinaten", "Fahrzeug", "Kennzeichen", "Gemeldet") | % {
            $column = New-Object System.Windows.Forms.ColumnHeader
            $column.Width = -2
            $column.Text = $_
            $objListViewSent.Columns.Add($column)
        }

        # BUTTON: Export Photos
        $objButtonExportPhoto = makeButton -form $objForm -text "Foto(s)`nexportieren" -posX 785 -posY 255 -width 85 -height 40
        $objButtonExportPhoto.Add_Click({

            # ask user for export folder            
            $folderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
            $folderBrowserDialog.ShowNewFolderButton = $true
            $folderBrowserDialog.Description = "Bitte Export Ordner wählen"
            $folderBrowserDialog.ShowDialog() | Out-Null
            $sFolderPath = $folderBrowserDialog.SelectedPath

            # set mouse cursor to waiting icon
            $objForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

            # export photos
            ForEach($item in $objListViewSent.SelectedItems) {
                # load incident from DB
                $incident = fetchIncident -id $item.SubItems[0].Text
                
                # save incident
                $sFileName = "$(($incident.timestamp).ToString("yyyyMMdd-HHmm")).jpg"
                saveBlobToFile -blob $incident.photo -sFilePath "$($sFolderPath)\$sFileName"
            }

            # reset mouse cursor
            $objForm.Cursor = [System.Windows.Forms.Cursors]::Default
        })

        # BUTTON: Export GPX
        $objButtonExportGPX = makeButton -form $objForm -text "Koordinaten`nexportieren" -posX 785 -posY 305 -width 85 -height 40
        $objButtonExportGPX.Add_Click({
            # load selected incidents from DB
            $incidents = @()
            ForEach($item in $objListViewSent.SelectedItems) { $incidents += fetchIncident -id $item.SubItems[0].Text }
            
            # export incidents to GPX
            exportGPX -incidents $incidents
        })


        # ---------------------------------------------
        # load data
        # ---------------------------------------------
        
        # fetch all users from DB
        $users = fetchUsers

        # fill combo box
        $objComboUser.Items.Clear()
        ForEach($user in $users) { $objComboUser.Items.Add($user.mail_address) }
        $objComboUser.SelectedIndex = 0

        # fill list views
        guiMain_pullDataFromDb -objComboUser $objComboUser -objListViewUnsent $objListViewUnsent -objListViewSent $objListViewSent



        # ---------------------------------------------
        # finally show dialog
        # ---------------------------------------------
        $objForm.ShowDialog() | Out-Null
    }





# ------------------------------------------------------------------------------------------------------------
#  > MAIN
# ------------------------------------------------------------------------------------------------------------

    # open DB connection
    openDbConnection
    
    # create Tables if not existing
    createTablesIfNotExisting

    # check if at least 1 user is setup
    If (-not (fetchUsers)) { guiAddUser }

    # show GUI
    guiMain

    # close DB connection
    closeDbConnection
