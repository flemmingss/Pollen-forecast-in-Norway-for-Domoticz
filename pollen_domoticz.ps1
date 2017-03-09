# Pollen Forecast In Norway for Domoticz | https://github.com/flemmingss/

##############################################################################################

### configuration ###

# Html Agility Pack (Download from https://htmlagilitypack.codeplex.com/)
    $HtmlAgilityPack_folder = "C:\HtmlAgilityPack.1.4.6\Net45"
    $HtmlAgilityPack_file = "HtmlAgilityPack.dll"

# Location, options: "Østlandet med Oslo", "Sørlandet", "Rogaland", "Hordaland", "Sogn og Fjordane", "Møre og Romsdal", "Indre Østlandet", "Sentrale fjellstrøk i Sør-Norge", "Trøndelag", "Nordland", "Troms", "Finnmark"
       $location = "Sørlandet"

# Select active categories, "1" = enabled, "0" = disabled.
	$or_today = "0"
	$or_tomorrow = "0"
	$hassel_today = "0"
	$hassel_tomorrow = "0"
	$salix_today = "0"
	$salix_tomorrow = "0"
	$bjork_today = "0"
	$bjork_tomorrow = "0"
	$gress_today = "0"
	$gress_tomorrow = "0"
	$burot_today = "0"
	$burot_tomorrow = "0"

# IDX of dummy devices in Domoticz, only needed for enabled categories
	$or_today_idx = "0"
	$or_tomorrow_idx = "0"
	$hassel_today_idx = "0"
	$hassel_tomorrow_idx = "0"
	$salix_today_idx = "0"
	$salix_tomorrow_idx = "0"
	$bjork_today_idx = "0"
	$bjork_tomorrow_idx = "0"
	$gress_today_idx = "0"
	$gress_tomorrow_idx = "0"
	$burot_today_idx = "0"
	$burot_tomorrow_idx = "0"

# Domoticz Credentials + Address
    $user = "x" #Domoticz username
    $pass = "x" #Domoticz password
    $domoticz_address = "http://192.168.1.123:8080/" #Domoticz Address


### End of configuration ###

##############################################################################################

### Web scraping ###

cd $HtmlAgilityPack_folder
add-type -Path .\$HtmlAgilityPack_file
$web = "http://www.pollenvarslingen.no/Forsiden/Varsel.aspx"
$doc = New-Object HtmlAgilityPack.HtmlDocument
$doc.LoadHtml((iwr $web).RawContent)
 
$output = $doc.DocumentNode.SelectNodes("//*[@id='forecastTable']//tr[contains(.,'$location')]")

### Results to variables ###

if ($or_today -eq "1")
    {
        $or_today_textvalue = $output.SelectNodes("./td[2]").Attributes["title"].value
    if ($or_today_textvalue -eq "Ingen spredning"){$or_today_numbervalue = "0"}
    if ($or_today_textvalue -eq "Beskjeden spredning"){$or_today_numbervalue = "1"}
    if ($or_today_textvalue -eq "Moderat spredning"){$or_today_numbervalue = "2"}
    if ($or_today_textvalue -eq "Kraftig spredning"){$or_today_numbervalue = "3"}
    if ($or_today_textvalue -eq "Ekstrem spredning"){$or_today_numbervalue = "4"}
    }

if ($or_tomorrow -eq "1")
    {
        $or_tomorrow_textvalue = $output.SelectNodes("./td[3]").Attributes["title"].value
    if ($or_tomorrow_textvalue -eq "Ingen spredning"){$or_tomorrow_numbervalue = "0"}
    if ($or_tomorrow_textvalue -eq "Beskjeden spredning"){$or_tomorrow_numbervalue = "1"}
    if ($or_tomorrow_textvalue -eq "Moderat spredning"){$or_tomorrow_numbervalue = "2"}
    if ($or_tomorrow_textvalue -eq "Kraftig spredning"){$or_tomorrow_numbervalue = "3"}
    if ($or_tomorrow_textvalue -eq "Ekstrem spredning"){$or_tomorrow_numbervalue = "4"}
    }

if ($hassel_today -eq "1")
    {
        $hassel_today_textvalue = $output.SelectNodes("./td[4]").Attributes["title"].value
    if ($hassel_today_textvalue -eq "Ingen spredning"){$hassel_today_numbervalue = "0"}
    if ($hassel_today_textvalue -eq "Beskjeden spredning"){$hassel_today_numbervalue = "1"}
    if ($hassel_today_textvalue -eq "Moderat spredning"){$hassel_today_numbervalue = "2"}
    if ($hassel_today_textvalue -eq "Kraftig spredning"){$hassel_today_numbervalue = "3"}
    if ($hassel_today_textvalue -eq "Ekstrem spredning"){$hassel_today_numbervalue = "4"}
    }

if ($hassel_tomorrow -eq "1")
    {
        $hassel_tomorrow_textvalue = $output.SelectNodes("./td[5]").Attributes["title"].value
    if ($hassel_tomorrow_textvalue -eq "Ingen spredning"){$hassel_tomorrow_numbervalue = "0"}
    if ($hassel_tomorrow_textvalue -eq "Beskjeden spredning"){$hassel_tomorrow_numbervalue = "1"}
    if ($hassel_tomorrow_textvalue -eq "Moderat spredning"){$hassel_tomorrow_numbervalue = "2"}
    if ($hassel_tomorrow_textvalue -eq "Kraftig spredning"){$hassel_tomorrow_numbervalue = "3"}
    if ($hassel_tomorrow_textvalue -eq "Ekstrem spredning"){$hassel_tomorrow_numbervalue = "4"}
    }

if ($salix_today -eq "1")
    {
        $salix_today_textvalue = $output.SelectNodes("./td[6]").Attributes["title"].value
    if ($salix_today_textvalue -eq "Ingen spredning"){$salix_today_numbervalue = "0"}
    if ($salix_today_textvalue -eq "Beskjeden spredning"){$salix_today_numbervalue = "1"}
    if ($salix_today_textvalue -eq "Moderat spredning"){$salix_today_numbervalue = "2"}
    if ($salix_today_textvalue -eq "Kraftig spredning"){$salix_today_numbervalue = "3"}
    if ($salix_today_textvalue -eq "Ekstrem spredning"){$salix_today_numbervalue = "4"}
    }

if ($salix_tomorrow -eq "1")
    {
        $salix_tomorrow_textvalue = $output.SelectNodes("./td[7]").Attributes["title"].value
    if ($salix_tomorrow_textvalue -eq "Ingen spredning"){$salix_tomorrow_numbervalue = "0"}
    if ($salix_tomorrow_textvalue -eq "Beskjeden spredning"){$salix_tomorrow_numbervalue = "1"}
    if ($salix_tomorrow_textvalue -eq "Moderat spredning"){$salix_tomorrow_numbervalue = "2"}
    if ($salix_tomorrow_textvalue -eq "Kraftig spredning"){$salix_tomorrow_numbervalue = "3"}
    if ($salix_tomorrow_textvalue -eq "Ekstrem spredning"){$salix_tomorrow_numbervalue = "4"}
    }

if ($bjork_today -eq "1")
    {
        $bjork_today_textvalue = $output.SelectNodes("./td[8]").Attributes["title"].value
    if ($bjork_today_textvalue -eq "Ingen spredning"){$bjork_today_numbervalue = "0"}
    if ($bjork_today_textvalue -eq "Beskjeden spredning"){$bjork_today_numbervalue = "1"}
    if ($bjork_today_textvalue -eq "Moderat spredning"){$bjork_today_numbervalue = "2"}
    if ($bjork_today_textvalue -eq "Kraftig spredning"){$bjork_today_numbervalue = "3"}
    if ($bjork_today_textvalue -eq "Ekstrem spredning"){$bjork_today_numbervalue = "4"}
    }

if ($bjork_tomorrow -eq "1")
    {
        $bjork_tomorrow_textvalue = $output.SelectNodes("./td[9]").Attributes["title"].value
    if ($bjork_tomorrow_textvalue -eq "Ingen spredning"){$bjork_tomorrow_numbervalue = "0"}
    if ($bjork_tomorrow_textvalue -eq "Beskjeden spredning"){$bjork_tomorrow_numbervalue = "1"}
    if ($bjork_tomorrow_textvalue -eq "Moderat spredning"){$bjork_tomorrow_numbervalue = "2"}
    if ($bjork_tomorrow_textvalue -eq "Kraftig spredning"){$bjork_tomorrow_numbervalue = "3"}
    if ($bjork_tomorrow_textvalue -eq "Ekstrem spredning"){$bjork_tomorrow_numbervalue = "4"}
    }

if ($gress_today -eq "1")
    {
        $gress_today_textvalue = $output.SelectNodes("./td[10]").Attributes["title"].value
    if ($gress_today_textvalue -eq "Ingen spredning"){$gress_today_numbervalue = "0"}
    if ($gress_today_textvalue -eq "Beskjeden spredning"){$gress_today_numbervalue = "1"}
    if ($gress_today_textvalue -eq "Moderat spredning"){$gress_today_numbervalue = "2"}
    if ($gress_today_textvalue -eq "Kraftig spredning"){$gress_today_numbervalue = "3"}
    if ($gress_today_textvalue -eq "Ekstrem spredning"){$gress_today_numbervalue = "4"}
    }

if ($gress_tomorrow -eq "1")
    {
        $gress_tomorrow_textvalue = $output.SelectNodes("./td[11]").Attributes["title"].value
    if ($gress_tomorrow_textvalue -eq "Ingen spredning"){$gress_tomorrow_numbervalue = "0"}
    if ($gress_tomorrow_textvalue -eq "Beskjeden spredning"){$gress_tomorrow_numbervalue = "1"}
    if ($gress_tomorrow_textvalue -eq "Moderat spredning"){$gress_tomorrow_numbervalue = "2"}
    if ($gress_tomorrow_textvalue -eq "Kraftig spredning"){$gress_tomorrow_numbervalue = "3"}
    if ($gress_tomorrow_textvalue -eq "Ekstrem spredning"){$gress_tomorrow_numbervalue = "4"}
    }

if ($burot_today -eq "1")
    {
        $burot_today_textvalue = $output.SelectNodes("./td[12]").Attributes["title"].value
    if ($burot_today_textvalue -eq "Ingen spredning"){$burot_today_numbervalue = "0"}
    if ($burot_today_textvalue -eq "Beskjeden spredning"){$burot_today_numbervalue = "1"}
    if ($burot_today_textvalue -eq "Moderat spredning"){$burot_today_numbervalue = "2"}
    if ($burot_today_textvalue -eq "Kraftig spredning"){$burot_today_numbervalue = "3"}
    if ($burot_today_textvalue -eq "Ekstrem spredning"){$burot_today_numbervalue = "4"}
    }

if ($burot_tomorrow -eq "1")
    {
        $burot_tomorrow_textvalue = $output.SelectNodes("./td[13]").Attributes["title"].value
    if ($burot_tomorrow_textvalue -eq "Ingen spredning"){$burot_tomorrow_numbervalue = "0"}
    if ($burot_tomorrow_textvalue -eq "Beskjeden spredning"){$burot_tomorrow_numbervalue = "1"}
    if ($burot_tomorrow_textvalue -eq "Moderat spredning"){$burot_tomorrow_numbervalue = "2"}
    if ($burot_tomorrow_textvalue -eq "Kraftig spredning"){$burot_tomorrow_numbervalue = "3"}
    if ($burot_tomorrow_textvalue -eq "Ekstrem spredning"){$burot_tomorrow_numbervalue = "4"}
    }

### Domoticz Authorization ###

$pair = "${user}:${pass}"
$bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
$base64 = [System.Convert]::ToBase64String($bytes)
$basicAuthValue = "Basic $base64"
$headers = @{ Authorization = $basicAuthValue }


### JSON ###

if ($or_today -eq "1")
	{
	$idx = $or_today_idx
	$value = $or_today_numbervalue
	$domoticz_json = "json.htm?type=command&param=udevice&idx=$idx&nvalue=0&svalue=$value"
	Invoke-RestMethod -Method Get -Uri "$domoticz_address$domoticz_json" -Headers $headers
	}

if ($or_tomorrow -eq "1")
	{
	$idx = $or_tomorrow_idx
	$value = $or_tomorrow_numbervalue
	$domoticz_json = "json.htm?type=command&param=udevice&idx=$idx&nvalue=0&svalue=$value"
	Invoke-RestMethod -Method Get -Uri "$domoticz_address$domoticz_json" -Headers $headers
	}
	
if ($hassel_today -eq "1")
	{
	$idx = $hassel_today_idx
	$value = $hassel_today_numbervalue
	$domoticz_json = "json.htm?type=command&param=udevice&idx=$idx&nvalue=0&svalue=$value"
	Invoke-RestMethod -Method Get -Uri "$domoticz_address$domoticz_json" -Headers $headers
	}

if ($hassel_tomorrow -eq "1")
	{
	$idx = $hassel_tomorrow_idx
	$value = $hassel_tomorrow_numbervalue
	$domoticz_json = "json.htm?type=command&param=udevice&idx=$idx&nvalue=0&svalue=$value"
	Invoke-RestMethod -Method Get -Uri "$domoticz_address$domoticz_json" -Headers $headers
	}
	
if ($salix_today -eq "1")
	{
	$idx = $salix_today_idx
	$value = $salix_today_numbervalue
	$domoticz_json = "json.htm?type=command&param=udevice&idx=$idx&nvalue=0&svalue=$value"
	Invoke-RestMethod -Method Get -Uri "$domoticz_address$domoticz_json" -Headers $headers
	}

if ($salix_tomorrow -eq "1")
	{
	$idx = $salix_tomorrow_idx
	$value = $salix_tomorrow_numbervalue
	$domoticz_json = "json.htm?type=command&param=udevice&idx=$idx&nvalue=0&svalue=$value"
	Invoke-RestMethod -Method Get -Uri "$domoticz_address$domoticz_json" -Headers $headers
	}
	
if ($bjork_today -eq "1")
	{
	$idx = $bjork_today_idx
	$value = $bjork_today_numbervalue
	$domoticz_json = "json.htm?type=command&param=udevice&idx=$idx&nvalue=0&svalue=$value"
	Invoke-RestMethod -Method Get -Uri "$domoticz_address$domoticz_json" -Headers $headers
	}

if ($bjork_tomorrow -eq "1")
	{
	$idx = $bjork_tomorrow_idx
	$value = $bjork_tomorrow_numbervalue
	$domoticz_json = "json.htm?type=command&param=udevice&idx=$idx&nvalue=0&svalue=$value"
	Invoke-RestMethod -Method Get -Uri "$domoticz_address$domoticz_json" -Headers $headers
	}
	
if ($gress_today -eq "1")
	{
	$idx = $gress_today_idx
	$value = $gress_today_numbervalue
	$domoticz_json = "json.htm?type=command&param=udevice&idx=$idx&nvalue=0&svalue=$value"
	Invoke-RestMethod -Method Get -Uri "$domoticz_address$domoticz_json" -Headers $headers
	}

if ($gress_tomorrow -eq "1")
	{
	$idx = $gress_tomorrow_idx
	$value = $gress_tomorrow_numbervalue
	$domoticz_json = "json.htm?type=command&param=udevice&idx=$idx&nvalue=0&svalue=$value"
	Invoke-RestMethod -Method Get -Uri "$domoticz_address$domoticz_json" -Headers $headers
	}

if ($burot_today -eq "1")
	{
	$idx = $burot_today_idx
	$value = $burot_today_numbervalue
	$domoticz_json = "json.htm?type=command&param=udevice&idx=$idx&nvalue=0&svalue=$value"
	Invoke-RestMethod -Method Get -Uri "$domoticz_address$domoticz_json" -Headers $headers
	}

if ($burot_tomorrow -eq "1")
	{
	$idx = $burot_tomorrow_idx
	$value = $burot_tomorrow_numbervalue
	$domoticz_json = "json.htm?type=command&param=udevice&idx=$idx&nvalue=0&svalue=$value"
	Invoke-RestMethod -Method Get -Uri "$domoticz_address$domoticz_json" -Headers $headers
	}

# End