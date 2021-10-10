$apiKey = "APIKEY"
$City = "City"
$uri = "api.openweathermap.org/data/2.5/weather?q=$City&appid=$apikey"
$response = Invoke-RestMethod -Uri $uri
$pressure = $response.main.pressure
$humid = $response.main.humidity
$Clouds = $response.weather.description
$temp = [int]$response.main.temp - 273
"It is currently $temp degres outside with $clouds. The current pressure is $pressure and the humidity is $humid"
