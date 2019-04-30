$weekdayMap = [System.Collections.Generic.Dictionary[string, int]]::new()
$weekdayMap.Add("monday", 1)
$weekdayMap.Add("tuesday", 2)
$weekdayMap.Add("wednesday", 3)
$weekdayMap.Add("thursday", 4)
$weekdayMap.Add("friday", 5)
$weekdayMap.Add("saturday", 6)
$weekdayMap.Add("sunday", 7)

[string]$title = "Monday"
[datetime]$creationDate = (Get-Date).Date

$originalAssignmentDate = $null
if ($weekdayMap.ContainsKey($title.ToLower())) {
    $originalAssignmentDate = $creationDate
    while ($originalAssignmentDate.DayOfWeek.ToString().ToLower() -ne $title.ToLower()) {
        $originalAssignmentDate = $originalAssignmentDate.AddDays(1)
    }
}
$originalAssignmentDate