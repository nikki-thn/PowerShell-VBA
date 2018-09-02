$day = (get-date).Day
$month = (get-date).Month
$year = (get-date).year

$crossRefTestScript = "
select
       
  and (DATEPART(yy, trigger_dt) = $year
AND    DATEPART(mm, trigger_dt) = $month
AND    DATEPART(dd, trigger_dt) = $day)
  order by trigger_dt desc, value_ref_id;
 "

Invoke-Sqlcmd -Query $crossRefTestScript -ServerInstance "CGSD01" -Database "AMR" | Out-GridView
