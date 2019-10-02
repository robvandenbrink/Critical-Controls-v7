$checktime = (get-date).adddays(-30)
$newusers = get-aduser -searchbase "DC=mittenvinyl,DC=local" -Properties whencreated -filter {whencreated -ge $checktime}
$newusers | select samaccountname, name, whencreated