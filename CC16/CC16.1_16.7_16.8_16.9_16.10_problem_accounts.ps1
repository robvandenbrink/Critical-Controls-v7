$probusers = get-aduser -filter * -properties * | select samaccountname, name, enabled, scriptpath, passwordlastset, passwordexpired, passwordneverexpires, passwordnotrequired, lockedout, lastlogon, lastlogondate, lastlogontimestamp, lockedout, cannotchangepassword, accountexpirationdate, mobilephone, officephone, telephoneNumber 
$probusers| export-csv accountinfo.csv

