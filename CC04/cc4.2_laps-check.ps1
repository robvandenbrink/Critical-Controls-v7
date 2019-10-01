# this collects all passwords set by LAPS (Local Admin Password Solution)
# don't store these passwords as a clear-text file
# this script is mainly to verify that LAPS is deployed, and is configured
# correctly.  View passwords to ensure that they meet the LAPS desired 
# policy and are unique

$localpwds = get-adcomputer -filter -Properties * | select name, operatingsystem, ms-MCS-AdmPwd

$localpwds | out-gridview