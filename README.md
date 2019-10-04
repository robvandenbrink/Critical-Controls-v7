This repo outlines various scripts to implement or assess the CIS Critical Controls

Blogs that reference or document these scripts are also listed

The Sector 2019 Powerpoint outlines the use of PowerShell to Assess and/or Implement these controls in a Corporate environment.
Many of these scripts take advantage of Active Directory, but they can be easily modified to use IP addresses or ranges instead

The 20 Controls (version 7) are:

     1/ Inventory and Control of Hardware Assets
     2/ Inventory and Control of Software Assets
     3/ Continuous Vulnerability Management
     4/ Controlled Use of Administrative Privileges
     5/ Secure Configuration for Hardware and Software on Mobile Devices, Laptops, Workstations and Servers
     6/ Maintenance, Monitoring and Analysis of Audit Logs
     7/ Email and Web Browser Protections
     8/ Malware Defenses
     9/ Limitation and Control of Network Ports, Protocols, and Services
    10/ Data Recovery Capabilities
    11/ Secure Configuration for Network Devices, such as Firewalls, Routers and Switches
    12/ Boundary Defense
    13/ Data Protection
    14/ Controlled Access Based on the Need to Know
    15/ Wireless Access Control
    16/ Account Monitoring and Control
    17/ Implement a Security Awareness and Training Program
    18/ Application Software Security
    19/ Incident Response and Management
    20/ Penetration Tests and Red Team Exercises

A great audit tool for the Critical Controls can be found at auditscripts.com: 

https://www.auditscripts.com/download/4229/

In a multi-OS environment, PowerShell can still be used, but depending on your environment OSSEC, Ansible, Puppet or Chef may also be a good fits - the approach remains the same though, no matter what the underlying API calls might be.

If any errors are found, or if you have suggestions for improvement, please contact me directly at:

rob@coherentsecurity.com

https://www.coherentsecurity.com
