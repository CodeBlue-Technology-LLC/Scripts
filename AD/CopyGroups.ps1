Get-ADUser "sourceUser" -Properties MemberOf | Select-Object -ExpandProperty MemberOf | Add-ADGroupMember -Members "targetUser"
