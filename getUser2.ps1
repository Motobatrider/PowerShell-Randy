﻿get-aduser -SearchBase "ou=china,ou=users,ou=root2,dc=vstage,dc=co" -searchscope "subtree" -filter * -Properties * | select name,lastlogondate,DistinguishedName,samaccountname,description | Export-Csv C:\0_WH\lastlogon.csv -Encoding UTF8 -NoType