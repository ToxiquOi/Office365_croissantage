#OEM850 Requis par outlook

powershell ^
 $Outlook = New-Object -ComObject Outlook.Application; ^
 $Mail = $Outlook.CreateItem(0); ^
 $Mail.To = ''; ^
 $Mail.Subject = 'Croissantage'; ^
 $Mail.Body = \""Bonjour � tous ! `nJ'ai le plaisir de vous informer que je ram�ne les petits pains demain matin ! `nNe me remerciez pas, je penserai � verrouiller mon poste la prochaine fois ! ;) \""; ^
 $Mail.Send(); ^