<#     
       .NOTES
       ==============================================================================
       Created on:         2023/02/01 
       Created by:         Drago Petrovic
       Organization:       MSB365.blog
       Filename:           SetCallingPolicy.ps1
       Current version:    V1.00     
       Find us on:
             * Website:         https://www.msb365.blog
             * Technet:         https://social.technet.microsoft.com/Profile/MSB365
             * LinkedIn:        https://www.linkedin.com/in/drago-petrovic/
             * MVP Profile:     https://mvp.microsoft.com/de-de/PublicProfile/5003446
       ==============================================================================
       .DESCRIPTION
       Set Calling Policies for multiple Users by using CSV file.
            
       
       .NOTES

       .EXAMPLE
       .\SetCallingPolicy.ps1 
             
       .COPYRIGHT
       Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), 
       to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, 
       and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
       The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
       THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
       FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, 
       WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
       ===========================================================================
       .CHANGE LOG
        V0.10, 2022/01/31 - DrPe - Initial version
	V0.20,	2022/10/02 - DrPe - Group Names modified
        V0.21, 2022/10/06 - DrPe - Bugfixing License assigning
	V1.00, 2022/10/07 - DrPe - Major Versioning, Module updates, array fixing
			 
--- keep it simple, but significant ---
--- by MSB365 Blog ---
#>
Write-Host "   _____      __                _____                            ___           " -ForegroundColor Magenta
Write-Host "  / ___/___  / /_   _________ _/ / (_)___  ____ _   ____  ____  / (_)______  __" -ForegroundColor Magenta
Write-Host "  \__ \/ _ \/ __/  / ___/ __ `/ / / / __ \/ __ `/  / __ \/ __ \/ / / ___/ / / /" -ForegroundColor Magenta
Write-Host " ___/ /  __/ /_   / /__/ /_/ / / / / / / / /_/ /  / /_/ / /_/ / / / /__/ /_/ / " -ForegroundColor Magenta
Write-Host "/____/\___/\__/   \___/\__,_/_/_/_/_/ /_/\__, /  / .___/\____/_/_/\___/\__, /  " -ForegroundColor Magenta
Write-Host "                                        /____/  /_/                   /____/   " -ForegroundColor Magenta
Start-Sleep -s 3

# Get Location ID
			write-host "Gettering Tenant location ID..." -ForegroundColor Cyan
			start-sleep -s 2
			$Lid = Get-CsOnlineLisLocation | Sort-Object LocationID | select-object -ExpandProperty LocationID
			write-host "Tenant LocationID is: $Lid" -ForegroundColor White -BackgroundColor Black
			Start-Sleep -s 2

			# Getting CSV Information
			write-host "Please select and import the CSV File from your device:" -ForegroundColor Cyan
			Write-Host ""
			Write-Host ""
			Start-Sleep -s 4
			Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
			Write-Host "NOTE!" -ForegroundColor Yellow -BackgroundColor Black
			Write-Host "The following information are needed in the CSV file:" -ForegroundColor White -BackgroundColor black
			Write-Host '"UserPrincipalName","DisplayName","CallingPolicy"' -ForegroundColor Gray -BackgroundColor Black
			Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
			Start-Sleep -s 6
			$File = New-Object System.Windows.Forms.OpenFileDialog
			$null = $File.ShowDialog()
			$FilePath = $File.FileName
			$users = Import-Csv $FilePath
			Start-Sleep -s 3
			Write-Host "The following data is imported from CSV:" -ForegroundColor Cyan
			$users | ft
			Start-Sleep -s 3
			do
			{
				$selection = Read-Host "Are the data correct? - Choose between [Y] and [N]"
				switch ($selection)
				{
					'y' {
						
					} 'n' {
						$File = New-Object System.Windows.Forms.OpenFileDialog
						$null = $File.ShowDialog()
						$FilePath = $File.FileName
						$users = Import-Csv $FilePath
						Start-Sleep -s 3
						Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
						Write-Host "The following data is imported from CSV:" -ForegroundColor Cyan
						$users | ft
						Start-Sleep -s 3
					}
				}
			}
			until ($selection -eq "y")
			

			start-sleep -s 3

			# Configuring Phone Number for Teams users
			write-host "Setting the CallingPolicy..." -ForegroundColor cyan 
			foreach($user in $users)
			{
				try
				{
                    Grant-CsTeamsCallingPolicy -identity $user.UserPrincipalName -PolicyName $user.CallingPolicy
					Write-Host "CallingPolicy for the users $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set the CallingPolicy for the user $($user.UserPrincipalName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
			start-sleep -s 3
			Write-Host "All CallingPolicy for the Users!" -ForegroundColor Green -BackgroundColor Black
			pause