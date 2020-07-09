<#
This script uses the cmdlets Start-Transaction and Complete-Transaction, which 
are only present in Windows PowerShell but are absent from 
PowerShell Core 6+.

ActiveX settings are stored in two different hives in the Registry: 
- HKCU contains user-specific settings
- HKLM contains machine-wide settings

Value    Setting
   ----------------------------------------------------------------------------------
   1001     ActiveX controls and plug-ins: Download signed ActiveX controls
   1004     ActiveX controls and plug-ins: Download unsigned ActiveX controls
   1200     ActiveX controls and plug-ins: Run ActiveX controls and plug-ins
   1201     ActiveX controls and plug-ins: Initialize and script ActiveX controls not marked as safe for scripting
   1206     Miscellaneous: Allow scripting of Internet Explorer Web browser control ^
   1207     Reserved #
   1208     ActiveX controls and plug-ins: Allow previously unused ActiveX controls to run without prompt ^
   1209     ActiveX controls and plug-ins: Allow Scriptlets
   120A     ActiveX controls and plug-ins: ActiveX controls and plug-ins: Override Per-Site (domain-based) ActiveX restrictions
   120B     ActiveX controls and plug-ins: Override Per-Site (domain-based) ActiveX restrictions
   1400     Scripting: Active scripting
   1402     Scripting: Scripting of Java applets
   1405     ActiveX controls and plug-ins: Script ActiveX controls marked as safe for scripting
   1406     Miscellaneous: Access data sources across domains
   1407     Scripting: Allow Programmatic clipboard access
   1408     Reserved #
   1409     Scripting: Enable XSS Filter
   1601     Miscellaneous: Submit non-encrypted form data
   1604     Downloads: Font download
   1605     Run Java #
   1606     Miscellaneous: Userdata persistence ^
   1607     Miscellaneous: Navigate sub-frames across different domains
   1608     Miscellaneous: Allow META REFRESH * ^
   1609     Miscellaneous: Display mixed content *
   160A     Miscellaneous: Include local directory path when uploading files to a server ^
   1800     Miscellaneous: Installation of desktop items
   1802     Miscellaneous: Drag and drop or copy and paste files
   1803     Downloads: File Download ^
   1804     Miscellaneous: Launching programs and files in an IFRAME
   1805     Launching programs and files in webview #
   1806     Miscellaneous: Launching applications and unsafe files
   1807     Reserved ** #
   1808     Reserved ** #
   1809     Miscellaneous: Use Pop-up Blocker ** ^
   180A     Reserved # 
   180B     Reserved #
   180C     Reserved #
   180D     Reserved #
   180E     Allow OpenSearch queries in Windows Explorer #
   180F     Allow previewing and custom thumbnails of OpenSearch query results in Windows Explorer #
   1A00     User Authentication: Logon
   1A02     Allow persistent cookies that are stored on your computer #
   1A03     Allow per-session cookies (not stored) #
   1A04     Miscellaneous: Don't prompt for client certificate selection when no certificates or only one certificate exists * ^
   1A05     Allow 3rd party persistent cookies *
   1A06     Allow 3rd party session cookies *
   1A10     Privacy Settings *
   1C00     Java permissions #
   1E05     Miscellaneous: Software channel permissions
   1F00     Reserved ** #
   2000     ActiveX controls and plug-ins: Binary and script behaviors
   2001     .NET Framework-reliant components: Run components signed with Authenticode
   2004     .NET Framework-reliant components: Run components not signed with Authenticode
   2007     .NET Framework-Reliant Components: Permissions for Components with Manifests
   2100     Miscellaneous: Open files based on content, not file extension ** ^
   2101     Miscellaneous: Web sites in less privileged web content zone can navigate into this zone **
   2102     Miscellaneous: Allow script initiated windows without size or position constraints ** ^
   2103     Scripting: Allow status bar updates via script ^
   2104     Miscellaneous: Allow websites to open windows without address or status bars ^
   2105     Scripting: Allow websites to prompt for information using scripted windows ^
   2200     Downloads: Automatic prompting for file downloads ** ^
   2201     ActiveX controls and plug-ins: Automatic prompting for ActiveX controls ** ^
   2300     Miscellaneous: Allow web pages to use restricted protocols for active content **
   2301     Miscellaneous: Use Phishing Filter ^
   2400     .NET Framework: XAML browser applications
   2401     .NET Framework: XPS documents
   2402     .NET Framework: Loose XAML
   2500     Turn on Protected Mode [Vista only setting] #
   2600     Enable .NET Framework setup ^
   2702     ActiveX controls and plug-ins: Allow ActiveX Filtering
   2708     Miscellaneous: Allow dragging of content between domains into the same window
   2709     Miscellaneous: Allow dragging of content between domains into separate windows
   270B     Miscellaneous: Render legacy filters
   270C     ActiveX Controls and plug-ins: Run Antimalware software on ActiveX controls 

   {AEBA21FA-782A-4A90-978D-B72164C80120}   First Party Cookie *
   {A8A88C49-5EB2-4990-A1A2-0876022C854F}   Third Party Cookie *

*  indicates an Internet Explorer 6 or later setting
** indicates a Windows XP Service Pack 2 or later setting
#  indicates a setting that is not displayed in the user interface in Internet Explorer
^  indicates a setting that only has two options, enabled or disabled

(table found at https://support.microsoft.com/en-us/help/182569/internet-explorer-security-zones-registry-entries-for-advanced-users)

Internet Zone ActiveX Settings
- [x] 2702: Allow ActiveX Filtering = Enable
- [x] 1208: Allow previously unused ActiveX controls to run without prompt = Disable
- [x] 1209: Allow Scriptlets = Disable
- [x] 2201: Automatic prompting for ActiveX Controls = Disable
- [ ] Binary and script behaviors = Enable
- [ ] Display video and animation on a web page that does not use external media player = Disable
- [ ] Download signed ActiveX controls = Prompt
- [ ] Download unsigned ActiveX controls = Disable
- [ ] Initialize and script ActiveX controls not marked as safe for scripting = Disable
- [ ] Only allow approved domains to use ActiveX without prompt = Enable
- [ ] Run ActiveX controls and plug-ins = Enable
- [ ] Run antimalware software on ActiveX controls = Enable
- [ ] Script ActiveX controls marked safe for scripting* = Enable
#>

$hklm_inet = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3'
$hkcu_inet = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3'


Start-Transaction
Use-Transaction {
  # Internet Zone: Allow ActiveX Filtering = Enable
  Set-ItemProperty -Path $hklm_inet -Name 2702 -Value 0
  Set-ItemProperty -Path $hkcu_inet -Name 2702 -Value 0

  # Internet Zone: Allow previously unused ActiveX controls to run without prompt = Disable
  Set-ItemProperty -Path $hkcu_inet -Name 1208 -Value 3
  Set-ItemProperty -Path $hklm_inet -Name 1208 -Value 3

  # Internet Zone: Allow Scriptlets = Disable
  Set-ItemProperty -Path $hklm_inet -Name 1209 -Value 3
  Set-ItemProperty -Path $hkcu_inet -Name 1209 -Value 3
  
  # Automatic prompting for ActiveX Controls = Disable
  Set-ItemProperty -Path $hklm_inet -Name 2201 -Value 3
  Set-ItemProperty -Path $hkcu_inet -Name 2201 -Value 3
  

} -UseTransaction
Complete-Transaction




