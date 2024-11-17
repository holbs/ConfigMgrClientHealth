# ConfigMgrClientHealth
This is a fork of [Anders Rodland's ConfigMgr Client Health script](https://github.com/AndersRodland/ConfigMgrClientHealth), with some minor updates as detailed below in the change log, and with installation, deployment, and distribution steps using SYSVOL
## Installation
- In SYSVOL create the folder ConfigMgrClientHealth then copy the following files, folders and their contents from this repositry to this folder.
  - Logs
  - Config.xml
  - ConfigMgrClientHealth.ps1
- SYSVOL is used for distribution, rather than using a single server like the standard deployment which could exist far away geographically.
- CONTOSO\Everyone should have read and execute access to the ConfigMgrClientHealth folder.
- CONTOSO\Everyone should have read, execute, write and modify to the Logs folder.
- Edit the Config.xml to suit your environment. Anders has documented the options you can use [here](https://www.andersrodland.com/configmgr-client-health/).
- In this version there are additional updates to Config.xml to check for, and turn off metered network connections, and IPv6
## Deployment
- Deploy as a scheduled task with Group Policy to start at 08:00 daily, with an 8 hour random delay with the action as below:
```
%windir%\System32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy bypass -File "\\contoso.com\SYSVOL\contoso.com\ConfigMgrClientHealth\ConfigMgrClientHealth.ps1" -Config "\\contoso.com\SYSVOL\contoso.com\ConfigMgrClientHealth\Config.xml"
```
## Change Log
- Updated WMI repair function to a PowerShell version of the script [here](https://www.reddit.com/r/sysadmin/comments/15uux4z/wmi_repair_script_built_in_native_windows_command/)
- Ignored event ID 7017 when checking for Group Policy errors as detailed [here](https://www.reddit.com/r/SCCM/comments/1aow39q/updates_and_feature_updates_stuck_at_0_download/)
- Registration errors are remediated by forcing MeteredNetworkUsage to 1 (allowed)
- Client reinstalls now disable Metered Connections before reinstalling the client. A client on a metered connection won't register after installation as detailed [here](https://www.asquaredozen.com/2020/05/22/lockdown-diary-metered-internet-connections-and-broken-configmgr-clients/)
