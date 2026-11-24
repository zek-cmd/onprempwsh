# On Prem AD PowerShell scripts

A set of PowerShell scripts for managing an on-premises Windows Active Directory (AD) domain: creating and moving users, resetting passwords, cleaning up inactive accounts, redirecting folders, deploying software through Group Policy, and sending messages or shutting down computers remotely.

## Requirements

- Windows Server or a Windows PC with the **ActiveDirectory** PowerShell module (RSAT on workstations)
- PowerShell 5.1 for most scripts; some scripts need **PowerShell 7.0+** (noted below)
- Run as an account with the right permissions in the domain (usually Domain Admin or delegated rights)

## Scripts

### Start here
| Script | What it does |
|---|---|
| `menu.ps1` | Numbered menu that lists every script in this folder, explains what it does and how risky it is, then asks y/n before starting it. Run ``.\menu.ps1 -WhatIf`` to run the scripts that support it in their own preview mode - the y/n question is still asked first. |

### User accounts
| Script | What it does |
|---|---|
| `USERCREATE.ps1` | Creates AD users from a CSV file (`FullName` column) using an OU tree picker. Existing users are skipped. |
| `USERCREATE2.ps1` | Alternate version of the user creation script. |
| `nameconvert.ps1` | Converts names from "LastName, FirstName MI" to "FirstName MI. LastName" from a CSV. |

### Moving users
| Script | What it does |
|---|---|
| `MOVEUSER.ps1` | Moves all users from one OU to another, with a dialog to pick both OUs. |
| `MOVEUSER2.ps1`, `MOVEUSER3.ps1` | Alternate versions of the move script. |

### Passwords
| Script | What it does |
|---|---|
| `RESETPASSWORD.ps1` | Resets passwords for all users in a chosen OU. |
| `RESETPASSWORD2.ps1` | Alternate version of the reset script. |

### Inactive accounts and folders
| Script | What it does |
|---|---|
| `INACTIVITY.ps1` | Finds inactive accounts and their redirected folders, writes a CSV report, and can optionally move or delete them in an interactive grid. |
| `Find-Duplicate-AD-Names.ps1` | Read-only scan for duplicate, reordered and similar (close) display names. Saves a CSV report and shows the result in a grid. |
| `folderbased/` | Folder redirection and lookup helpers, plus a DFS namespace setup in `folderbased/DFS_script/`. The two hardening scripts - `setgpo.ps1` and `New-UsbEventCollector.ps1` - live in here too; see the table below. |

### Hardening through Group Policy
| Script | What it does |
|---|---|
| `folderbased/setgpo.ps1` | Fills an **existing** GPO (it never creates one) and links it to the whole domain or one OU: blocks regedit, wallpaper changes, lock screen image changes and taskbar settings, turns off the Microsoft Store and Offline Files, stops Windows caching domain logons, turns on verbose logon messages, deletes stale user profiles after 30 days, sets the three interactive logon Security Options (locked-session display, sign-in banner title and text), turns on WinRM, writes a baseline set of firewall rules (ICMP, WinRM, DNS, NTP, web - Domain profile only) and switches on USB plug / unplug auditing. Every step previews first and asks y/n, and a second run changes nothing that is already set. Run it in Windows PowerShell 5.1 - started from PowerShell 7 it hands the run over to 5.1 itself. |
| `folderbased/New-UsbEventCollector.ps1` | Run **once on the server** after `setgpo.ps1`. Clients each log USB plug and unplug in their own Security log (event 6416) and USB driver log; this script creates a Windows Event Collector subscription that pulls those events from every PC you name and stores them in one place, the server's **ForwardedEvents** log. It asks before replacing an existing subscription and deletes nothing. |

### Software and computers
| Script | What it does |
|---|---|
| `SOFTWAREDEPLOY.ps1` | Creates and shares a folder for GPO software deployment, with the right share and NTFS permissions. |
| `SHUTDOWN.ps1` | Remote shutdown manager for domain computers (PowerShell 7.0+). |
| `message.ps1` | Sends a message to remote computers (run from a domain controller). |
| `JIT/Start-JITServer.ps1` | Local web server on `http://localhost:40000/` for Just-In-Time (JIT) admin access: grant, revoke and check access, with a session log. (PowerShell 7.0+) |

## How to use

1. Open PowerShell as Administrator on a machine that can reach the domain controller.
2. Go to the scripts folder: `cd path\to\onprempwsh`
3. Run the script you need, for example `.\USERCREATE.ps1` - or open the menu with `.\menu.ps1` and pick it from the list.
4. Follow the dialog boxes or prompts.

Read the header comments at the top of each script before running it. Some scripts change or delete data in AD.

## Heads-up

- Several scripts contain a default initial or reset password in plain text. Change it before running them in production.
- Test on a non-production OU first.
- The USB collector needs each client to allow this server to read its Security log (the client must run WinRM and this server's account must be able to read the log, for example through the **Event Log Readers** group). If no events arrive in `ForwardedEvents`, check that first.
- Keep passwords, keys and CSV files with personal data out of the repository.

## Author

Execquielle John Goingco - ejaygoingco4@gmail.com
