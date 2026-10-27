# On Prem AD PowerShell scripts

A set of PowerShell scripts for managing an on-premises Windows Active Directory (AD) domain: creating and moving users, resetting passwords, cleaning up inactive accounts, redirecting folders, deploying software through Group Policy, and sending messages or shutting down computers remotely.

## Requirements

- Windows Server or a Windows PC with the **ActiveDirectory** PowerShell module (RSAT on workstations)
- PowerShell 5.1 for most scripts; some scripts need **PowerShell 7.0+** (noted below)
- Run as an account with the right permissions in the domain (usually Domain Admin or delegated rights)

## Scripts

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
| `folderbased/` | Folder redirection and lookup helpers, plus a DFS namespace setup in `folderbased/DFS_script/`. |

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
3. Run the script you need, for example `.\USERCREATE.ps1`.
4. Follow the dialog boxes or prompts.

Read the header comments at the top of each script before running it. Some scripts change or delete data in AD.

## Heads-up

- Several scripts contain a default initial or reset password in plain text. Change it before running them in production.
- Test on a non-production OU first.
- Keep passwords, keys and CSV files with personal data out of the repository.

## Author

Execquielle John Goingco - ejaygoingco4@gmail.com
