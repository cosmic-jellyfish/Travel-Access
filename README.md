# Travel Access Scheduler

If you block sign-in from other countries, you still have staff who go overseas for work. The block is the right default. The exception is one person, for the days they are away, and then it should stop.

Entra already has a place for that. A security group on the Conditional Access policy. While someone is in the group, the country block does not apply to them. When they come out, it applies again. The messy part is the membership. People get added the morning they fly, or they stay in the group for weeks after they are home, because nobody went back and took them out.

This is a scheduled task that keeps that one group up to date. You write down who is travelling and the dates. On those dates the task puts them in the group. After the end date it takes them out. It does not edit the policy, and it does not turn off MFA or any other sign-in rule. Someone you added to the group yourself in Entra, who has never been on the trip list, is left alone.

You still attach the group to the policy. That is once. Use the policy you already have for overseas access, or create one and call it Travel access, then add the group. If the policy blocks other countries, the group goes under Exclude. Exclude is the side that lets those people through.

After that it is just the trip list. Travel Desk is the window for adding and removing people. Save, then close it. The task runs when the computer starts, and about every two hours, whether the window is open or not. It tells you the next update in ordinary language, for example today at 2:05 pm.

## Privacy matters

This runs on the computer you install it on. There is no account, no subscription, and no service in the middle. It does not send telemetry, and it does not report usage back to anyone. The trip list is a file in the `data` folder. It stays there.

The computer does talk to your Microsoft 365 tenant. That is the add-or-remove on the travel group, using the app you set up during install. Microsoft sees that, same as any other group change. The schedule itself is not uploaded anywhere.

If the disk dies, or the server is rebuilt, those files are gone unless you kept a copy. That backup is yours to do. Nobody else has it.

Copy the `data` folder somewhere safe. The files that matter:

- `TravelSchedule.csv`, the trips. Lose this and you lose who is travelling.
- `Config.json`, the tenant, the app, and the group. The installer needs this to recognise the existing setup.
- `State.json`, the people the tool has already managed. Without it, a rebuilt server will not know who to take back out of the group.
- `Logs`, only if you want the history of who was added and removed.

The certificate private key is in the Windows certificate store on that computer, not in `data`. A replacement machine means running `Install.cmd` again so a new certificate can be created and put on the app. The backup is what saves the trips and the setup. The logs can be skipped if you do not care about them.

## What you need

- Windows Server or Windows 10/11, with Windows PowerShell 5.1
- A Global Administrator for the first install
- The Conditional Access policy that lets people in from overseas, or permission to create one
- A permanent folder, for example `C:\TravelAccess`. Do not run it from inside the ZIP

## What you do, and what the installer does

You do not make the group. The installer makes it. The default name is `CA - Approved Overseas Travellers`. If you already have a normal security group you want to use, type that name instead.

You do make sure the group is on the overseas access policy. That is the step the installer cannot do for you.

| You | The installer |
| --- | --- |
| Add this group to the overseas policy, or create a policy named Travel access and then add the group | One group for the travellers |
| Leave MFA and the other sign-in rules alone | One scheduled task, at startup and about every 2 hours |
| Add and remove trips in Travel Desk, then close it | One app registration, if you ask it to, and one certificate on this computer |

Do not make a second group for taking access away. When a trip ends, the person is removed from this same group. Someone you add to the group only in Entra, who has never been in the schedule, is left alone.

## Install

1. Extract the ZIP to a permanent folder.
2. Double-click `Install.cmd` and approve the administrator prompt.
3. On a new install, enter a short organisation name. It is used for the task, the certificate and the app registration.
4. Choose **1** to create the app registration. Sign in as a Global Administrator.
   Or choose **2** if you already created an app with these **application** permissions and admin consent:
   - `GroupMember.ReadWrite.All`
   - `User.Read.All`
5. Accept the group name, or type an existing static security group.
6. Accept the Windows time zone, or type another one. The suggestion is this server's time zone. Examples: `AUS Eastern Standard Time`, `GMT Standard Time`, `Pacific Standard Time`.
7. Type `YES` to confirm the app. Then add the travel group to the overseas policy (create a policy named Travel access first, if you do not have one) and type `YES` again.

Running the installer again reuses the certificate while it has more than 30 days left, and replaces the certificate on the app registration so Entra still has one. Any older certificate with the same name is removed from this computer. A new certificate is created only when the current one is missing or due to expire within 30 days. An installation that already has `Config.json` keeps its task name, certificate name and app name. If a scheduled task of that name already points at a different folder, the installer stops before it changes anything.

## Day to day

Double-click `TravelDesk.cmd`.

The tabs across the top are **Schedule**, **Status** and **Setup**.

- **Schedule** — add, edit and remove trips. Dates are `dd/MM/yyyy` and include both ends. The top of the page shows when the task will next update access. Save, then close the window, or use **Run sync now**. The task also runs at startup and about every two hours.
- **Status** — task, certificate expiry and the latest log.
- **Setup** — the group and Conditional Access steps, the time zone, and buttons that open Entra.

A trip date is that calendar day in the time zone saved in `Config.json`, including daylight saving for that zone. A new install suggests the server's Windows time zone. An existing config keeps the zone already saved there. Change it on the Setup tab if the server clock is UTC but the travel day should follow another country. The next sync uses the saved zone.

You can still edit `data\TravelSchedule.csv` directly:

```text
UserPrincipalName,StartDate,EndDate
alex.chen@contoso.com,05/10/2026,16/10/2026
```

Delete a row to cancel that trip. Every row is active for its dates. There is no enable/disable column.

Check the same information in a console with `.\app\Status.ps1`.

## Certificate

The certificate is self-signed. In the certificate store, Issued To and Issued By are the same name. That is expected. The private key stays on this server and is not exported. There is one certificate, and the installer keeps it that way.

Open the store with `certlm.msc`. Do not use `certmgr.msc`; that is the current-user store, and this certificate is on the computer.

## Uninstall

Double-click `Uninstall.cmd`, or run it from an elevated command prompt:

```text
Uninstall.cmd
Uninstall.cmd -RemoveCertificate
```

The task is removed. `-RemoveCertificate` also removes the local certificate for this scheduler. The Entra app, the travel group, and the overseas policy are left in place. Delete the certificate under the app's **Certificates & secrets** blade if you are retiring the app. Remove the group from that policy when you no longer want the overseas access.

## Security

- Use a dedicated app registration. The installer replaces the certificate on that app.
- The app can read users and change group membership. It cannot change Conditional Access, roles or passwords.
- Add the travel group only to the policy that lets people in from overseas. If that policy blocks other countries, the group belongs under Exclude.
- Leave MFA, device compliance, and sign-in frequency alone.
- `data\Config.json` contains your tenant ID, application ID and group ID. Do not publish it. It is listed in `.gitignore`.

## Files

Day to day, double-click `TravelDesk.cmd`. Setup is `Install.cmd`.

| File | What it is |
| --- | --- |
| `TravelDesk.cmd` | Opens the window |
| `Install.cmd` | Setup |
| `Uninstall.cmd` | Removes the scheduled task |
| `app\` | The programs. The scheduled task runs `app\Sync.ps1`. |
| `data\` | Created on install. Holds the schedule, settings, and logs. Do not publish it. |

## Logs

| File | What it is |
| --- | --- |
| `data\Logs\Install.log` | Installer transcript |
| `data\Logs\TravelAccess.log` | Each sync, including who was added or removed |

## Licence

MIT. See [LICENSE](LICENSE). Replace the copyright notice if your organisation should hold it.
