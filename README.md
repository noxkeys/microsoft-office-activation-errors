# Microsoft Office Activation Errors and Fixes

A complete guide to fixing **"Unlicensed Product"**, **"Product Activation Failed"** and other Microsoft Office activation errors. It covers Office 2024, 2021, 2019, 2016 and Microsoft 365 on Windows.

Maintained by [Noxkeys](https://noxkeys.com/), an online store for genuine software licences. Corrections are welcome.

---

## Table of contents

- [Quick fixes to try first](#quick-fixes-to-try-first)
- [How to check your Office activation status](#how-to-check-your-office-activation-status)
- [Common Office activation messages](#common-office-activation-messages)
- [Office activation error codes](#office-activation-error-codes)
- [Step-by-step fixes](#step-by-step-fixes)
- [When you need a new product key](#when-you-need-a-new-product-key)
- [FAQ](#faq)
- [Related guides](#related-guides)
- [References](#references)

---

## Quick fixes to try first

Most Office activation problems are fixed by one of these steps:

1. **Sign in with the right account:** open any Office app, go to **File → Account** and make sure you're signed in with the account you used to buy Office
2. **Fix the date and time:** turn on **Set time automatically** and **Set time zone automatically** in Windows settings
3. **Run Office as administrator:** right-click Word or Excel and choose **Run as administrator**
4. **Remove extra copies of Office:** having two versions installed is a common cause of activation failure
5. **Restart the PC** and open an Office app again

If none of these work, find your message or error code below.

---

## How to check your Office activation status

**In any Office app:** go to **File → Account**. Under *Product Information* you'll see one of these:

| Status shown                     | Meaning                                         |
| -------------------------------- | ----------------------------------------------- |
| **Product Activated**            | Office is activated. Nothing to fix             |
| **Activation Required**          | Office is installed but not activated           |
| **Unlicensed Product**           | No valid licence was found. Most features are disabled |
| **Subscription Product**         | Microsoft 365 is active through a subscription  |

**From Command Prompt (Admin):**

```bat
:: 64-bit Office
cd /d "C:\Program Files\Microsoft Office\Office16"

:: 32-bit Office on 64-bit Windows
:: cd /d "C:\Program Files (x86)\Microsoft Office\Office16"

:: Show licence status, edition, channel and the last 5 characters of the key
cscript ospp.vbs /dstatus
```

Office 2016, 2019, 2021 and 2024 all use the `Office16` folder. `LICENSE STATUS: ---LICENSED---` means Office is activated.

---

## Common Office activation messages

| Message                                                    | Most likely cause                                        | Fix                                                                         |
| ---------------------------------------------------------- | -------------------------------------------------------- | --------------------------------------------------------------------------- |
| **"Unlicensed Product"** in the title bar                  | Not activated, wrong account, or an invalid key          | See [Fix "Unlicensed Product"](#1-fix-unlicensed-product)                   |
| **"Product Activation Failed"**                            | The installed key is invalid, blocked or for a different version | [Remove the old key](#4-remove-and-replace-a-product-key) and enter a genuine one |
| **"We couldn't find any Office products"**                 | Signed in with the wrong account, or the subscription expired | Sign in with the account used to buy Office, or renew Microsoft 365   |
| **"This product key has already been used"**               | The key was redeemed on another account or PC            | Contact the seller, or use a new genuine key                                |
| **"This product key isn't valid"**                         | Typo, or a key for a different Office version            | Check the key matches your version (e.g. a 2021 key for Office 2021)       |
| **"Microsoft Office can't find your license for this application"** | The licence is damaged or missing                | [Repair Office](#3-repair-office), then activate again                      |
| **"Non-commercial use"** with **"Unlicensed Product"**     | A trial or unactivated copy                              | Enter a genuine product key                                                 |

---

## Office activation error codes

| Error code   | What it means                                          | Fix                                                                         |
| ------------ | ------------------------------------------------------ | --------------------------------------------------------------------------- |
| `0xC004F074` | No KMS server could be contacted                       | A volume (KMS) key is installed on a home PC. Replace it with a retail key |
| `0xC004C060` | The product key is no longer valid                     | Replace it with a genuine key from an authorised seller                     |
| `0xC004C003` | The product key was blocked                            | Use a different genuine key                                                 |
| `0xC004C008` | The key has been used on too many devices              | Each PC needs its own licence. Contact the seller for a replacement        |
| `0x8004FC12` | Activation failed, often after a Windows upgrade       | Install the latest Windows updates, or use the [commands below](#6-fix-error-0x8004fc12) |
| `0xC004F050` | The product key is invalid                             | Check for typos and make sure the key matches the edition (e.g. Pro Plus)  |

> **KMS key on a home PC?** Office volume (KMS) keys only work inside organisations with a KMS server. See [Microsoft Office 2024, 2021 & 2019 Product Keys](https://github.com/webex9/microsoft-office-2024-2021-2019-product-keys) for what those keys do.

---

## Step-by-step fixes

### 1. Fix "Unlicensed Product"

Microsoft lists these causes. Work through them in order:

**a) Signed in with the wrong account**
1. Open Word, go to **File → Account** and click **Sign out**
2. Sign in again with the email address you used to buy Office
3. If you use Office from work or school, sign in with that account instead

**b) More than one copy of Office is installed**
1. Open **Control Panel → Programs → Programs and Features**
2. Type `office` in the search box
3. Uninstall every Office version you don't use
4. Restart the PC, open an Office app and sign in

**c) Microsoft 365 subscription expired**
1. Close all Office apps
2. Check your subscription status in your Microsoft account, under **Services & subscriptions**
3. Renew it if needed, then restart Office

**d) Office 2024, 2021, 2019 or 2016 not activated**
Enter a genuine product key. See [Enter a product key](#2-enter-a-product-key).

### 2. Enter a product key

**In an Office app:**
1. Open Word or Excel
2. Go to **File → Account**
3. Click **Change Product Key** (or **Activate Product**)
4. Enter your 25-character key and follow the prompts

**From Command Prompt (Admin):**

```bat
cd /d "C:\Program Files\Microsoft Office\Office16"

:: Install the key
cscript ospp.vbs /inpkey:XXXXX-XXXXX-XXXXX-XXXXX-XXXXX

:: Activate
cscript ospp.vbs /act
```

### 3. Repair Office

1. Close all Office apps
2. Open **Settings → Apps → Installed apps**
3. Find **Microsoft Office** (or **Microsoft 365**), click **⋯** and choose **Modify**
4. Choose **Quick Repair** first. If the problem continues, run **Online Repair**
5. Restart the PC and activate again

### 4. Remove and replace a product key

Use this when an old, invalid or KMS key is stuck on the PC.

```bat
cd /d "C:\Program Files\Microsoft Office\Office16"

:: 1. Find the last 5 characters of the installed key
cscript ospp.vbs /dstatus

:: 2. Remove it (replace XXXXX with those 5 characters)
cscript ospp.vbs /unpkey:XXXXX

:: 3. Install your genuine key and activate
cscript ospp.vbs /inpkey:XXXXX-XXXXX-XXXXX-XXXXX-XXXXX
cscript ospp.vbs /act
```

If `/dstatus` shows more than one key, remove each one before adding the new key.

### 5. Fix network, firewall and time problems

Microsoft recommends checking these when activation can't connect:

- **Date and time:** turn on *Set time automatically* and *Set time zone automatically*, then restart Office
- **Firewall:** check Windows Firewall or your third-party firewall settings, and temporarily turn it off while activating
- **Antivirus:** pause it while Office activates, then turn it back on
- **Proxy:** if you switch between home and work networks, turn off any manual proxy under *Settings → Network & internet → Proxy*
- **Updates:** install the latest Office updates (*File → Account → Update Options → Update Now*)

### 6. Fix error 0x8004FC12

Microsoft's fix is to **install the latest Windows updates** first. If you can't update right away, open **Command Prompt (Admin)** and run each command:

```bat
net localgroup Administrators localservice /add
fsutil resource setautoreset true C:\
netsh int ip reset resetlog.txt
```

Restart the PC and try activating Office again.

---

## When you need a new product key

Some problems can't be fixed without a new key:

- The key is blocked or no longer valid (`0xC004C003`, `0xC004C060`)
- The key was already used on another PC or account (`0xC004C008`, "already been used")
- A volume (KMS) key is installed on a home PC (`0xC004F074`)
- Office 2019 or 2016 is installed and you want security updates again (both reached end of support on October 14, 2025)
- You have Microsoft 365 but want a **one-time purchase** instead of a subscription

Genuine **Office 2024 Professional Plus** and **Office 2021 Professional Plus** keys with instant email delivery are available at **[noxkeys.com](https://noxkeys.com/)**.

> Avoid unofficial activators and "KMS crack" tools. They often contain malware and stop working after Office updates.

---

## FAQ

**What does "Unlicensed Product" mean in Office?**
Office couldn't find a valid licence. Most editing features are disabled until you sign in with the right account or enter a genuine product key.

**Can I still use Office if it's unlicensed?**
You can open and read documents, but editing and saving are restricted.

**Can I use an Office 2021 key on Office 2024?**
No. Each key only works with its own version and edition.

**Why does Office keep asking me to activate?**
Usually the wrong account is signed in, the date and time are wrong, or two copies of Office are installed. Work through the [quick fixes](#quick-fixes-to-try-first).

**Do I need to reinstall Office after changing the key?**
No. Remove the old key and enter the new one. A reinstall is only needed if the key is for a different edition from the one installed.

**How do I move Office to a new PC?**
For retail Office 2024 or 2021, uninstall it from the old PC first, then install and activate it on the new one with the same key or account.

---

## Related guides

More free references maintained by [Noxkeys](https://noxkeys.com/):

- [Windows 11 & 10 Product Keys](https://github.com/webex9/windows-11-10-product-keys): generic and KMS client keys for every Windows edition
- [Microsoft Office 2024, 2021 & 2019 Product Keys](https://github.com/webex9/microsoft-office-2024-2021-2019-product-keys): Office KMS keys, ospp.vbs commands and version comparison
- [Windows Server 2025, 2022 & 2019 Product Keys](https://github.com/webex9/windows-server-2025-2022-2019-product-keys): Server KMS keys, Standard vs Datacenter and Evaluation conversion
- [Windows 11 & 10 Activation Errors and Fixes](https://github.com/webex9/windows-11-10-activation-errors): Windows activation error codes explained
- [How to Upgrade Windows 11 Home to Pro](https://github.com/webex9/upgrade-windows-11-home-to-pro): step-by-step upgrade without reinstalling

---

## References

- [Microsoft Support – Unlicensed Product and activation errors in Office](https://support.microsoft.com/office/0d23d3c0-c19c-4b2f-9845-5344fedc4380)
- [Microsoft Support – Office error code 0x8004FC12 when activating Office](https://support.microsoft.com/topic/f32c97d5-11f7-4e56-925c-94959b222bd3)
- [Microsoft Learn – Tools to manage volume activation of Office (ospp.vbs)](https://learn.microsoft.com/en-us/office/volume-license-activation/tools-to-manage-volume-activation-of-office)
- [Microsoft Office 2024, 2021 & 2019 Product Keys](https://github.com/webex9/microsoft-office-2024-2021-2019-product-keys) – KMS keys and version comparison
- [Windows 11 & 10 Activation Errors and Fixes](https://github.com/webex9/windows-11-10-activation-errors) – Windows activation error codes
- [How to Upgrade Windows 11 Home to Pro](https://github.com/webex9/upgrade-windows-11-home-to-pro)
- [Noxkeys – Genuine Microsoft Office keys](https://noxkeys.com/)

---

⭐ **If this helped, star the repo** so others can find it. Last updated: October 2026.
