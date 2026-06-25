# Security Policy

## Supported Versions

Security fixes are applied to the latest released version only.

| Version | Supported          |
| ------- | ------------------ |
| 1.0.x   | :white_check_mark: |
| < 1.0   | :x:                |

## Reporting a Vulnerability

**Please do not report security vulnerabilities through public GitHub issues or
discussions.**

Report them privately through GitHub Security Advisories:

1. Go to the **Security** tab of this repository
2. Click **Report a vulnerability**
3. Include a description, steps to reproduce, affected version, and your
   environment (OS, Python version)

This keeps the report private until a fix is available. As this is a project
maintained in spare time, please allow a few days for an initial response.

## How This Application Handles Your Data

This is a desktop application that connects to your email account over IMAP to
download attachments. Understanding what it does — and does not do — with your
data is part of using it safely.

- **Credentials stay in memory.** Your email address and password (or App
  Password) are held only in memory for the duration of the session and are not
  written to disk or logged by the application.
- **Read-only access.** The app only reads messages and downloads attachments;
  it does not send, modify, move, or delete email.
- **Encrypted transport.** All connections to the mail server use SSL/TLS
  (IMAP over port 993).
- **Local storage only.** Downloaded attachments are saved to a local folder you
  choose. Nothing is uploaded anywhere else.

## Recommendations for Users

- **Use an App Password, not your main password.** For Gmail and for Outlook
  accounts with 2FA, generate a dedicated App Password and use that. You can
  revoke it at any time without changing your main password.
- **Download the installer only from official Releases.** Get the `.exe` from
  this repository's [Releases](../../releases) page, and verify its SHA-256
  checksum against the value published with the release before running it.
- **Review the source.** Because the app handles email credentials, you are
  encouraged to read `src/email_client.py` and confirm for yourself how
  credentials are used before entering them.

## Threat Model & Limitations

- The application is only as trustworthy as the machine it runs on. Malware,
  keyloggers, or a compromised OS can capture credentials regardless of what the
  app does.
- App Passwords grant mailbox access. Treat them like passwords: do not share
  them, and revoke any you no longer use.
- The maintainer cannot guarantee the security of builds distributed outside the
  official Releases of this repository.

## Scope

In scope: credential handling, logging of sensitive data, insecure transport,
unsafe file handling of downloaded attachments, and dependency vulnerabilities
in this project.

Out of scope: vulnerabilities in Gmail, Outlook, or the upstream libraries
(`customtkinter`, `tkcalendar`, `Pillow`, etc.) — report those to their
respective maintainers.
