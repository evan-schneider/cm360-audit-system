# Security Policy

## Supported Versions

| Version | Supported          |
| ------- | ------------------ |
| 1.0.x   | :white_check_mark: |

## Reporting a Vulnerability

We take the security of CM360 Audit System seriously. If you believe you have found a security vulnerability, please report it to us as described below.

**Please do not report security vulnerabilities through public GitHub issues.**

Instead, please send a report to: **evan.schneider98@gmail.com**

You should receive a response within 48 hours. If for some reason you do not, please follow up to ensure we received your original message.

Please include the requested information listed below (as much as you can provide) to help us better understand the nature and scope of the possible issue:

* Type of issue (e.g., improper access controls, data exposure, injection, etc.)
* Full paths of source file(s) related to the manifestation of the issue
* The location of the affected source code (tag/branch/commit or direct URL)
* Any special configuration required to reproduce the issue
* Step-by-step instructions to reproduce the issue
* Proof-of-concept or exploit code (if possible)
* Impact of the issue, including how an attacker might exploit the issue

This information will help us triage your report more quickly.

## Security Considerations

This system handles sensitive campaign data and integrates with multiple Google services. Key security areas include:

### Data Handling
- Campaign performance data
- Email addresses and distribution lists
- Google Drive file access
- Gmail message processing

### Access Controls
- Apps Script execution permissions
- Google Workspace service access
- Spreadsheet sharing permissions
- Drive folder permissions

### Common Security Best Practices
- Keep external configuration spreadsheets properly secured
- Regularly review Gmail label filters and permissions
- Monitor Apps Script execution logs for anomalies
- Limit sharing of deployment scripts and configuration data
- Use staging mode for testing to prevent accidental data exposure

## Responsible Disclosure

We kindly ask that you:
- Allow us reasonable time to investigate and address the issue before any disclosure
- Avoid privacy violations, destruction of data, or disruption of services
- Only interact with accounts you own or have explicit permission to access
- Do not access or modify data that does not belong to you

## Recognition

We appreciate the security research community's efforts to keep our users safe. Valid security reports will be acknowledged, and we're happy to credit researchers (with their permission) for their responsible disclosure.