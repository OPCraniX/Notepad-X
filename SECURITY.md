# Security Policy

## Supported Versions

Security fixes are provided for:

| Version | Supported |
| --- | --- |
| Latest `main` branch | Yes |
| Older commits / forks | No |

## Reporting a Vulnerability

If you find a security issue in Notepad-X:

1. Do not post the details in a public issue.
2. Use GitHub's private vulnerability reporting for this repository if it is available.
3. If private reporting is not available, contact the repository owner directly through GitHub before disclosing details publicly.

When reporting, please include:

- A short description of the issue
- Steps to reproduce it
- The affected file or feature
- Any proof-of-concept details needed to confirm the issue
- Your estimate of the impact

You can expect:

- An acknowledgement after the report is reviewed
- A confirmation if the issue is reproducible
- A fix or mitigation plan when appropriate

Please allow reasonable time for review and remediation before any public disclosure.

## Current Security Focus

The current Notepad-X security model focuses on:

- atomic file replacement for normal saves
- atomic writes for session, recovery, editor identity, note sidecars, and JSON note exports
- validation and sanitization of local support files before trusting them
- safer handling of malformed, oversized, unusual-encoding, and binary-like files
- conflict detection before overwriting files that changed on disk
- text-only file handling with no automatic execution of scripts or markup
- non-destructive recovery that restores content into tabs instead of overwriting user files
- user-visible permission and filesystem errors
- stale editor cleanup for shared note tracking

Notepad-X is a local desktop text editor and is not intended to provide anti-tamper, DRM, sandboxing, or high-assurance runtime integrity protection.
