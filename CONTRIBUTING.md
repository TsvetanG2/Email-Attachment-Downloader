# Contributing

Thanks for your interest in contributing to Email Attachment Downloader! This
guide explains how to get set up and submit changes.

## Getting Started

1. Fork the repository and clone your fork:
   ```bash
   git clone https://github.com/<your-username>/Email-Attachment-Downloader.git
   cd Email-Attachment-Downloader
   ```
2. (Recommended) Create and activate a virtual environment:
   ```bash
   python -m venv .venv
   source .venv/bin/activate      # Windows: .venv\Scripts\activate
   ```
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
4. Run the app:
   ```bash
   python main.py
   ```

## Project Structure

```
main.py                  # Application entry point
src/
  email_client.py        # IMAP connection & email fetching
  downloader.py          # Attachment extraction & saving
  renamer.py             # File renaming logic
  date_picker.py         # Calendar date picker widget
  preview_window.py      # Email preview/selection window
  gui.py                 # Main GUI application
```

## How to Contribute

1. Create a branch:
   ```bash
   git checkout -b feature/your-feature
   ```
2. Make your changes, keeping them focused.
3. Test manually against a real (or test) mailbox — connect, search, preview,
   and download — since the GUI cannot be fully tested in CI.
4. Add an entry to `CHANGELOG.md` under `[Unreleased]`.
5. Push and open a Pull Request describing what you changed and why.

## Code Style

- Follow [PEP 8](https://peps.python.org/pep-0008/).
- Run a linter before submitting:
  ```bash
  pip install flake8
  flake8 .
  ```

## Security-Sensitive Code

This app handles email credentials. When contributing, please uphold these
rules:

- **Never log or print credentials** (email address is fine; passwords are not).
- **Never write credentials to disk** — they must stay in memory only.
- **Keep all mail connections over SSL/TLS.**
- **Do not add features that send, modify, or delete email** — the app is
  intentionally read-only.
- **Never commit downloaded attachments, build artifacts, or `.env` files**
  (see `.gitignore`).

If you find a security issue, do **not** open a public issue — follow
[SECURITY.md](SECURITY.md).

## Good First Contributions

- Additional email providers (Yahoo, iCloud, custom IMAP)
- More rename patterns
- Better error handling for flaky connections / rate limits
- Accessibility and UI polish
- Packaging for macOS / Linux

## Code of Conduct

By participating, you agree to abide by the [Code of Conduct](CODE_OF_CONDUCT.md).

## License

By contributing, you agree that your contributions will be licensed under the
[MIT License](LICENSE).
