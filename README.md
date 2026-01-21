# Microsoft Skill

A Claude Code plugin for Microsoft Graph API integration, focused on Outlook/Hotmail email access.

## Installation

Add this plugin to Claude Code:

```bash
claude mcp add-plugin /path/to/microsoft-skill
```

Or add to your Claude Code configuration.

## Features

- **OAuth 2.0 Authentication** with PKCE
- **Email Operations**: List messages, download as `.eml`
- **Secure Storage**: 1Password or local file credentials
- **Per-project Tokens**: Different Microsoft accounts per project

## Quick Start

Once installed, the skill triggers when you ask about emails:

- "Check my outlook inbox"
- "Read my emails"
- "Download recent messages from hotmail"
- "List my microsoft mail"

### First-time Setup

1. **Configure credentials**:
   ```bash
   pnpm tsx skills/microsoft-outlook/scripts/microsoft.ts setup
   ```

2. **Authenticate**:
   ```bash
   pnpm tsx skills/microsoft-outlook/scripts/microsoft.ts auth
   ```

3. **Verify**:
   ```bash
   pnpm tsx skills/microsoft-outlook/scripts/microsoft.ts me
   ```

## Plugin Structure

```
microsoft-skill/
├── .claude-plugin/
│   └── plugin.json          # Plugin manifest
├── skills/
│   └── microsoft-outlook/
│       ├── SKILL.md         # Skill instructions
│       ├── scripts/         # CLI implementation
│       │   ├── microsoft.ts
│       │   └── lib/
│       └── references/      # Detailed documentation
│           ├── setup-guide.md
│           └── api-reference.md
├── package.json
└── README.md
```

## Commands

| Command | Description |
|---------|-------------|
| `setup` | Configure API credentials |
| `auth` | Run OAuth flow |
| `check` | Verify auth status |
| `me` | Get user profile |
| `messages` | List recent messages |
| `download <id>` | Download message as .eml |
| `download-all` | Download recent messages |

## Credential Setup

Register an app in [Microsoft Entra admin center](https://entra.microsoft.com/) or the [Azure Portal](https://portal.azure.com/):

1. **Sign in**: Use your Microsoft account. If you see an error about not being in a directory, follow the prompt to "Sign up for Azure" (free) to create a default directory.
2. **App registrations** > **New registration**
3. **Name**: `microsoft-skill-cli`
4. **Supported account types**: Select **"Accounts in any organizational directory (Any Microsoft Entra ID tenant - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)"**.
   * *Note: This is required if you want to use a personal @outlook.com or @hotmail.com account.*
5. **Redirect URI**: Select **Web** and enter `http://localhost:3000/callback`.
6. **Register**: Click the button.
7. **Client ID**: Copy the **Application (client) ID** from the Overview page.
8. **Client Secret**:
   * Go to **Certificates & secrets** > **New client secret**.
   * Add a description and click **Add**.
   * Copy the **Value** (not the Secret ID) immediately.
9. **API permissions**:
   * Go to **API permissions** > **Add a permission**.
   * Select **Microsoft Graph** > **Delegated permissions**.
   * Add `Mail.Read` and `User.Read`.
   * (Optional) Click **Grant admin consent** if you are the tenant admin.

### Storage Options

**1Password** (recommended):
The script looks for Secure Notes in your **Personal** vault:
- `op://Personal/Microsoft Client ID/notesPlain`
- `op://Personal/Microsoft Client Secret/notesPlain`

**Local file**: `~/.config/microsoft-skill/credentials.json`
Run `pnpm tsx scripts/microsoft.ts setup` and choose "Enter manually".

## Token Storage

- **Per-project**: `.claude/microsoft-skill.local.json`
- **Global**: `~/.config/microsoft-skill/tokens.json`

Add `.claude/*.local.*` to `.gitignore`.

## License

MIT
