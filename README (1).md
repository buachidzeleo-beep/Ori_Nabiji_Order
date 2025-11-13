# Config Directory

This folder contains configuration files for the **Two-Step Order Cleaner**.

## `client_removal_template.xlsx`

- Sheet: `clients_to_clear`
- Columns:
  - `shop_code` (required) — numeric shop code from the address row token `#ID#`.
  - `shop_nickname_optional` (optional) — shop nickname as in the first row of the order file.
  - `notes_optional` (optional) — free text, ignored by logic.

The app will load this template by default and clear orders only for rows/columns
matching these shop codes or nicknames (except for the protected supplier).

You can always override this default in the UI by uploading another template.
