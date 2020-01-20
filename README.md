# VimEverywhere

Use some of vim's key binding almost everywhere.

### Requirements:

PyHook3, PyWin32, PyperClip, PySimpleGUIQt

### Usage:

Directly run 'vimeverywhere.pyw'. The system tray icon will show you the edit state. And you can right click it to temporarily release the key binding or exit the script. Three edit modes support: normal mode, insert mode and visual mode.

### Support bindings:

* Normal mode:
  * \d*[wW]
  * \d*[hjkl]
  * [fFtT].
  * ;
  * [\\^\\$]
  * [iIaA]
  * r.
  * \d*[xs]
  * \d*((dd)|(yy)|(cc))
  * [DC]
  * [oO]
  * [dyc]\d*[wW]
  * \[dyc][\\^\\$]
  * \[dyc][fFtT].
  * \\.
  * :w
  * u
  * v

* Insert mode:
  * Escape key

* Visual mode:
  * \d*[hjkl]
  * [xdycs]
  * [\\^\\$]
  * Escape key

### Notice:

1. The copy(yank) and cut(delete) commands directly use system clipboard.
2. Normal mode and visual mode will block all the keys except composite keys and some control keys, so you can use keys like control+C, control+V to copy and paste.
3. This script uses the first character next to the cursor as "the character under the cursor" in vim.