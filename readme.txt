.DESCRIPTION
    - Downloads each subfolder (excluding ZZZ_COMPRESSION_HISTORY) under a target folder from SPO to local temp.
    - If a ZIP with matching 6-digit reference exists at root level, it downloads and updates it; otherwise creates new ZIP.
    - Uploads ZIP back to the same location.
    - Writes a per-root folder compression log to ZZZ_COMPRESSION_HISTORY.
    - Deletes the original (now compressed) folder.
    - Handles long local paths by truncating filenames with a stable short SHA1 suffix.


## License

This project is licensed under the GNU General Public License v3.0 or later.

See the [LICENSE](LICENSE) file for details.

SPDX-License-Identifier: GPL-3.0-or-later
