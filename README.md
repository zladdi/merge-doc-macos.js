# merge-doc-macos.js

A port of TortoiseSVN/TortoiseGit merge-doc.js to Open Scripting Architecture (OSA) on macOS.
For the source script, see https://github.com/TortoiseGit/TortoiseGit/blob/master/contrib/diff-scripts/merge-doc.js.

The code is distributed under the GNU General Public License. 

## Prerequisites

* Requires OS X Yosemite or later.
* Microsoft Word needs to be installed.

## Usage 

`merge-doc-macos.js <absolute-path-to-merged.doc> <absolute-path-to-theirs.doc> <absolute-path-to-mine.doc> <absolute-path-to-base.doc>`

## Known issues

* Does presently not support version 16 of Microsoft Office 2016 for Mac.
* Relative paths to documents are not supported.
* Newer versions of Microsoft Office apps are sandboxed and do not allow modifying the document in-memory if the app does not have write access to the underlying file. Thus, `merge-doc-macos.js` temporarily saves a copy of comparison results to a folder where the app has write access to (`~/Library/Group Containers/UBF8T346G9.Office`). Normally, `merge-doc-macos.js` deletes the documents containing the comparison results automatically as soon as possible. However, it might happen that the program exits unexpectedly without having removed the documents. So, you might want to routinely check for stale documents in that folder.
* Newer versions of Microsoft Office apps are sandboxed. This leads to the annoying
"Grant File Access" dialog to pop up for each of the documents to be compared in cases
where Word does not have permission to access the respective file already.

## Future work

Create a formula for [`brew`](https://github.com/Homebrew).
