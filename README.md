pBMSC
=====
pBMSC is a modified version of uBMSC (which is a modified version of iBMSC) with a primary focus on quality of life functionalities such as keyboard shortcuts.
See README.md.iBMSC for the original iBMSC README file and REAME.md.uBMSC for the last uBMSC README file.

* Special thanks to Shimi9999#7602 for the Japanese translation with the new features.

### Discretion
* Themes from Bluvel's uBMSC fork may not be compatible, please use themes from Zadoru's uBMSC fork.
* Having .NET Framework 4.8 installed is recommended.
* Most of the new features are available in English only (with limited support for Chinese and Japanese).

# Changes
Listed in the order added.

## Bugfixes
* Added keybindings for DP and PMS by default. Now also added support for rebinding keys. See **Keyboard Shortcuts** for more information.
* Fixed the search function such that notes on lane A8 and D8 are now searchable.
* Fixed the mirror function such that notes between A1 and D8 are reflected locally. Supports PMS as well.
* Fixed the Statistic Label not including notes between D1-D8. Statistic window still not fixed.
* Fixed the total note count on the toolbar.
* Reorganized the sidebar so you can tab between textboxes properly.
* Prevented notes in Expansion Code from being loaded.
* Rare bug that occurs when the mouse is highlighting a long note while toggling between NT and BMSE simultaneously.
* Smoothened Selection Box animation.

## Functionalities
* Added Random, R-Random, S-Random and H-Random. Supports PMS as well.
* The application now saves the option "Disable Vertical Moves".
* Changed the temporary bms file extension from .bms to .bmsc.
* Added advanced statistics (Ctrl+Shift+T).
* Removed restriction for drag and dropping files, as well as opening files. Mainly for opening bms template files, not tested thoroughly.
* Added note search function (goto measure except it's goto note). One note per VPosition only.
* Added sort function. Selected notes are sorted based on their VPosition and Value.
* Added mBMplay as a default player.
* Added basic support for #RANDOM. Supports expansion field and main data field only. Not tested thoroughly, nested #RANDOM in "Modify Section" only. Accessible via the "Select Section" button in Expansion Code or via Ctrl+Alt+R.
* Added color overriding options where you can specify a range of notes to appear a specific color (or a color gradient), such as red notes for drums and green notes for the piano. Accessible via the Options tab or via Shift+F12.
* Added comment notes. Comment notes will be saved as #ECMD and #ECOM within the same bms/pms file. Not tested thoroughly.
* The window will now follow notes being moved by arrow keys.
* When creating LNs in NT mode, the VPosition will snap to the highlighted note if any.
* Added "Show Waveform on Notes".
* Added "Define Measure" in Time Selection Tool. Allows defining a custom measure, add a measure line or remove measure lines. Also accessible via the Insert key while Time Selection Tool is selected.
* Added "Insert/Remove Space" in Time Selection Tool. Allows inserting space into the measure, either measure only, notes only or both. Also accessible via Ctrl+Insert, Shift+Insert and Ctrl+Shift+Insert respectively.
* Added ability to copy and paste measure lengths over multiple measures.
* Added ability to duplicate #WAV's.
* Added technical error check (Jack notes, impossible chords for PMS, and impossible scratches for DP). Not a toggle because running it live will severely reduce performance. Accessible via the Options tab or via Ctrl+Alt+E.
* Added better support for all audio files (supported by CSCore + ogg).
* Added support for rebinding keys.
* Added #TOTAL Options window to set a recommended #TOTAL value based on the selected calculation formula.
* Added lines to scroll along the panel while previewing a keysound and to indicate the keysound's end point.
* Added support to move notes to their template position.
* Added ability to convert notes to Long notes in a series by VPosition or by column.
* Added ability to convert between LNObj's and NT LN's.
* Added ability to copy and paste patterns to highlighted notes. Accessible via Ctrl+Shift+V.
* Added BMS file tabs. See below for details.
* Added #BMP list.
* Added "brush" tool to BMP notes. Select a range of BMP files in the #BMP list then create a long note in NTInput mode, and the BMP notes will be added correspondingly.

## Keyboard shortcuts
* While rebinding keys is now supported, the default keybinding has also been changed to allow note placement between D1 and D8 or only on PMS lanes in PMS mode:
  * Numpad keys are assigned to 2P lanes when making DP charts.
  * QWERTYUI keys are also assigned to 2P lanes when making DP charts.
  * 1 to 7 are now assigned to A2 to A8, and 8 is now assigned to A1.
  * Ctrl+1 to Ctrl+8 are now assigned to D1-D8.
  * For PMS: Number keys 1-9 assign the notes to PMS lanes when a PMS theme is used.
* Added Save As keyboard shortcut (Ctrl+Alt+S)
* Added recent bms keyboard shortcuts (Alt+1 to Alt+5)
* Added shortcuts for toggling lanes:
  * Alt+B - BPM lane
  * Alt+S - Stop lane
  * Alt+R - Scroll lane
  * Alt+G - BGA/Layer/Poor
* Added shortcuts for the panel splitter (Alt+Left and Alt+Right).
* Added Options shortcut:
  * F9 - Player Options
  * F10 - General Options
  * F12 - Visual Options
* Added advanced statistics (Ctrl+Shift+T).
* Added keyboard shortcuts for previewing and replacing keysounds in the Sounds List (Spacebar to preview, enter to replace).
* Added "Select Section" (Ctrl+Alt+R by default).
* Added color overriding options (Shift+F12).
* Revised shortcuts for grid division values:
  * Period now increases the division value, and comma now decreases the division value.
  * Hold Ctrl for incremental value change
  * Hold Shift for x3 multiplier or divisor
  * Hold Alt to change the sub value
  * For example, Shift+Alt+Comma = divide the sub value by 3 (rounded)
* Added shortcut for scrolling through #WAV selection (Shift+Scroll).
* Added shortcut for defining measures with Time Selection Tool (Insert).
* Added shortcut for inserting/removing space with Time Selection Tool (Ctrl+Insert for measure only, Shift+Insert for notes only, and Ctrl+Shift+Insert for both)
* Added shortcut for techincal error check (Ctrl+Alt+E by default)
* Added shortcut for keybinding options (Shift+F10)
* Added shortcut for an experimental feature (Preview highlighted notes - Shift+F4)

## New dialog/option boxes

### Advanced Statistics (Ctrl+Shift+T)
* Displays note statistics over individual lanes from A1 to D8.
* Displays note usages, including #WAV assigned notes with 0 note count and #WAV unassigned notes with non-zero note count.

### Select Expansion Code (Ctrl+Alt+R)
* Display ghost notes: Select an #if section to display its notes in ghost form.
* Display ghost notes (All): Display all notes in ghost form.
* Modify notes: Select an #if section and jump straight to modifying them.
* Modify section: Open an instance of pBMSC, allowing you to modify the #if section individually.
* Remove ghost notes: Remove all ghost notes.

### Visual Override Options (Shift+F12)
* Creates a list of note ranges with a specified color and replaces specified notes' display color.
* Allows a range of notes to have a gradient of colors, or a single color.
* Settings can be saved to either Editor, Chart, or Song.

### Keybinding Options (Shift+F10)
* Allows for rebinding keys for a lot of functions, such as note lane assignment.
* Some other functions are also rebindable, such as Snap to Grid, Undo, and Redo.

### #TOTAL Options
* Allows you to set a recommended #TOTAL value based on either IIDX supposition 1, IIDX supposition 2, or a multiplier.
* Allows you to auto-set the #TOTAL value for the chart.

### BMS File Tabs
* Allows multiple files to be opened simultaneously.
* Middle-clicking on the file tab will close it.
* Saves a list of the opened files on exiting application and reopens them on relaunching application.

###### Unavailable features
* Autosaving all files. Currently it behaves the same as previously, which is autosaving the active BMS file.
* Integration with Modify Section in Select Expansion Code.
* Moving file tabs. Not sure if this will be doable with the current arrangement.
* Saving all files on exiting application. Unsaved changes will be gone upon closing.

# Planned To-Do List

### #RANDOM Editor
* Multiple #IF sections in one file (notes only).
* Allow ghost notes to be editable.
* Integration of Modify Section with FileTab.

### Misc
* Add support for saving new functions to iBMSC.
* Optimization?
* Unavailable BMS File Tab features listed above.