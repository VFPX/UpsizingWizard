# How to contribute to Upsizing Wizard

## Report a bug
- Please check [issues](https://github.com/VFPX/UpsizingWizard/issues) if the bug has already been reported.
- If you're unable to find an open issue addressing the problem, open a new one. Be sure to include a title and clear description, as much relevant information as possible, and a code sample or an executable test case demonstrating the expected behavior that is not occurring.

## Fix a bug or add an enhancement
- Fork the project: see this [guide](https://www.dataschool.io/how-to-contribute-on-github/) for setting up and using a fork.
- Make whatever changes are necessary. **Be sure to build the project using VFP 9 SP 2 rather than VFP Advanced since the APP format is different.**
- Use [FoxBin2Prg](https://github.com/fdbozzo/foxbin2prg) to create text files for all VFP binary files (SCX, VCX, DBF, etc.)
- Update the Releases section of README.md and describe the changes.
- Right-click CreateThorUpdate.ps1 in the ThorUpdater folder and choose Run with PowerShell to update the file the Thor Check for Updates function needs.
- Commit the changes.
- Push to your fork.
- Create a pull request; ensure the description clearly describes the problem and solution or the enhancement.

----
Last changed: 2022-04-09