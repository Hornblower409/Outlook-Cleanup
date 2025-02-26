# Outlook-Cleanup
Cleanup formatting of an Outlook Response

This is a Developer ONLY release. Not for general usage.

## Install:
- Import the two .BAS files.
- Tools -> References. [/] Microsoft Word NN.M Object Library

## Usage:

Cleanup was designed to run automatically on a Reply/Send event. But this version does not have any of the event hooks, so you have to create a response and then trigger it manually.
- Create a response (reply) in the Active Inspector.
- Macros -> Cleanup_Response_Lnk
