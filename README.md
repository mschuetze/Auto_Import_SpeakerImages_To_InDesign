# Auto_Import_SpeakerImages_To_InDesign
Script for Adobe InDesign that will import/place an image for each speaker from our server in to the .indd file. It does this based on a paragraph format that is reserved for speaker names only.

## What it does
- Check if document is already open – if not, throw error and stop script.
- Check if paragraph style **"speakers"** exists in the document – if not, throw error and stop script.
- Loop through each page and find paragraph with style **"speakers"**.
- Copy content of found paragraph to variable.
- Trim content of variable to only include the speaker name – everything after the opening bracket is removed.
