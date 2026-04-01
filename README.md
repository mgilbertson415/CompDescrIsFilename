# CompDescrIsFilename
## SW macro to initialize part/assembly descriptions to their (extension-removed) filename

Algorithm:
- Respond to "dry run or not" prompt
- Browse to folder
- Code selects every part and assembly
- For each file, "Description" property is set to filename without extension
    - Two cases covered: add property if it is missing; update it in place if it exists
