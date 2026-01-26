# Word-VBA

A collection of original and adapted VBA macros for Microsoft Word to enhance document automation and productivity.

## Macros

### TCA_input

A macro for managing and importing dictionary entries from a TSV file into Word documents.

#### TCA Dictionary File Checklist

Before using the dictionary file:
- [ ] Every non-blank line contains exactly one TAB
- [ ] No trailing spaces in keys
- [ ] No duplicate keys
- [ ] Placeholders are bracketed ([...])
- [ ] Encoding is UTF-8 (Excel default is fine)

#### Recommended Workflow

1. Maintain `TCA_dictionary.xlsx` as the authoritative file
2. Save/export to `TCA_dictionary.tsv`
3. Version control (date or Git)
4. Import into Word via macro

**Note:** Never edit the `.tsv` directly unless necessary.
