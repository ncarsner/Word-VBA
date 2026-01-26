# Word-VBA
Original or adapted VBA scripts for Microsoft Word.


TCA Dictionary checklist
- Before using the file:
[ ] Every non-blank line contains exactly one TAB
[ ] No trailing spaces in keys
[ ] No duplicate keys
[ ] Placeholders are bracketed ([...])
[ ] Encoding is UTF-8 (Excel default is fine)

- Recommended workflow:
1. Maintain TCA_dictionary.xlsx as the authoritative file
2. Save/export to TCA_dictionary.tsv
3. Version control (date or Git)
4. Import into Word via macro

- Never edit the `.tsv` directly unless necessary.
