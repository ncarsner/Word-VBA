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
    a. Column A = KEY
    b. Column B = VALUE
2. Save/export to `TCA_dictionary.tsv`
    - TCA|T=16
    - TCA|T=16|C=3
    - TCA|T=16|C=3|P=3
    - TCA|T=16|C=3|P=3|S=101

3. Version control (date or Git)
4. Import into Word via macro

**Note:** Never edit the `.tsv` directly unless necessary.

```
TCA|T=16	Courts

TCA|T=16|C=1	General Provisions
TCA|T=16|C=1|P=1	[RESERVED – PART NOT YET CODIFIED]

TCA|T=16|C=2	Court Administration
TCA|T=16|C=2|P=1	Administrative Office of the Courts
TCA|T=16|C=2|P=2	Judicial Conferences
TCA|T=16|C=2|P=3	[RESERVED – PART NOT YET IDENTIFIED]

TCA|T=16|C=3	Supreme Court
TCA|T=16|C=3|P=1	Organization and Powers
TCA|T=16|C=3|P=2	Procedural Authority
TCA|T=16|C=3|P=3	Terms
TCA|T=16|C=3|P=4	[RESERVED – FUTURE LEGISLATION]

TCA|T=16|C=4	Court of Appeals
TCA|T=16|C=4|P=1	Composition and Jurisdiction
TCA|T=16|C=4|P=2	[RESERVED – PART NUMBER UNUSED]

TCA|T=16|C=5	Criminal Court
TCA|T=16|C=5|P=1	Jurisdiction
TCA|T=16|C=5|P=2	Judges and Terms
```
