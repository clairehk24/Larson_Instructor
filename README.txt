LARSON INSTRUCTOR VERSION
=========================

START
Open index.html in a current desktop browser. The instructor guide does not require an internet connection or web server.

CURRENT MANUSCRIPT COVERAGE
- The instructor Introduction is built from assets/Manuscripts/L1715_Introduction.docx.
- Instructor simulation manuscripts currently available: 1, 3, 4, 5, 6, 7, 8, 9, and 10.
- Simulation 9 is partial and is presented only to the extent supplied in its manuscript.
- Supporting instructor resources are available for simulations 2, 16, 17, 21, and 22.
- Simulations without a supplied instructor manuscript display a status notice instead of student-version copy.
- Simulation Checklist, Debriefing Methods, and Simulation Finder appear after the simulations in the navigation.

REBUILD FROM MANUSCRIPTS
Run:

    python scripts/build_instructor_version.py

The build reads the L1715 Word manuscripts, writes instructor HTML pages, creates the marked instructor Word downloads, and creates an All Instructor Downloads ZIP for each ready simulation. Production directions such as \qqID blocks and production tags such as <a> are omitted from reader-facing output; manuscript wording is not rewritten.

To rebuild a single simulation while developing, append its number. For example:

    python scripts/build_instructor_version.py 6

VERIFY MANUSCRIPT WORDING
Run:

    python scripts/validate_manuscript_text.py

The validation checks reader-facing HTML and generated Word content against the source manuscripts. It fails when wording, punctuation, capitalization, or ordering differs.

ADDING A MANUSCRIPT
1. Copy the approved L1715 Word manuscript into assets/Manuscripts using the L1715_SimNN naming pattern.
2. Add any supplied figure PNG files to assets/images. The figure filename must match the manuscript's \qqINSERT source stem.
3. Run the build and validation commands above.
4. Confirm the simulation title in data/navigation.js matches the approved manuscript.

NOTES
- Do not hand-edit generated manuscript wording in pages/simulation-N.html or generated Word downloads. Update the source manuscript and rebuild.
- Completion is stored in the instructor's browser only. LMS reporting requires a SCORM/xAPI wrapper or LMS-specific integration.
- No external libraries or packaged font files are used by the browser interface.
