# ReRead
Read raw csv-files, trim and format into Labware LIMS readable files.

One of our aliquoting robots has been excluding empty tubes from resultfile, causing a discrepancy in our LIMS. LIMS needs to be aware that plates have tubes in those positions, even if the tube itself is empty.
To rectify this we'll use the raw csv files to "reread" and set a non empty flag on positions containing tubes on the plates.
