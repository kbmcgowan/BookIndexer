For an overview of how an Index works in Word, see:

     https://support.microsoft.com/en-us/word/create-and-update-an-index


here's an interesting previous solution to the first part: https://github.com/karzmei/Book-index

If NLTK's copy of the Brown corpus has not been downloaded, install it with `python3 -m nltk.downloader brown`.

To generate a list of candidate terms:

`python3 ./suggest_index_terms.py`

The generator uses the NLTK Brown corpus as background documents for IDF, considers one-, two-, and three-word phrases, and adds up to 10 extra trigram candidates per chapter beyond the main TF-IDF list.

To review candidates in a desktop interface:

`python3 ./review_index_candidates.py`

The app opens `index_candidates.csv` by default. You can also pass another CSV, such as `index_candidates_with_ner.csv`. Select a row to see its extracted DOCX text with the term highlighted, check **Keep** for entries to retain, and save your choices.

The app requires PySide6 and python-docx:

`python3 -m pip install PySide6 python-docx`

The preview shows extracted text, not Word's page layout or pagination.

After reviewing, open the book in Word and run the VBA script:

     MarkReviewedIndexTerms.bas

It creates XE fields only for rows whose `Keep` value is `True`, in the Word document named on that row.
